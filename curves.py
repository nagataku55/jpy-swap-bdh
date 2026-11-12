"""JPY TONA OISスワップのスポットレートからカーブを構築し、割引率・フォワードレート・
フォワードスワップレート・PV・バケット感応度を計算するモジュール。

excel-fastapiプロジェクトのcurves.pyを移植し、forward_swap_rate()を新規追加した。
それ以外の既存関数は内容の変更なし。
"""
import hashlib
import json
from collections import OrderedDict
from typing import Optional

import QuantLib as ql

from qlutils import (
    approx_months,
    get_business_day_convention,
    get_calendar,
    get_compounding,
    get_day_counter,
    get_end_of_month,
    get_overnight_index,
    get_swap_position,
    parse_period,
)

_ON_ALIASES = {"ON", "O/N"}

_MAX_CACHE_SIZE = 200

# ブートストラップの収束判定(割引率の反復ごとの変化幅)。QuantLibの既定は1e-12だが、
# Kruger法のようなグローバル補間では反復のたびに全ノードが少しずつ動くため、
# 実データで「last improvement 5e-12〜9e-12」のように既定値のわずか手前で
# 99回の反復上限に達して失敗する日が多数あった(2025-11〜2026-11で約30日)。
# 1e-10は割引率の誤差として0.0001bp未満で、レートへの影響は実質ゼロ。
BOOTSTRAP_ACCURACY = 1e-10
# curve_id -> {"curve": ql.YieldTermStructure, "params": 構築時のpayload(instruments含む)}
# paramsを保持するのは、バケット感応度計算で個々のテナーだけレートをbumpして再構築するため。
_curve_cache: "OrderedDict[str, dict]" = OrderedDict()


def _to_ql_date(iso: str) -> ql.Date:
    y, m, d = (int(x) for x in iso.split("-"))
    return ql.Date(d, m, y)


def _iso(d: ql.Date) -> str:
    return f"{d.year():04d}-{d.month():02d}-{d.dayOfMonth():02d}"


_PERIOD_UNIT_LABELS = {
    ql.Days: "D",
    ql.Weeks: "W",
    ql.Months: "M",
    ql.Years: "Y",
}


def _period_label(period: ql.Period) -> str:
    return f"{period.length()}{_PERIOD_UNIT_LABELS[period.units()]}"


def _multiply_period(period: ql.Period, n: int) -> ql.Period:
    return ql.Period(period.length() * n, period.units())


def _curve_id(payload: dict) -> str:
    canonical = json.dumps(payload, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(canonical.encode()).hexdigest()[:16]


def build_curve(
    valuation_date_iso: str,
    instruments: list,
    calendar_name: str = "TKB",
    day_count_code: str = "ACT365F",
    overnight_index_code: str = "TONAR",
    convention_code: str = "MF",
    settlement_days: int = 2,
    eom_code: str = "STD",
    payment_lag_days: int = 2,
) -> str:
    """OISスワップ・スポットレート(ON + 翌日物指数の複利スワップ)からカーブをブートストラップし、curve_idを返す。

    ONはオーバーナイト参照金利そのものとしてDepositRateHelper(決済ラグ0日)、
    それ以外(1W〜)はOISRateHelperで複利ブートストラップする(支払いは満期日からpayment_lag_days営業日後、市場慣行)。
    同一の入力データからは同一のcurve_id(内容ハッシュ)になるため、
    Excelの再計算のたびにカーブが積み上がることはない。
    """
    payload = {
        "valuation_date": valuation_date_iso,
        "instruments": sorted(
            ({"tenor": i["tenor"].strip().upper(), "rate": float(i["rate"])} for i in instruments),
            key=lambda x: x["tenor"],
        ),
        "calendar": calendar_name.strip().upper(),
        "day_count": day_count_code.strip().upper(),
        "overnight_index": overnight_index_code.strip().upper(),
        "convention": convention_code.strip().upper(),
        "settlement_days": settlement_days,
        "eom": eom_code.strip().upper(),
        "payment_lag_days": payment_lag_days,
    }
    curve_id = _curve_id(payload)

    if curve_id in _curve_cache:
        _curve_cache.move_to_end(curve_id)
        return curve_id

    calendar = get_calendar(calendar_name)
    day_count = get_day_counter(day_count_code)
    convention = get_business_day_convention(convention_code)
    is_eom = get_end_of_month(eom_code)
    overnight_index = get_overnight_index(overnight_index_code)

    ql_valuation_date = _to_ql_date(valuation_date_iso)
    ql.Settings.instance().evaluationDate = ql_valuation_date

    helpers = []
    for inst in payload["instruments"]:
        tenor_period = parse_period(inst["tenor"])
        quote = ql.QuoteHandle(ql.SimpleQuote(inst["rate"]))

        is_overnight = inst["tenor"] in _ON_ALIASES or (
            tenor_period.units() == ql.Days and tenor_period.length() <= 1
        )
        if is_overnight:
            # オーバーナイト参照金利そのもの: 決済ラグなし。
            # 文字列が"ON"/"O/N"の場合に加えて、"1D"のように1日物として表記される
            # 実データ向けに、パース後の期間が1日以下かどうかでも判定する(文字列の
            # 表記ゆれに頼らない)。これを素通りさせてOISRateHelper(settlement_days=2日)
            # 扱いにすると、「2日後スタート・1日後満期」という極端に短い縮退した
            # インストゥルメントになり、カーブの短期側でKruger法の収束が不安定になる。
            helper = ql.DepositRateHelper(quote, tenor_period, 0, calendar, convention, is_eom, day_count)
        else:
            payment_frequency = ql.Once if approx_months(tenor_period) <= 12 else ql.Annual
            # pillar(カーブのノード)は支払日ではなく満期日(=複利計算期間の終了日)に置く。
            # OISのパーレートは「スポット日〜満期日」の割引率の比だけで決まり、支払日
            # (満期日+payment_lag営業日)の割引率には依存しない(両レッグとも同じ日に
            # 払うので打ち消し合う)。既定のLastRelevantDate(=支払日)だと、年末年始・
            # GW・シルバーウィークのように祝日が続く時期に、短期テナーの支払日が次の
            # テナーの満期日付近まで後ろにずれ、「そのノードを動かしても自分のレートが
            # 変わらない」状態になってルート探索が失敗していた
            # (RuntimeError: root not bracketed: f[...] -> [x, x])。
            helper = ql.OISRateHelper(
                settlement_days,
                tenor_period,
                quote,
                overnight_index,
                ql.YieldTermStructureHandle(),
                False,
                payment_lag_days,
                convention,
                payment_frequency,
                calendar,
                ql.Period(0, ql.Days),  # forwardStart
                0.0,                    # overnightSpread
                ql.Pillar.MaturityDate,
            )
        helpers.append(helper)

    # 既定のLogCubic(natural spline, monotonic制約なし)は実データの構成(短期は密・
    # 長期の一部が疎)でブートストラップが収束しないことがあった。Kruger法(単調性を
    # 保つ3次スプライン)なら同じデータでも収束し、フォワードレートも滑らかになる。
    log_cubic = ql.LogCubic(ql.CubicInterpolation.Kruger, True)
    bootstrap = ql.IterativeBootstrap(BOOTSTRAP_ACCURACY)
    curve = ql.PiecewiseLogCubicDiscount(ql_valuation_date, helpers, day_count, [], [], log_cubic, bootstrap)
    curve.enableExtrapolation()

    if len(_curve_cache) >= _MAX_CACHE_SIZE:
        _curve_cache.popitem(last=False)
    _curve_cache[curve_id] = {"curve": curve, "params": payload}
    return curve_id


def _get_cache_entry(curve_id: str) -> dict:
    entry = _curve_cache.get(curve_id)
    if entry is None:
        raise ValueError(f"Unknown curve_id: {curve_id!r} (curve may have expired, rebuild it)")
    _curve_cache.move_to_end(curve_id)
    return entry


def get_curve(curve_id: str) -> ql.YieldTermStructure:
    return _get_cache_entry(curve_id)["curve"]


def discount_factor(curve_id: str, end_date_iso: str) -> float:
    """計算期間終了日(end_date)を受け取り、カーブ構築時と同じカレンダー・
    payment_lag_days規則で実際の支払日(PaymentDate)にシフトしてから割引係数を返す。

    割引は支払日に対して行うべきだが(キャッシュフローが実際に動く日)、
    Excel側からは各期間のend_dateを渡すだけで済むように、ラグ調整はここで吸収する。
    """
    entry = _get_cache_entry(curve_id)
    curve = entry["curve"]
    params = entry["params"]

    calendar = get_calendar(params["calendar"])
    convention = get_business_day_convention(params["convention"])

    end_date = _to_ql_date(end_date_iso)
    payment_date = calendar.advance(end_date, params["payment_lag_days"], ql.Days, convention)

    return curve.discount(payment_date)


def forward_rate(
    curve_id: str,
    start_iso: str,
    end_iso: str,
    day_count_code: str = "ACT365F",
    compounding_code: Optional[str] = None,
) -> float:
    """start_isoからend_isoまでの単一区間のポイントフォワードレートを返す。

    これは2点間の割引率だけから出す値であり、複数回払い(年次など)のスワップの
    アニュイティを考慮しない。forward_tenorが1年以内(一括払い)ならフォワード
    スワップレートと一致するが、1年超では一致しない。1年超のフォワード"スワップ"
    レートが欲しい場合はforward_swap_rate()を使うこと。

    compounding_codeを省略した場合、期間の長さに応じて自動選択する
    (カーブ構築時の支払い頻度と同じ規則: 1年以下は単利/満期一括、1年超は年次複利)。
    """
    curve = get_curve(curve_id)
    day_count = get_day_counter(day_count_code)
    start_date = _to_ql_date(start_iso)
    end_date = _to_ql_date(end_iso)

    if compounding_code is None:
        year_frac = day_count.yearFraction(start_date, end_date)
        compounding = ql.Simple if year_frac <= 1.0 else ql.Compounded
    else:
        compounding = get_compounding(compounding_code)

    ir = curve.forwardRate(start_date, end_date, day_count, compounding, ql.Annual)
    return ir.rate()


def forward_swap_rate(
    curve_id: str,
    start_iso: str,
    end_iso: str,
    forward_tenor: str,
    calendar_name: str = "TKB",
    day_count_code: str = "ACT365F",
    convention_code: str = "MF",
    eom_code: str = "STD",
    overnight_index_code: str = "TONAR",
    payment_lag_days: int = 2,
) -> float:
    """start_isoを開始日とする、forward_tenor分のフォワードスワップの
    パーレート(PVがゼロになる固定レート)をcurve_id上で返す。

    swap_pv()と同じスケジュール構築をしてOvernightIndexedSwapを作る。
    固定レグの支払頻度はbuild_curve()のOIS実装と同じ規則を適用する:
    forward_tenorが12ヶ月以下なら一括払い(スケジュールの区切りがstart/endの2点のみ)、
    12ヶ月超なら年次払い(Backward生成、終了日基準で1年ごとに区切る)。

    パーレートの求め方: QuantLibの swap.fairRate() (内部でNewton法のソルバーを使う)
    は、このOIS×LogCubicカーブの組み合わせだと収束に失敗するケースがあった
    (RuntimeError: convergence not reached)。OISスワップのPVは固定レートrに対して
    厳密に線形(PV(r) = 浮動レグPV - r×想定元本×アニュイティ)なので、ソルバーに頼らず
    r=0とr=0.01の2点でNPVを計算し、直線の式から直接(ゼロ割を気にせず)パーレートを
    解く方が速く確実。
    """
    entry = _get_cache_entry(curve_id)
    curve = entry["curve"]
    handle = ql.YieldTermStructureHandle(curve)

    calendar = get_calendar(calendar_name)
    day_count = get_day_counter(day_count_code)
    convention = get_business_day_convention(convention_code)
    is_eom = get_end_of_month(eom_code)
    overnight_index = get_overnight_index(overnight_index_code, handle)

    start_date = _to_ql_date(start_iso)
    end_date = _to_ql_date(end_iso)

    forward_period = parse_period(forward_tenor)
    payment_tenor = forward_period if approx_months(forward_period) <= 12 else ql.Period(1, ql.Years)

    schedule = ql.Schedule(
        start_date, end_date, payment_tenor, calendar, convention, convention, ql.DateGeneration.Backward, is_eom
    )

    def _npv_at(fixed_rate: float) -> float:
        swap = ql.OvernightIndexedSwap(
            ql.OvernightIndexedSwap.Payer,
            1.0,
            schedule,
            fixed_rate,
            day_count,
            overnight_index,
            0.0,
            payment_lag_days,
            convention,
            calendar,
        )
        swap.setPricingEngine(ql.DiscountingSwapEngine(handle))
        return swap.NPV()

    npv_0 = _npv_at(0.0)
    probe_rate = 0.01
    npv_probe = _npv_at(probe_rate)

    # PV(r) = npv_0 - slope*r (slope = 想定元本×アニュイティ、r=0での傾き)。
    # PV(r)=0となるrを直線の式から直接求める。
    slope = (npv_0 - npv_probe) / probe_rate
    return npv_0 / slope


def instantaneous_forward_rate(curve_id: str, date_iso: str, day_count_code: str = "ACT365F") -> float:
    """瞬間フォワードレート(ゼロカーブの瞬間的な傾き -d/dT ln P(0,T))を連続複利で返す。

    start=endでforwardRateを呼ぶと、QuantLibが内部で微小区間(dt=0.0001年)の
    フォワードレートを計算し、瞬間フォワードの近似値を返す(LogCubicカーブなので滑らか)。
    """
    curve = get_curve(curve_id)
    day_count = get_day_counter(day_count_code)
    target_date = _to_ql_date(date_iso)
    ir = curve.forwardRate(target_date, target_date, day_count, ql.Continuous)
    return ir.rate()


def forward_grid(
    curve_id: str,
    forward_tenor: str = "1M",
    step_tenor: str = "1M",
    max_start_tenor: str = "2Y",
    day_count_code: str = "ACT365F",
) -> list:
    """スポット~max_start_tenorの範囲を、step_tenor刻みのスタート地点グリッドに区切り、
    各スタート地点からforward_tenor物のフォワードレートを計算する。

    例: forward_tenor="1M", step_tenor="1M", max_start_tenor="2Y"なら、
    Spot, 1M, 2M, ..., 23M をそれぞれスタートとする「xxx into 1M」フォワードの列になる
    (23M+1M=24M=2Yで打ち切り)。
    グリッドの起点(スタート=0)はカーブ構築時のスポット日(valuation_date +
    settlement_days営業日)。カレンダー・営業日調整・EOM規則もカーブ構築時のものを使う。
    """
    entry = _get_cache_entry(curve_id)
    params = entry["params"]

    calendar = get_calendar(params["calendar"])
    convention = get_business_day_convention(params["convention"])
    is_eom = get_end_of_month(params["eom"])
    ql_valuation = _to_ql_date(params["valuation_date"])
    ql_spot = calendar.advance(ql_valuation, params["settlement_days"], ql.Days)

    forward_period = parse_period(forward_tenor)
    step_period = parse_period(step_tenor)
    max_period = parse_period(max_start_tenor)

    step_months = approx_months(step_period)
    max_months = approx_months(max_period)
    n_steps = round(max_months / step_months)
    if n_steps < 1:
        raise ValueError(f"max_start_tenor {max_start_tenor!r} is shorter than step_tenor {step_tenor!r}")

    rows = []
    for i in range(n_steps):
        offset_period = _multiply_period(step_period, i)
        ql_start = ql_spot if i == 0 else calendar.advance(ql_spot, offset_period, convention, is_eom)
        ql_end = calendar.advance(ql_start, forward_period, convention, is_eom)

        start_iso = _iso(ql_start)
        end_iso = _iso(ql_end)
        rate = forward_rate(curve_id, start_iso, end_iso, day_count_code)

        rows.append(
            {
                "start_label": "SPOT" if i == 0 else _period_label(offset_period),
                "start_date": start_iso,
                "end_date": end_iso,
                "rate": rate,
            }
        )

    return rows


def swap_pv(
    curve_id: str,
    start_iso: str,
    end_iso: str,
    fixed_rate: float,
    notional: float,
    position_code: str = "PAY",
    payment_tenor: str = "1Y",
    calendar_name: str = "TKB",
    day_count_code: str = "ACT365F",
    convention_code: str = "MF",
    eom_code: str = "STD",
    overnight_index_code: str = "TONAR",
    payment_lag_days: int = 2,
) -> float:
    """OISスワップ(固定 vs TONA複利)のNPVを、curve_idのカーブで評価する。

    割引・フォワードプロジェクションとも同じカーブ(シングルカーブ)を使う。
    position_code: "PAY"(固定支払/変動受取) or "RECEIVE"(固定受取/変動支払)
    """
    curve = get_curve(curve_id)
    handle = ql.YieldTermStructureHandle(curve)

    calendar = get_calendar(calendar_name)
    day_count = get_day_counter(day_count_code)
    convention = get_business_day_convention(convention_code)
    is_eom = get_end_of_month(eom_code)
    position = get_swap_position(position_code)
    tenor_period = parse_period(payment_tenor)
    overnight_index = get_overnight_index(overnight_index_code, handle)

    start_date = _to_ql_date(start_iso)
    end_date = _to_ql_date(end_iso)

    schedule = ql.Schedule(
        start_date, end_date, tenor_period, calendar, convention, convention, ql.DateGeneration.Backward, is_eom
    )

    swap = ql.OvernightIndexedSwap(
        position, notional, schedule, fixed_rate, day_count, overnight_index, 0.0, payment_lag_days, convention, calendar
    )
    swap.setPricingEngine(ql.DiscountingSwapEngine(handle))
    return swap.NPV()


def bucket_sensitivity(curve_id: str, swap_kwargs: dict, bump: float = 0.0001) -> list:
    """カーブを構成する各テナーのレートをbump(既定1bp)だけ動かして再構築し、
    スワップPVの変化量(バケットDV01)をテナーごとに返す。
    """
    entry = _get_cache_entry(curve_id)
    params = entry["params"]

    base_pv = swap_pv(curve_id, **swap_kwargs)

    results = []
    for i, inst in enumerate(params["instruments"]):
        bumped_instruments = [dict(x) for x in params["instruments"]]
        bumped_instruments[i]["rate"] = bumped_instruments[i]["rate"] + bump

        bumped_curve_id = build_curve(
            valuation_date_iso=params["valuation_date"],
            instruments=bumped_instruments,
            calendar_name=params["calendar"],
            day_count_code=params["day_count"],
            overnight_index_code=params["overnight_index"],
            convention_code=params["convention"],
            settlement_days=params["settlement_days"],
            eom_code=params["eom"],
            payment_lag_days=params["payment_lag_days"],
        )
        bumped_pv = swap_pv(bumped_curve_id, **swap_kwargs)
        results.append({"tenor": inst["tenor"], "dv01": bumped_pv - base_pv})

    return results
