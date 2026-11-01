"""swap_history(スポットレート)から、フォワードスワップレートのヒストリカルデータを
作成・更新するモジュール。

設計のポイント:
  - 「今日だけ計算」と「過去分のバックフィル(初回構築・エラー後の欠損・新規スペック追加)」
    を区別しない。refresh_missing()は毎回「swap_historyにある日付のうち、
    forward_historyにまだ無いもの」を全部計算する1本のロジックにしている。
    通常運用では実質「前回実行からの差分(=だいたい当日1件)」だけになる。
  - カーブ構築(curves.build_curve)は日付ごとに1回だけ行い、その1つのカーブに対して
    アクティブな全スペックのフォワードスワップレートを計算する(逆に「スペックごとに
    日付をループ」すると同じ日のカーブをスペック数分re-buildしてしまい無駄)。
  - バックフィルの対象はmax_backfill_days日前までに限定する(既定365日)。
    それより古い欠損は自動では埋めない(意図しない全期間再計算のガード)。
  - 計算結果が前日比でWARN_THRESHOLD(既定50bp)を超えて動いた場合、取り込み自体は
    継続しつつ警告メッセージを集めて返す(呼び出し元がUpdateLogシートに書く想定)。
  - 計算対象のあったrunでは、監査用CSV(csv_audit/ForwardHistory_YYYY-MM-DD.csv)を
    追記する。Excelのシートを経由しない(forward_historyはそもそもシート上に無いデータ
    なので、サーバー側で直接CSVに残す)。
  - 1日分のカーブ構築/計算が失敗しても(実データ特有のカーブ形状でQuantLibの
    ブートストラップが収束しない等)、その日だけスキップして残りの日付の計算を
    続ける。1日の異常で全体が500エラーになり日次更新全体が止まるのを防ぐため。
    スキップした日は警告として返す。
"""
import csv
from datetime import date, timedelta
from itertools import groupby
from pathlib import Path

import history_store
from curves import build_curve, forward_swap_rate

WARN_THRESHOLD = 0.0050  # 50bp(前日比でこれを超えて動いたら警告)
AUDIT_DIR = Path(__file__).parent / "csv_audit"


def _write_audit_csv(rows: list[dict]) -> Path:
    """計算した行をcsv_audit/ForwardHistory_<実行日>.csvに追記する。
    同日に複数回実行された場合は上書き(その日の最新の計算結果を残す)。
    """
    AUDIT_DIR.mkdir(exist_ok=True)
    out_path = AUDIT_DIR / f"ForwardHistory_{date.today().isoformat()}.csv"

    fieldnames = ["as_of_date", "index_name", "ccy", "start_tenor", "forward_tenor", "rate"]
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for r in rows:
            writer.writerow({k: r[k] for k in fieldnames})

    return out_path


def refresh_missing(
    max_backfill_days: int = 365,
    calendar_name: str = "TKB",
    day_count_code: str = "ACT365F",
    convention_code: str = "MF",
    eom_code: str = "STD",
    overnight_index_code: str = "TONAR",
    settlement_days: int = 2,
    payment_lag_days: int = 2,
) -> dict:
    """アクティブなforward_specごとに、swap_historyにあるがforward_historyに無い
    日付をまとめて計算してUPSERTする。戻り値: {"computed": 件数, "warnings": [...]}。

    カレンダー・日数計算・コンベンションの各引数は、swap_history/curve構築時に使って
    いるものと同じ既定値(TKB/ACT365F/MF/STD/TONAR/2/2)。異なる規則で構築したカーブを
    使っている場合はここも合わせて変更すること。
    """
    specs = history_store.load_active_forward_specs()
    if not specs:
        return {"computed": 0, "warnings": [], "message": "アクティブなforward_specがありません"}

    since_iso = (date.today() - timedelta(days=max_backfill_days)).isoformat()

    computed = 0
    warnings: list[str] = []
    all_new_rows: list[dict] = []

    # この処理全体で接続を1回だけ開いて使い回す。何百回ものクエリを1回ずつ
    # connect()し直すと接続オーバーヘッドが積み重なって大きく遅くなるため
    # (計測: 600件の計算で約136秒→接続を使い回す形に直して1秒未満に短縮)。
    conn = history_store.connect()
    try:
        # 同じ(index_name, ccy)のスペックはswap_historyの対象日が共通なので、
        # グループ化してカーブ構築の対象日リストを1回だけ取得する。
        key = lambda s: (s["index_name"], s["ccy"])
        for (index_name, ccy), group_iter in groupby(sorted(specs, key=key), key=key):
            group_specs = list(group_iter)
            dates = history_store.list_swap_history_dates(conn, index_name, ccy, since_iso)

            for date_iso in dates:
                missing = history_store.missing_specs_for_date(conn, date_iso, index_name, ccy, group_specs)
                if not missing:
                    continue

                instruments = history_store.load_instruments_for_date(conn, date_iso, index_name, ccy)
                if not instruments:
                    continue

                try:
                    # カーブ構築はこの日付につき1回だけ。missing内の全スペックがこの1つの
                    # curve_idを再利用する(カーブ再構築の回数をスペック数分減らすため)。
                    # QuantLibのPiecewiseLogCubicDiscountは遅延評価のため、実際の
                    # ブートストラップ計算(収束しない場合に例外を投げる)はbuild_curve()
                    # の時点ではなく、下のforward_swap_rate()内で初めて割引率を読む時に
                    # 走る。そのため例外はbuild_curve()ではなくここで起きる想定。
                    curve_id = build_curve(
                        valuation_date_iso=date_iso,
                        instruments=instruments,
                        calendar_name=calendar_name,
                        day_count_code=day_count_code,
                        overnight_index_code=overnight_index_code,
                        convention_code=convention_code,
                        settlement_days=settlement_days,
                        eom_code=eom_code,
                        payment_lag_days=payment_lag_days,
                    )

                    new_rows_for_date = []
                    for spec in missing:
                        start_iso, end_iso = _resolve_period(
                            date_iso, spec["start_tenor"], spec["forward_tenor"],
                            calendar_name, convention_code, eom_code, settlement_days,
                        )
                        rate = forward_swap_rate(
                            curve_id,
                            start_iso,
                            end_iso,
                            spec["forward_tenor"],
                            calendar_name=calendar_name,
                            day_count_code=day_count_code,
                            convention_code=convention_code,
                            eom_code=eom_code,
                            overnight_index_code=overnight_index_code,
                            payment_lag_days=payment_lag_days,
                        )

                        prev_rate = history_store.get_previous_forward_rate(conn, spec, date_iso)
                        if prev_rate is not None and abs(rate - prev_rate) > WARN_THRESHOLD:
                            diff_bp = (rate - prev_rate) * 10000
                            warnings.append(
                                f"{date_iso} {spec['start_tenor']}x{spec['forward_tenor']}: "
                                f"前日比{diff_bp:+.1f}bp ({prev_rate:.4f} -> {rate:.4f})"
                            )

                        new_rows_for_date.append(
                            {
                                "as_of_date": date_iso,
                                "start_tenor": spec["start_tenor"],
                                "forward_tenor": spec["forward_tenor"],
                                "index_name": index_name,
                                "ccy": ccy,
                                "rate": rate,
                                "source": "COMPUTED",
                            }
                        )
                except Exception as exc:
                    # この日はカーブが不安定(収束しない等)でまるごとスキップし、
                    # 他の日付の計算は続ける。原因調査用に警告として記録する。
                    warnings.append(
                        f"{date_iso}: カーブ構築/計算に失敗したためスキップしました "
                        f"({type(exc).__name__}: {exc})"
                    )
                    continue

                history_store.upsert_forward_rows(conn, new_rows_for_date)
                computed += len(new_rows_for_date)
                all_new_rows.extend(new_rows_for_date)

        conn.commit()
    finally:
        conn.close()

    if all_new_rows:
        _write_audit_csv(all_new_rows)

    return {"computed": computed, "warnings": warnings}


def _resolve_period(
    valuation_date_iso: str,
    start_tenor: str,
    forward_tenor: str,
    calendar_name: str,
    convention_code: str,
    eom_code: str,
    settlement_days: int,
) -> tuple[str, str]:
    """スポット日(評価日からsettlement_days営業日後)を起点に、start_tenor分進めた
    開始日と、そこからforward_tenor分進めた終了日を返す。

    curves.forward_grid()と同じ規則(スポット日基準)。スポットは市場慣行上の
    フォワードスワップの起点であり、評価日を直接の起点にするのは誤り
    (以前は旧excel-fastapi/forward_history.pyのhistorical_forward_rate()を
    そのまま踏襲して評価日基準にしていたが、同じ旧プロジェクト内のcurves.
    forward_grid()はスポット日基準であり矛盾していた。市場慣行に合わせてこちらに統一する)。
    """
    import QuantLib as ql

    from qlutils import get_business_day_convention, get_calendar, get_end_of_month, parse_period

    calendar = get_calendar(calendar_name)
    convention = get_business_day_convention(convention_code)
    is_eom = get_end_of_month(eom_code)

    y, m, d = (int(x) for x in valuation_date_iso.split("-"))
    ql_valuation = ql.Date(d, m, y)
    ql_spot = calendar.advance(ql_valuation, settlement_days, ql.Days)

    start_period = parse_period(start_tenor)
    fwd_period = parse_period(forward_tenor)

    ql_start = calendar.advance(ql_spot, start_period, convention, is_eom)
    ql_end = calendar.advance(ql_start, fwd_period, convention, is_eom)

    start_iso = f"{ql_start.year():04d}-{ql_start.month():02d}-{ql_start.dayOfMonth():02d}"
    end_iso = f"{ql_end.year():04d}-{ql_end.month():02d}-{ql_end.dayOfMonth():02d}"
    return start_iso, end_iso
