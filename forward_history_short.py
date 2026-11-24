"""短期セクターのフォワードスワップレート(例: 1M1M, 3M3M, 6M6M)のヒストリカルデータを
swap_history(スポットレート)から作成・更新するモジュール。

長期側(forward_history.py)とは別のモジュールにしている理由:
  - 短期側は、将来カーブの補間方法を長期側と変える可能性がある(短期は中銀の
    政策金利の段差が効くので、長期と同じLogCubic(Kruger)のままがよいとは限らない)。
    補間を変える時は、このファイルの_build_curve_for_date()だけを差し替えればよく、
    長期側の計算結果には一切影響しない。
  - 計算結果の保存先も分けている(forward_spec_short / forward_history_short)。
    同じテーブルに混ぜると、どの補間方法で計算した値なのか区別がつかなくなるため。

現時点の_build_curve_for_date()は長期側と同じcurves.build_curve()を使っている。
同じ日付・同じ入力なら同じcurve_id(内容ハッシュ)になり、curves側のキャッシュが
効くので、長期側の計算直後に実行すればカーブの再構築はほぼ発生しない。

それ以外の設計(欠損日をまとめて計算する1本のロジック、max_backfill_days、前日比警告、
1日失敗しても続行、監査CSV出力)は長期側のforward_history.pyと同じ。
"""
import csv
from datetime import date, timedelta
from itertools import groupby
from pathlib import Path

import history_store
from curves import build_curve, forward_swap_rate
from forward_history import _resolve_period, _tenor_sort_key

SECTOR = "short"

# 前日比でこれを超えて動いたら警告(長期側と同じ50bp)。短期は日々の変動が小さいので、
# 取得ミスをもっと敏感に拾いたければここを小さくする(例: 0.0025 = 25bp)。
WARN_THRESHOLD = 0.0050

AUDIT_DIR = Path(__file__).parent / "csv_audit"
AUDIT_FILE_PREFIX = "ForwardHistoryShort"


def _build_curve_for_date(
    date_iso: str,
    instruments: list[dict],
    calendar_name: str,
    day_count_code: str,
    overnight_index_code: str,
    convention_code: str,
    settlement_days: int,
    eom_code: str,
    payment_lag_days: int,
) -> str:
    """短期フォワード計算用のカーブを1日分構築し、curve_idを返す。

    【短期側の補間方法を変えたい時はここを差し替える】
    現在は長期側と同じcurves.build_curve()(LogCubic・Kruger法)をそのまま使っている。
    別の補間(例: 割引率の対数線形=フラットフォワード)にする場合は、curves.pyに
    その方法でカーブを作る関数を追加し(戻り値は同じくcurve_idで、curves._curve_cacheに
    登録されていること)、ここで呼び出す関数をそれに変えればよい。
    forward_swap_rate()はcurve_idさえあれば補間方法に関係なく動く。
    """
    return build_curve(
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


def _write_audit_csv() -> "Path | None":
    """forward_history_shortの全履歴を、長期側と同じワイド形式
    (1列目=Date、2列目以降=スペック(例: "3M3M")ごとのレート列)で
    csv_audit/ForwardHistoryShort_<実行日>.csvに書き出す。
    同日に複数回実行された場合は上書き。データがまだ無い場合はNoneを返す。
    """
    rows = history_store.load_forward_history(sector=SECTOR)
    if not rows:
        return None

    specs = sorted({(r["start_tenor"], r["forward_tenor"]) for r in rows}, key=_tenor_sort_key)
    spec_labels = {spec: f"{spec[0]}{spec[1]}" for spec in specs}

    by_date: dict[str, dict[str, float]] = {}
    for r in rows:
        label = spec_labels[(r["start_tenor"], r["forward_tenor"])]
        by_date.setdefault(r["as_of_date"], {})[label] = r["rate"]

    AUDIT_DIR.mkdir(exist_ok=True)
    out_path = AUDIT_DIR / f"{AUDIT_FILE_PREFIX}_{date.today().isoformat()}.csv"

    fieldnames = ["Date"] + [spec_labels[spec] for spec in specs]
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for d in sorted(by_date.keys()):
            row = {"Date": d}
            row.update(by_date[d])
            writer.writerow(row)

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
    """アクティブなforward_spec_shortごとに、swap_historyにあるがforward_history_shortに
    無い日付をまとめて計算してUPSERTする。戻り値: {"computed": 件数, "warnings": [...]}。

    スタート日・終了日の決め方は長期側と同じ(スポット日を起点にstart_tenor分進め、
    そこからforward_tenor分進める。Modified Following)。forward_tenorが12ヶ月以下の
    スペックは固定レグ一括払い(curves.forward_swap_rate()の規則)になる。
    """
    specs = history_store.load_active_forward_specs(sector=SECTOR)
    if not specs:
        return {"computed": 0, "warnings": [], "message": "アクティブなforward_spec_shortがありません"}

    since_iso = (date.today() - timedelta(days=max_backfill_days)).isoformat()

    computed = 0
    warnings: list[str] = []

    # 長期側と同じく、接続はこの処理全体で1回だけ開いて使い回す(速度のため)。
    conn = history_store.connect()
    try:
        key = lambda s: (s["index_name"], s["ccy"])
        for (index_name, ccy), group_iter in groupby(sorted(specs, key=key), key=key):
            group_specs = list(group_iter)
            dates = history_store.list_swap_history_dates(conn, index_name, ccy, since_iso)

            for date_iso in dates:
                missing = history_store.missing_specs_for_date(
                    conn, date_iso, index_name, ccy, group_specs, sector=SECTOR
                )
                if not missing:
                    continue

                instruments = history_store.load_instruments_for_date(conn, date_iso, index_name, ccy)
                if not instruments:
                    continue

                try:
                    # カーブ構築はこの日付につき1回だけ(全スペックで使い回す)。
                    # QuantLibの遅延評価のため、収束失敗の例外は下のforward_swap_rate()で起きる。
                    curve_id = _build_curve_for_date(
                        date_iso,
                        instruments,
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

                        prev_rate = history_store.get_previous_forward_rate(conn, spec, date_iso, sector=SECTOR)
                        if prev_rate is not None and abs(rate - prev_rate) > WARN_THRESHOLD:
                            diff_bp = (rate - prev_rate) * 10000
                            warnings.append(
                                f"[短期] {date_iso} {spec['start_tenor']}x{spec['forward_tenor']}: "
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
                    # この日だけスキップして他の日付の計算は続ける(長期側と同じ方針)。
                    warnings.append(
                        f"[短期] {date_iso}: カーブ構築/計算に失敗したためスキップしました "
                        f"({type(exc).__name__}: {exc})"
                    )
                    continue

                history_store.upsert_forward_rows(conn, new_rows_for_date, sector=SECTOR)
                computed += len(new_rows_for_date)

        conn.commit()
    finally:
        conn.close()

    _write_audit_csv()

    return {"computed": computed, "warnings": warnings}
