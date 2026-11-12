"""11/10・11/11のスポット(GitHub issue「11/11」に貼られた値)で、日次処理と同じ
build_curve() + forward_swap_rate()を実行し、フォワードスワップレートが計算できるか
検証する。DBは読むだけで、書き込みは一切しない。

使い方(プロジェクトのフォルダで):
    python verify_forward_dates.py          # 11/10・11/11だけ検証
    python verify_forward_dates.py --all    # DBにある直近365日分すべてを検証

出力:
  1. DB(swap_history / forward_history)に各日付の行が何件あるか
  2. DB上のスポットとissueの値が一致しているか
  3. 各日付のフォワードスワップレート計算結果(失敗した場合はエラー全文)
  4. 11/10についてはDBに保存済みのフォワードとの差
"""
import sqlite3
import traceback

import history_store
from curves import build_curve, forward_swap_rate
from forward_history import _resolve_period

TENORS = ["1D", "1W", "2W", "3W", "1M", "2M", "3M", "4M", "5M", "6M", "7M", "8M", "9M",
          "10M", "11M", "1Y", "15M", "18M", "21M", "2Y", "3Y", "4Y", "5Y", "6Y", "7Y", "8Y",
          "9Y", "10Y", "11Y", "12Y", "15Y", "20Y", "25Y", "30Y", "35Y", "40Y"]

# issue「11/11」に貼られた値(%表記)
ISSUE_DATA = {
    "2026-11-11": [1.22700, 1.22750, 1.22750, 1.22750, 1.22875, 1.31000, 1.36625, 1.39375,
                   1.43625, 1.47625, 1.51250, 1.54500, 1.57625, 1.60250, 1.62625, 1.65000,
                   1.70875, 1.75750, 1.79750, 1.83875, 1.97250, 2.08875, 2.19625, 2.29750,
                   2.39875, 2.50000, 2.59750, 2.69000, 2.77625, 2.85625, 3.05500, 3.29625,
                   3.42625, 3.48250, 3.4975, 3.495],
    "2026-11-10": [1.22800, 1.22750, 1.22750, 1.22750, 1.22875, 1.29875, 1.34875, 1.37625,
                   1.41000, 1.44125, 1.47125, 1.49750, 1.52375, 1.54750, 1.56750, 1.58625,
                   1.63750, 1.67875, 1.71500, 1.75125, 1.87750, 1.98875, 2.09375, 2.19500,
                   2.29875, 2.40250, 2.50250, 2.59625, 2.68375, 2.76500, 2.97125, 3.22125,
                   3.35875, 3.42125, 3.4375, 3.435],
}

# 日次処理と同じ既定値(forward_history.refresh_missingと揃える)
CONV = dict(calendar_name="TKB", day_count_code="ACT365F", convention_code="MF",
            eom_code="STD", overnight_index_code="TONAR", settlement_days=2, payment_lag_days=2)

DEFAULT_SPECS = [("1Y", "1Y"), ("2Y", "1Y"), ("1Y", "2Y")]


def db_status(conn: sqlite3.Connection, d: str) -> None:
    n_spot = conn.execute(
        "SELECT COUNT(*) FROM swap_history WHERE as_of_date = ? AND index_name='TONA' AND ccy='JPY'", (d,)
    ).fetchone()[0]
    n_fwd = conn.execute(
        "SELECT COUNT(*) FROM forward_history WHERE as_of_date = ? AND index_name='TONA' AND ccy='JPY'", (d,)
    ).fetchone()[0]
    print(f"  DB: swap_history {n_spot}件 / forward_history {n_fwd}件")

    if n_spot:
        db = dict(conn.execute(
            "SELECT tenor, rate FROM swap_history WHERE as_of_date = ? AND index_name='TONA' AND ccy='JPY'", (d,)
        ).fetchall())
        diffs = []
        for t, pct in zip(TENORS, ISSUE_DATA[d]):
            if t not in db:
                diffs.append(f"{t}: DBに無し")
            elif abs(db[t] * 100 - pct) > 1e-9:
                diffs.append(f"{t}: DB={db[t] * 100:.5f} issue={pct:.5f}")
        extra = sorted(set(db) - set(TENORS))
        if extra:
            diffs.append(f"issueに無いテナーがDBにある: {extra}")
        print("  DBとissueの値: " + ("一致" if not diffs else "差あり → " + "; ".join(diffs)))


def main() -> None:
    specs = [(s["start_tenor"], s["forward_tenor"]) for s in history_store.load_active_forward_specs()]
    if not specs:
        print("アクティブなforward_specがDBに無いため、既定の3本で検証します")
        specs = DEFAULT_SPECS
    print(f"検証するスペック: {', '.join(a + b for a, b in specs)}\n")

    conn = history_store.connect()
    try:
        for d in sorted(ISSUE_DATA):
            print(f"=== {d} ===")
            db_status(conn, d)

            instruments = [{"tenor": t, "rate": pct / 100.0} for t, pct in zip(TENORS, ISSUE_DATA[d])]
            try:
                curve_id = build_curve(valuation_date_iso=d, instruments=instruments, **CONV)
                for start_tenor, fwd_tenor in specs:
                    start_iso, end_iso = _resolve_period(
                        d, start_tenor, fwd_tenor, CONV["calendar_name"], CONV["convention_code"],
                        CONV["eom_code"], CONV["settlement_days"],
                    )
                    rate = forward_swap_rate(
                        curve_id, start_iso, end_iso, fwd_tenor,
                        calendar_name=CONV["calendar_name"], day_count_code=CONV["day_count_code"],
                        convention_code=CONV["convention_code"], eom_code=CONV["eom_code"],
                        overnight_index_code=CONV["overnight_index_code"],
                        payment_lag_days=CONV["payment_lag_days"],
                    )
                    stored = conn.execute(
                        """SELECT rate FROM forward_history WHERE as_of_date = ? AND start_tenor = ?
                           AND forward_tenor = ? AND index_name='TONA' AND ccy='JPY'""",
                        (d, start_tenor, fwd_tenor),
                    ).fetchone()
                    note = "DBに無し" if stored is None else f"DB {stored[0] * 100:.5f}% (差 {(rate - stored[0]) * 10000:+.3f}bp)"
                    print(f"  {start_tenor}{fwd_tenor:<4} {start_iso}→{end_iso}  {rate * 100:.5f}%   [{note}]")
                print("  → 計算成功")
            except Exception:
                print("  → 計算失敗:")
                traceback.print_exc()
            print()
    finally:
        conn.close()


def main_all_db_dates(days: int = 365) -> None:
    """DBのswap_historyにある直近days日分すべてで計算を試し、失敗日と、DBに保存済みの
    フォワードとの差(カーブ構築の変更が既存値にどれだけ影響するか)をまとめて表示する。"""
    from datetime import date, timedelta

    specs = [(s["start_tenor"], s["forward_tenor"]) for s in history_store.load_active_forward_specs()] or DEFAULT_SPECS
    since = (date.today() - timedelta(days=days)).isoformat()
    conn = history_store.connect()
    failed, max_diff, n_ok = [], (0.0, None), 0
    try:
        dates = history_store.list_swap_history_dates(conn, "TONA", "JPY", since)
        print(f"{len(dates)}日分を検証します(スペック: {', '.join(a + b for a, b in specs)})")
        for d in dates:
            instruments = history_store.load_instruments_for_date(conn, d, "TONA", "JPY")
            try:
                curve_id = build_curve(valuation_date_iso=d, instruments=instruments, **CONV)
                for start_tenor, fwd_tenor in specs:
                    start_iso, end_iso = _resolve_period(
                        d, start_tenor, fwd_tenor, CONV["calendar_name"], CONV["convention_code"],
                        CONV["eom_code"], CONV["settlement_days"],
                    )
                    rate = forward_swap_rate(
                        curve_id, start_iso, end_iso, fwd_tenor,
                        calendar_name=CONV["calendar_name"], day_count_code=CONV["day_count_code"],
                        convention_code=CONV["convention_code"], eom_code=CONV["eom_code"],
                        overnight_index_code=CONV["overnight_index_code"],
                        payment_lag_days=CONV["payment_lag_days"],
                    )
                    stored = conn.execute(
                        """SELECT rate FROM forward_history WHERE as_of_date = ? AND start_tenor = ?
                           AND forward_tenor = ? AND index_name='TONA' AND ccy='JPY'""",
                        (d, start_tenor, fwd_tenor),
                    ).fetchone()
                    if stored is not None:
                        diff = abs(rate - stored[0]) * 10000
                        if diff > max_diff[0]:
                            max_diff = (diff, f"{d} {start_tenor}{fwd_tenor}")
                n_ok += 1
            except Exception as exc:
                failed.append(f"{d}: {type(exc).__name__}: {str(exc)[:120]}")
    finally:
        conn.close()

    print(f"\n成功: {n_ok}日 / 失敗: {len(failed)}日")
    for f in failed:
        print("  " + f)
    if max_diff[1]:
        print(f"\nDB保存済みの値との最大差: {max_diff[0]:.4f}bp ({max_diff[1]})")


if __name__ == "__main__":
    import sys

    if "--all" in sys.argv:
        main_all_db_dates()
    else:
        main()
