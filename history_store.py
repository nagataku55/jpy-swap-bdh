"""JPY TONAスワップのヒストリカルレートをSQLite(swap_history.db)に保存/読み出す。

テーブル構造は旧プロジェクト(excel-fastapi/history_store.py)のswap_historyを踏襲する。
UNIQUE(as_of_date, index_name, tenor)でUPSERTするため、同じ日・同じテナーの再取り込みは
上書きされ、重複しない。

レートはExcel/Bloomberg側の表記(%、例0.523)で受け取り、QuantLibが期待する小数
(0.00523)に変換して保存する。
"""
import sqlite3
from datetime import datetime, timezone
from pathlib import Path

DB_PATH = Path(__file__).parent / "swap_history.db"

_SCHEMA = """
CREATE TABLE IF NOT EXISTS swap_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    as_of_date TEXT NOT NULL,
    index_name TEXT NOT NULL,
    tenor TEXT NOT NULL,
    ccy TEXT NOT NULL,
    rate REAL NOT NULL,
    source TEXT NOT NULL,
    inserted_at TEXT NOT NULL,
    UNIQUE(as_of_date, index_name, tenor)
)
"""


def _connect() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.execute(_SCHEMA)
    return conn


def upsert_rows(index_name: str, ccy: str, source: str, rows: list[dict]) -> int:
    """rows: [{"as_of_date": "YYYY-MM-DD", "tenor": "ON", "rate_pct": 0.751}, ...]

    rate_pctは%表記(例0.751=0.751%)。小数(0.00751)に変換して保存する。
    戻り値: 取り込んだ行数。
    """
    now = datetime.now(timezone.utc).isoformat()
    conn = _connect()
    try:
        conn.executemany(
            """
            INSERT INTO swap_history (as_of_date, index_name, tenor, ccy, rate, source, inserted_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(as_of_date, index_name, tenor) DO UPDATE SET
                rate = excluded.rate,
                ccy = excluded.ccy,
                source = excluded.source,
                inserted_at = excluded.inserted_at
            """,
            [
                (r["as_of_date"], index_name, r["tenor"], ccy, float(r["rate_pct"]) / 100.0, source, now)
                for r in rows
            ],
        )
        conn.commit()
        return len(rows)
    finally:
        conn.close()


def load_history(index_name: str = "TONA", ccy: str = "JPY"):
    """日付×テナーの横持ちDataFrameを返す(相関分析・PCA用)。"""
    import pandas as pd

    conn = _connect()
    try:
        df = pd.read_sql_query(
            "SELECT as_of_date, tenor, rate FROM swap_history WHERE index_name = ? AND ccy = ?",
            conn,
            params=(index_name, ccy),
        )
    finally:
        conn.close()

    wide = df.pivot(index="as_of_date", columns="tenor", values="rate")
    wide.index = pd.to_datetime(wide.index)
    return wide.sort_index()
