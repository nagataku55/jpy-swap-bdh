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
from typing import Optional

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

# forward_spec: 「どのフォワードスワップ(start_tenor x forward_tenor)を追跡するか」の設定。
# Excelのconfig_fwdシートの日次ミラー(sync_forward_specで全件置き換えられる)という
# 位置づけなので、ここ自体を手で編集することは想定していない。
_FORWARD_SPEC_SCHEMA_TEMPLATE = """
CREATE TABLE IF NOT EXISTS {table} (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    start_tenor TEXT NOT NULL,
    forward_tenor TEXT NOT NULL,
    index_name TEXT NOT NULL,
    ccy TEXT NOT NULL,
    active INTEGER NOT NULL DEFAULT 1,
    UNIQUE(start_tenor, forward_tenor, index_name, ccy)
)
"""

# forward_history: 計算済みのフォワードスワップレート。forward_specをActive=FALSEに
# しても、ここの既存行は削除しない(過去分は残したまま、今後の計算対象から外すだけ)。
_FORWARD_HISTORY_SCHEMA_TEMPLATE = """
CREATE TABLE IF NOT EXISTS {table} (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    as_of_date TEXT NOT NULL,
    start_tenor TEXT NOT NULL,
    forward_tenor TEXT NOT NULL,
    index_name TEXT NOT NULL,
    ccy TEXT NOT NULL,
    rate REAL NOT NULL,
    source TEXT NOT NULL,
    inserted_at TEXT NOT NULL,
    UNIQUE(as_of_date, index_name, ccy, start_tenor, forward_tenor)
)
"""

# フォワードのセクター(長期/短期)ごとの(スペック表, 計算結果表)の対応。
#   - "long" : 既存のconfig_fwdシート → forward_spec / forward_history
#   - "short": 短期用のconfig_fwd_shortシート → forward_spec_short / forward_history_short
# 短期側は将来カーブの補間方法を長期側と変える可能性があるため、計算結果を同じ
# テーブルに混ぜず、テーブルごと分けておく(どちらの方法で計算した値かが明確になる)。
# テーブル名はSQLに直接埋め込むので、必ずこの辞書経由で決める(外部入力を使わない)。
FORWARD_SECTORS = {
    "long": {"spec_table": "forward_spec", "history_table": "forward_history"},
    "short": {"spec_table": "forward_spec_short", "history_table": "forward_history_short"},
}


def _tables(sector: str) -> tuple[str, str]:
    """sector("long"/"short")から(スペック表名, 計算結果表名)を返す。"""
    try:
        t = FORWARD_SECTORS[sector]
    except KeyError:
        raise ValueError(f"Unknown forward sector: {sector!r} (supported: {', '.join(FORWARD_SECTORS)})")
    return t["spec_table"], t["history_table"]


def connect() -> sqlite3.Connection:
    """新しい接続を開く(テーブルが無ければ作成)。forward_history.refresh_missing()の
    ように同じ接続で何十〜何百回もクエリを発行する処理では、呼び出しごとにconnect()
    し直すとSQLiteへの接続オーバーヘッド(スキーマ作成チェック含む)が積み重なって
    大きく遅くなるため、呼び出し側で1回だけ開いて使い回すこと(下のconn引数を取る
    関数群はそのために存在する)。
    """
    conn = sqlite3.connect(DB_PATH)
    conn.execute(_SCHEMA)
    for t in FORWARD_SECTORS.values():
        conn.execute(_FORWARD_SPEC_SCHEMA_TEMPLATE.format(table=t["spec_table"]))
        conn.execute(_FORWARD_HISTORY_SCHEMA_TEMPLATE.format(table=t["history_table"]))
    return conn


_connect = connect  # 内部呼び出し用の別名(簡潔さのため)


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


# ---------------------------------------------------------------------------
# forward_spec: Excelのconfig_fwdシートのミラー(POST /forwardspec/sync用)
# ---------------------------------------------------------------------------

def sync_forward_spec(specs: list[dict], sector: str = "long") -> int:
    """forward_specテーブルの内容をspecsで完全に置き換える(全削除→全挿入)。

    specs: [{"start_tenor": "1Y", "forward_tenor": "1Y", "index_name": "TONA",
             "ccy": "JPY", "active": True}, ...]
    Excelのconfig_fwdシートが正データなので、ここは「今シートに書かれている
    内容そのもの」で毎回上書きする(差分更新ではない)。forward_history側のデータは
    一切触らないので、過去に計算済みのレートは常に残る。
    sector="short"の場合はconfig_fwd_shortシートのミラー(forward_spec_short)を置き換える。
    """
    spec_table, _ = _tables(sector)
    conn = _connect()
    try:
        conn.execute(f"DELETE FROM {spec_table}")
        conn.executemany(
            f"""
            INSERT INTO {spec_table} (start_tenor, forward_tenor, index_name, ccy, active)
            VALUES (?, ?, ?, ?, ?)
            """,
            [
                (s["start_tenor"], s["forward_tenor"], s["index_name"], s["ccy"], 1 if s["active"] else 0)
                for s in specs
            ],
        )
        conn.commit()
        return len(specs)
    finally:
        conn.close()


def load_active_forward_specs(sector: str = "long") -> list[dict]:
    """active=1のforward_spec(sector="short"ならforward_spec_short)を全件返す。"""
    spec_table, _ = _tables(sector)
    conn = _connect()
    try:
        cur = conn.execute(
            f"SELECT start_tenor, forward_tenor, index_name, ccy FROM {spec_table} WHERE active = 1"
        )
        return [
            {"start_tenor": r[0], "forward_tenor": r[1], "index_name": r[2], "ccy": r[3]}
            for r in cur.fetchall()
        ]
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# forward_history: 計算結果の読み書き(forward_history.pyのバックフィル処理から使う)
# ---------------------------------------------------------------------------

def list_swap_history_dates(conn: sqlite3.Connection, index_name: str, ccy: str, since_iso: str) -> list[str]:
    """since_iso以降でswap_historyに存在する観測日(as_of_date)を昇順で返す。

    「カーブが構築できる日」の一覧として使う(=その日のテナー別スポットレートが
    揃っている日)。バックフィル範囲の上限(max_backfill_days)をsince_isoで絞り込む。

    connは呼び出し側(forward_history.refresh_missing)が1回だけ開いて使い回す接続
    (以下のconn引数を取る関数も同様。何百回も呼ばれるループの中でconnect()し直すと
    接続オーバーヘッドが積み重なって大きく遅くなるため)。
    """
    cur = conn.execute(
        """
        SELECT DISTINCT as_of_date FROM swap_history
        WHERE index_name = ? AND ccy = ? AND as_of_date >= ?
        ORDER BY as_of_date ASC
        """,
        (index_name, ccy, since_iso),
    )
    return [r[0] for r in cur.fetchall()]


def load_instruments_for_date(conn: sqlite3.Connection, date_iso: str, index_name: str, ccy: str) -> list[dict]:
    """ある観測日のテナー別スポットレートを、curves.build_curve()が受け取れる形
    ([{"tenor":..., "rate":...}, ...])で返す。rateはswap_history上で既に小数
    (例0.00751)なので、そのままbuild_curve()に渡せる(%からの変換は不要)。
    """
    cur = conn.execute(
        """
        SELECT tenor, rate FROM swap_history
        WHERE as_of_date = ? AND index_name = ? AND ccy = ?
        """,
        (date_iso, index_name, ccy),
    )
    return [{"tenor": r[0], "rate": r[1]} for r in cur.fetchall()]


def missing_specs_for_date(
    conn: sqlite3.Connection, date_iso: str, index_name: str, ccy: str, specs: list[dict],
    sector: str = "long",
) -> list[dict]:
    """specsのうち、date_isoの行がforward_history(sector="short"ならforward_history_short)
    にまだ無いものだけを返す。"""
    _, history_table = _tables(sector)
    cur = conn.execute(
        f"""
        SELECT start_tenor, forward_tenor FROM {history_table}
        WHERE as_of_date = ? AND index_name = ? AND ccy = ?
        """,
        (date_iso, index_name, ccy),
    )
    existing = {(r[0], r[1]) for r in cur.fetchall()}
    return [s for s in specs if (s["start_tenor"], s["forward_tenor"]) not in existing]


def get_previous_forward_rate(
    conn: sqlite3.Connection, spec: dict, before_date_iso: str, sector: str = "long"
) -> Optional[float]:
    """spec(start_tenor/forward_tenor/index_name/ccy)について、before_date_isoより前の
    直近のforward_history.rateを返す(無ければNone)。前日比チェック(異常値警告)用。
    """
    _, history_table = _tables(sector)
    cur = conn.execute(
        f"""
        SELECT rate FROM {history_table}
        WHERE start_tenor = ? AND forward_tenor = ? AND index_name = ? AND ccy = ?
              AND as_of_date < ?
        ORDER BY as_of_date DESC
        LIMIT 1
        """,
        (spec["start_tenor"], spec["forward_tenor"], spec["index_name"], spec["ccy"], before_date_iso),
    )
    row = cur.fetchone()
    return row[0] if row else None


def upsert_forward_rows(conn: sqlite3.Connection, rows: list[dict], sector: str = "long") -> int:
    """rows: [{"as_of_date":..., "start_tenor":..., "forward_tenor":..., "index_name":...,
               "ccy":..., "rate":..., "source":...}, ...]
    rateは既に小数(例0.0123)。UNIQUE(as_of_date, index_name, ccy, start_tenor, forward_tenor)
    でUPSERTする。コミットは呼び出し側(forward_history.refresh_missing)がまとめて行う
    想定のため、ここではcommit()しない。
    """
    if not rows:
        return 0

    _, history_table = _tables(sector)
    now = datetime.now(timezone.utc).isoformat()
    conn.executemany(
        f"""
        INSERT INTO {history_table}
            (as_of_date, start_tenor, forward_tenor, index_name, ccy, rate, source, inserted_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(as_of_date, index_name, ccy, start_tenor, forward_tenor) DO UPDATE SET
            rate = excluded.rate,
            source = excluded.source,
            inserted_at = excluded.inserted_at
        """,
        [
            (
                r["as_of_date"],
                r["start_tenor"],
                r["forward_tenor"],
                r["index_name"],
                r["ccy"],
                float(r["rate"]),
                r["source"],
                now,
            )
            for r in rows
        ],
    )
    return len(rows)


def _tenor_years(tenor: str) -> Optional[float]:
    """テナー表記(ON, O/N, 1D, 1W, 3M, 10Y等)をおおよその年数に変換する。
    分析Excelでテナーを短い順に並べるためのソートキー用(文字列のままだと
    "10Y"が"1M"より前に来てしまうため)。解釈できない表記はNone。
    QuantLibを読み込まずに済むよう、qlutils.parse_periodは使わない。"""
    t = tenor.strip().upper()
    if t in ("ON", "O/N"):
        return round(1 / 365, 6)
    units = {"D": 1 / 365, "W": 7 / 365, "M": 1 / 12, "Y": 1.0}
    if len(t) >= 2 and t[-1] in units and t[:-1].isdigit():
        return round(int(t[:-1]) * units[t[-1]], 6)
    return None


def load_swap_history_long(
    since: Optional[str] = None, index_name: str = "TONA", ccy: str = "JPY"
) -> list[dict]:
    """GET /swaphistory用。swap_history(スポットのTONAスワップレート)をロング形式
    (1行=1日×1テナー)で返す。sinceを指定するとas_of_date >= sinceに絞り込む。
    rateは小数(例0.00751)のまま返す(/forwardhistoryと単位を揃えるため)。
    """
    sql = """
        SELECT as_of_date, tenor, index_name, ccy, rate
        FROM swap_history
        WHERE index_name = ? AND ccy = ?
    """
    params: list = [index_name, ccy]
    if since:
        sql += " AND as_of_date >= ?"
        params.append(since)
    sql += " ORDER BY as_of_date ASC"

    conn = _connect()
    try:
        cur = conn.execute(sql, params)
        rows = [
            {
                "as_of_date": r[0],
                "tenor": r[1],
                "tenor_years": _tenor_years(r[1]),
                "index_name": r[2],
                "ccy": r[3],
                "rate": r[4],
            }
            for r in cur.fetchall()
        ]
    finally:
        conn.close()
    # 同じ日付内はテナーの短い順に並べる(解釈できないテナーは末尾)
    rows.sort(key=lambda x: (x["as_of_date"], x["tenor_years"] is None, x["tenor_years"] or 0))
    return rows


def load_forward_history(since: Optional[str] = None, sector: str = "long") -> list[dict]:
    """GET /forwardhistory用(sector="short"ならGET /forwardhistory_short用)。
    ロング形式(1行=1日×1スペック)のレコード一覧を返す。
    sinceを指定すると、その日付以降(as_of_date >= since)だけに絞り込む。
    """
    _, history_table = _tables(sector)
    conn = _connect()
    try:
        if since:
            cur = conn.execute(
                f"""
                SELECT as_of_date, start_tenor, forward_tenor, index_name, ccy, rate
                FROM {history_table} WHERE as_of_date >= ?
                ORDER BY as_of_date ASC
                """,
                (since,),
            )
        else:
            cur = conn.execute(
                f"""
                SELECT as_of_date, start_tenor, forward_tenor, index_name, ccy, rate
                FROM {history_table}
                ORDER BY as_of_date ASC
                """
            )
        return [
            {
                "as_of_date": r[0],
                "start_tenor": r[1],
                "forward_tenor": r[2],
                "index_name": r[3],
                "ccy": r[4],
                "rate": r[5],
            }
            for r in cur.fetchall()
        ]
    finally:
        conn.close()
