"""FastAPIアプリ本体。ExcelのVBAからPOSTされたヒストリカルレートをSQLiteに取り込む。"""
from fastapi import FastAPI
from pydantic import BaseModel

import history_store

app = FastAPI()


class HistoryRow(BaseModel):
    """1本のレコード(1日×1テナー)。ExportToDb.basがJPY_TONA_Histシートの
    1セルごとにこの形で送ってくる。rate_pctはExcel/Bloomberg表記の%のまま
    (小数への変換はhistory_store.upsert_rows側で行う)。"""

    as_of_date: str
    tenor: str
    rate_pct: float


class HistoryIngestRequest(BaseModel):
    """/historyのリクエストボディ全体。rows以外(index_name/ccy/source)は
    その送信分すべてに共通する値として1回だけ渡す(1セルごとに繰り返さない)。"""

    index_name: str = "TONA"
    ccy: str = "JPY"
    source: str = "EXCEL_VBA"
    rows: list[HistoryRow]


@app.post("/history")
def ingest_history(payload: HistoryIngestRequest):
    """ExportToDb.basからのPOSTを受け、rows全件をswap_history.dbにUPSERTする。
    取り込んだ件数を返す(VBA側でMsgBox表示/ログ記録に使う)。"""
    n = history_store.upsert_rows(
        index_name=payload.index_name,
        ccy=payload.ccy,
        source=payload.source,
        rows=[r.model_dump() for r in payload.rows],
    )
    return {"ingested": n}
