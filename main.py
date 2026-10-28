"""FastAPIアプリ本体。ExcelのVBAからPOSTされたヒストリカルレートをSQLiteに取り込む。"""
from fastapi import FastAPI
from pydantic import BaseModel

import history_store

app = FastAPI()


class HistoryRow(BaseModel):
    as_of_date: str
    tenor: str
    rate_pct: float


class HistoryIngestRequest(BaseModel):
    index_name: str = "TONA"
    ccy: str = "JPY"
    source: str = "EXCEL_VBA"
    rows: list[HistoryRow]


@app.post("/history")
def ingest_history(payload: HistoryIngestRequest):
    n = history_store.upsert_rows(
        index_name=payload.index_name,
        ccy=payload.ccy,
        source=payload.source,
        rows=[r.model_dump() for r in payload.rows],
    )
    return {"ingested": n}
