"""FastAPIアプリ本体。ExcelのVBAからPOSTされたヒストリカルレートをSQLiteに取り込む。
また、フォワードスワップレートの計算(forward_history.py)のトリガーと、分析Excel
(Power Query)向けの読み出しも提供する。"""
from typing import Optional

from fastapi import FastAPI
from pydantic import BaseModel

import forward_history
import forward_history_short
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


class ForwardSpecItem(BaseModel):
    """追跡対象のフォワードスワップ1本。Excelのconfig_fwdシートの1行に対応する。"""

    start_tenor: str
    forward_tenor: str
    index_name: str = "TONA"
    ccy: str = "JPY"
    active: bool = True


@app.post("/forwardspec/sync")
def sync_forward_spec(specs: list[ForwardSpecItem]):
    """ForwardSpec.basからのPOSTを受け、forward_specテーブルをspecsの内容で
    完全に置き換える(Excelのconfig_fwdシートが正データ、DB側はその日次ミラー)。
    forward_history(計算済みレートの履歴)は一切変更しない。"""
    n = history_store.sync_forward_spec([s.model_dump() for s in specs])
    return {"synced": n}


class ForwardHistoryRefreshRequest(BaseModel):
    """/forwardhistory/refreshのリクエストボディ。max_backfill_daysは
    ExcelのControlシート(最大バックフィル日数セル)から渡される想定。"""

    max_backfill_days: int = 365


@app.post("/forwardhistory/refresh")
def refresh_forward_history(payload: ForwardHistoryRefreshRequest):
    """アクティブなforward_specについて、欠損している日付分のフォワードスワップ
    レートをまとめて計算しforward_historyにUPSERTする(forward_history.refresh_missing)。
    通常運用では実質「前回実行からの差分(だいたい当日1件)」だけが計算される。
    計算件数と、前日比で大きく変動した値への警告メッセージを返す。"""
    return forward_history.refresh_missing(max_backfill_days=payload.max_backfill_days)


@app.get("/swaphistory")
def get_swap_history(since: Optional[str] = None):
    """分析ExcelのPower Queryから呼ぶ想定の読み出し専用エンドポイント。
    スポットのTONAスワップレート(swap_history)をロング形式(1行=1日×1テナー)で
    全件(またはsince以降)返す。rateは小数で、/forwardhistoryと単位が揃っている。
    tenor_yearsはテナーを短い順に並べるためのソートキー。"""
    return {"rows": history_store.load_swap_history_long(since=since)}


@app.get("/forwardhistory")
def get_forward_history(since: Optional[str] = None):
    """分析ExcelのPower Queryから呼ぶ想定の読み出し専用エンドポイント。
    ロング形式(1行=1日×1スペック)で全件(またはsince以降)を返す。"""
    return {"rows": history_store.load_forward_history(since=since)}


# ---------------------------------------------------------------------------
# 短期セクターのフォワードスワップレート(config_fwd_shortシート)
# 長期側(/forwardspec/sync, /forwardhistory/refresh, /forwardhistory)と同じ形の
# エンドポイントを、保存先テーブル(forward_spec_short / forward_history_short)と
# 計算モジュール(forward_history_short.py)だけ変えて用意している。
# ---------------------------------------------------------------------------

@app.post("/forwardspec_short/sync")
def sync_forward_spec_short(specs: list[ForwardSpecItem]):
    """ForwardSpecShort.basからのPOSTを受け、forward_spec_shortテーブルをspecsの内容で
    完全に置き換える(Excelのconfig_fwd_shortシートが正データ)。
    forward_history_short(計算済みレートの履歴)は一切変更しない。"""
    n = history_store.sync_forward_spec([s.model_dump() for s in specs], sector="short")
    return {"synced": n}


@app.post("/forwardhistory_short/refresh")
def refresh_forward_history_short(payload: ForwardHistoryRefreshRequest):
    """アクティブなforward_spec_shortについて、欠損している日付分の短期フォワード
    スワップレートをまとめて計算しforward_history_shortにUPSERTする。"""
    return forward_history_short.refresh_missing(max_backfill_days=payload.max_backfill_days)


@app.get("/forwardhistory_short")
def get_forward_history_short(since: Optional[str] = None):
    """分析ExcelのPower Queryから呼ぶ想定の読み出し専用エンドポイント(短期セクター)。
    列構成は/forwardhistoryと同じロング形式(1行=1日×1スペック)。"""
    return {"rows": history_store.load_forward_history(since=since, sector="short")}
