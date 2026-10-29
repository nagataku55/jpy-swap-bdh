# swap (jpy-swap-bdh)

Excel(Bloomberg `BDH`関数)で取得したJPY TONAスワップのヒストリカルレートを、
FastAPI経由でSQLiteに取り込み、監査用CSVも残すためのプロジェクトです。
`excel-fastapi`プロジェクトを土台にしつつ、構成を整理して作り直しました。

## 全体構成

```
Bloomberg Terminal (BDH)
        │
        ▼
Excel ブック(jpy_tona_data.xlsm)
  ├─ Controlシート: BDH取得設定(日付範囲・ティッカー・データシート名)
  ├─ JPY_TONA_Histシート: ヒストリカルレート本体(ワイド形式)
  ├─ Configシート: CSV監査出力ジョブ一覧
  └─ UpdateLogシート: 日次実行ログ(自動生成)
        │
        ├─ ExportHistLevels.bas ──▶ CSV監査ログ(csv_audit/配下、日付付きファイル名)
        │
        └─ ExportToDb.bas ──▶ POST /history ──▶ FastAPI(main.py)
                                                      │
                                                      ▼
                                          swap_history.db (SQLite, history_store.py)

DailyUpdate.bas が上記2つ(CSV出力・DB取り込み)を1本の処理としてまとめ、
run_daily_update.ps1 + タスクスケジューラで16:30頃に無人実行する。
```

## ファイル一覧と役割

| ファイル | 役割 |
|---|---|
| `jpy_tona_data.xlsx`(運用時は`.xlsm`) | Excel本体。Control/JPY_TONA_Hist/Config/UpdateLogシートを持つ |
| `gen_bdh_workbook.py` | 上記ワークブックをゼロから生成するスクリプト(BDH数式付き)。テナー構成やレイアウトを変えたい時はこれを編集して再実行する |
| `main.py` | FastAPIアプリ本体。`POST /history` でヒストリカルレートを受け取り、SQLiteに取り込む |
| `history_store.py` | SQLite(`swap_history.db`)への保存・読み出し。レートは%表記→小数に変換して保存する |
| `requirements.txt` | Python依存パッケージ(`pip install -r requirements.txt`) |
| `ExportToDb.bas` | ExcelのVBAマクロ。`JPY_TONA_Hist`シートを読み、FastAPIの`/history`へJSONでPOSTする |
| `ExportHistLevels.bas` | ExcelのVBAマクロ。`Config`シートで指定したシートをCSVへ出力する(監査ログ用) |
| `DailyUpdate.bas` | 上記2つを「再計算待ち→エラーチェック→保存→CSV出力→DB取り込み→ログ記録」の1本の処理にまとめる |
| `run_daily_update.ps1` | Excelを非表示で開き`DailyUpdate.RunDailyUpdate_Silent`を実行する(タスクスケジューラから呼ぶ) |
| `start_api_server.ps1` | `uvicorn main:app`をバックグラウンドで起動する(ログオン時にタスクスケジューラから呼ぶ。ポートが既に使用中なら何もしないので多重起動しない) |
| `swap_history.db` | SQLite本体(git管理外。`history_store.py`が自動生成する) |

## セットアップ

### 1. Python環境

```
pip install -r requirements.txt
```

### 2. Excelワークブックの用意

- 既に`jpy_tona_data.xlsx`があるので、そのまま使ってよい
- レイアウトやテナー構成を変えたい/一から作り直したい場合は `python gen_bdh_workbook.py` を実行する
  (`Control`シートのティッカー欄B9:B22が空欄で生成されるので、実際のBloombergティッカーを入力すること)
- マクロを使うには、Excelで開いて「名前を付けて保存」→「Excelマクロ有効ブック(.xlsm)」で保存し直す

### 3. VBAモジュールの組み込み

1. `.xlsm`を開き、VBEditor(Alt+F11)を起動
2. 「ファイル→ファイルのインポート」で以下を取り込む
   - `ExportToDb.bas`
   - `ExportHistLevels.bas`
   - `DailyUpdate.bas`
3. シート上にボタンを置き、`ExportToDb.RunExportToDb_Manual`や`DailyUpdate.RunDailyUpdate_Manual`を割り当てると手動実行できる

## 使い方

### A. ヒストリカルレートを取得する(Control / JPY_TONA_Histシート)

1. `Control`シートで開始日・終了日・価格フィールド・ティッカーを設定する
2. `JPY_TONA_Hist`シートのB3以降にあるBDH数式が、Bloombergアドイン有効な端末で再計算されて値が入る
   (Bloombergが無い環境では、ダミー値を直接貼り付けて動作確認する運用でもよい)

### B. FastAPIサーバーを起動する

```
uvicorn main:app
```

デフォルトで`http://127.0.0.1:8000`に起動する。`ExportToDb.bas`の`API_URL`定数がこのURLを前提にしているので、
ポートを変える場合は両方合わせて変更すること。

### C. 手動でDBに取り込む

Excelでボタンを押して`ExportToDb.RunExportToDb_Manual`を実行する。
`JPY_TONA_Hist`シートの全データ(B列=Date, C列以降=テナー別レート)をJSONにして`/history`へPOSTし、
`swap_history.db`に取り込まれる(結果はMsgBoxで表示)。

### D. CSV監査ログを出力する

Excelでボタンを押して`ExportHistLevels.ExportAllSheetsToCsv`を実行する。
`Config`シートのテーブル(`SheetName | SaveFolder | FileName | DataStartCell`)に従って、
指定したシートをCSVに書き出す。

- `SaveFolder`が絶対パス(`C:\...`や`\\server\...`)でなければ、**ブックが置かれているフォルダからの相対パス**として解決される(フォルダが無ければ自動作成)。PCが変わってもセルの書き換えが不要
- `FileName`に`{date}`を含めると実行日(`yyyy-mm-dd`)に置き換わる(例: `JPY_TONA_Hist_{date}.csv` → `JPY_TONA_Hist_2026-10-29.csv`)。日次で別ファイルとして残るので、監査ログを上書きしない

### E. 日次更新を1本で行う(手動)

Excelでボタンを押して`DailyUpdate.RunDailyUpdate_Manual`を実行すると、以下を順に行う。

1. 完全再計算を実行し、BDHの再計算完了を待つ(最大60秒)
2. `JPY_TONA_Hist`シートにエラー値が無いか確認(あれば中止。中途半端なデータを取り込まないため)
3. ワークブックを保存
4. CSV監査ログを出力(`ExportHistLevels`)
5. FastAPI経由でDBに取り込み(`ExportToDb`)
6. 結果を`UpdateLog`シートに追記

### F. 16:30頃に自動実行する(タスクスケジューラ)

FastAPIサーバーを常駐させる(ログオン時に自動起動)。`start_api_server.ps1`は、
指定ポートが既に使用中なら何もしない(ログオンのたびに重複起動してWinError 10048
「アドレス使用中」になるのを防ぐ):

```
schtasks /Create /TN "SwapHistory_ApiServer" /TR "powershell -ExecutionPolicy Bypass -File C:\python\swap\start_api_server.ps1" /SC ONLOGON /RL LIMITED
```

16:30に`DailyUpdate.RunDailyUpdate_Silent`を実行:

```
schtasks /Create /TN "SwapHistory_DailyUpdate" /TR "powershell -ExecutionPolicy Bypass -File C:\python\swap\run_daily_update.ps1 -WorkbookPath C:\python\swap\jpy_tona_data.xlsm" /SC DAILY /ST 16:30 /RL LIMITED
```

タスクの実行設定は「ユーザーがログオンしているときのみ実行」にすること(Bloombergは対話ログイン前提のため)。
パスは実際の配置先に合わせて変更する。

## `swap_history`テーブル構造

| 列名 | 型 | 内容 |
|---|---|---|
| `id` | INTEGER PK | 自動採番 |
| `as_of_date` | TEXT | 観測日(`YYYY-MM-DD`) |
| `index_name` | TEXT | 参照金利(現状`TONA`のみ) |
| `tenor` | TEXT | テナー(`ON`, `1W`, `1M`, ... `30Y`) |
| `ccy` | TEXT | 通貨(`JPY`) |
| `rate` | REAL | スワップレート(**小数**。例`0.00751`=0.751%) |
| `source` | TEXT | 取得元(`EXCEL_VBA`) |
| `inserted_at` | TEXT | 取り込み日時(UTC ISO8601) |

`UNIQUE(as_of_date, index_name, tenor)` — 同じキーの再取り込みはUPSERT(上書き)される。

## 既知の注意点

- **レートの単位**: Excel/Bloomberg側は%表記(例`0.751`=0.751%)、SQLite側は小数(`0.00751`)。
  `ExportToDb.bas`はセルの値をそのまま(%表記のまま)送り、`history_store.upsert_rows`が小数に変換して保存する
- **ポート番号を変える場合**: `main.py`を`uvicorn main:app --port XXXX`で起動したら、
  `ExportToDb.bas`の`API_URL`定数も同じポートに変更すること
- **Excelファイルがロックされてエラーになる場合**: `.xlsx`/`.xlsm`をExcelで開いたまま
  Python側から書き込もうとすると`PermissionError`になる。Excelを閉じてから実行する
- **`swap_history.db`はgit管理外**: `history_store.py`が接続時に自動でテーブルを作成するため、
  ファイルが無い状態から始めても問題ない
