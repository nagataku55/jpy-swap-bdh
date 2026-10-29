# swap (jpy-swap-bdh)

Excel(Bloomberg `BDH`関数)で取得したJPY TONAスワップのヒストリカルレートを、
FastAPI経由でSQLiteに取り込み、そこからフォワードスワップレートのヒストリカルデータを
計算して分析Excelに取り込むためのプロジェクトです。
`excel-fastapi`プロジェクトを土台にしつつ、構成を整理して作り直しました。

## 全体構成

```
Bloomberg Terminal (BDH)
        │
        ▼
Excel ブック(jpy_tona_data.xlsm)
  ├─ Controlシート: BDH取得設定・データシート名・最大バックフィル日数(既定365日)
  ├─ JPY_TONA_Histシート: スポットレート本体(ワイド形式)
  ├─ config_csvシート: CSV監査出力ジョブ一覧(スポット用)
  ├─ config_fwdシート: 追跡するフォワードスワップ一覧(StartTenor|ForwardTenor|Active)
  └─ UpdateLogシート: 日次実行ログ(自動生成)
        │
        ├─ ExportHistLevels.bas ──▶ CSV監査ログ(csv_audit/配下、日付付きファイル名)
        │
        ├─ ExportToDb.bas ──▶ POST /history ─────────┐
        ├─ ForwardSpec.bas ──▶ POST /forwardspec/sync ─┤
        └─ ForwardHistory.bas ─▶ POST /forwardhistory/refresh ┤
                                                               ▼
                                                     FastAPI(main.py)
                                                               │
                                     ┌─────────────────────────┼─────────────────────┐
                                     ▼                         ▼                     ▼
                          swap_history.db              forward_spec           forward_history
                        (history_store.py)          (Excelのconfig_fwd      (計算済みフォワード
                                                      シートの日次ミラー)      スワップレート)
                                                               │
                                               curves.py / qlutils.py / forward_history.py
                                               (カーブ再構築→forward_swap_rate()でパーレート計算)

DailyUpdate.bas が上記を1本の処理としてまとめ、run_daily_update.ps1 + タスクスケジューラで
16:30頃に無人実行する。

分析Excel(別ワークブック、同一PC) ── Power Query(Webから取得) ──▶ GET /forwardhistory
  └─ RawDataシート(ロング形式テーブル) → ピボットテーブル/ピボットグラフで分析
```

## ファイル一覧と役割

| ファイル | 役割 |
|---|---|
| `jpy_tona_data.xlsx`(運用時は`.xlsm`) | Excel本体。Control/JPY_TONA_Hist/config_csv/config_fwd/UpdateLogシートを持つ |
| `gen_bdh_workbook.py` | 上記ワークブックをゼロから生成するスクリプト(BDH数式付き)。テナー構成やレイアウトを変えたい時はこれを編集して再実行する |
| `main.py` | FastAPIアプリ本体。`/history`・`/forwardspec/sync`・`/forwardhistory/refresh`・`/forwardhistory`を提供 |
| `history_store.py` | SQLite(`swap_history.db`)への保存・読み出し(`swap_history`/`forward_spec`/`forward_history`の3テーブル) |
| `curves.py` | QuantLibでOISカーブをブートストラップし、割引率・フォワードレート・フォワードスワップレート(`forward_swap_rate`)・PV・バケット感応度を計算する |
| `qlutils.py` | カレンダー・日数計算・営業日調整・テナー解析などのQuantLib変換ヘルパー |
| `forward_history.py` | `swap_history`にあるがまだ`forward_history`に無い日付をまとめて計算する(`refresh_missing`)。日次実行と初回/欠損バックフィルを区別しない1本のロジック |
| `requirements.txt` | Python依存パッケージ(`pip install -r requirements.txt`) |
| `ExportToDb.bas` | ExcelのVBAマクロ。`JPY_TONA_Hist`シートを読み、FastAPIの`/history`へJSONでPOSTする |
| `ExportHistLevels.bas` | ExcelのVBAマクロ。`config_csv`シートで指定したシートをCSVへ出力する(スポットの監査ログ用) |
| `ForwardSpec.bas` | ExcelのVBAマクロ。`config_fwd`シートの内容を`/forwardspec/sync`へPOSTし、DBの`forward_spec`を同期する |
| `ForwardHistory.bas` | ExcelのVBAマクロ。`/forwardhistory/refresh`を呼び、欠損しているフォワードスワップレートを計算させる |
| `DailyUpdate.bas` | 上記を「再計算待ち→エラーチェック→保存→CSV出力→DB取り込み→ForwardSpec同期→ForwardHistory更新→ログ記録」の1本の処理にまとめる |
| `run_daily_update.ps1` | Excelを非表示で開き`DailyUpdate.RunDailyUpdate_Silent`を実行する(タスクスケジューラから呼ぶ) |
| `start_api_server.ps1` | `uvicorn main:app`をバックグラウンドで起動する(ログオン時にタスクスケジューラから呼ぶ。ポートが既に使用中なら何もしないので多重起動しない) |
| `swap_history.db` | SQLite本体(git管理外。`history_store.py`が自動生成する) |
| `csv_audit/` | スポット監査CSV(`ExportHistLevels.bas`)とフォワード監査CSV(`forward_history.py`が自動出力)の保存先。git管理外 |

## セットアップ

### 1. Python環境

```
pip install -r requirements.txt
```

QuantLib(カーブ構築)・pandas・openpyxl(ワークブック生成)・fastapi/uvicorn(API)が入る。

### 2. Excelワークブックの用意

- 既に`jpy_tona_data.xlsx`があるので、そのまま使ってよい
- レイアウトやテナー構成を変えたい/一から作り直したい場合は `python gen_bdh_workbook.py` を実行する
  (`Control`シートのティッカー欄B9:B22が空欄で生成されるので、実際のBloombergティッカーを入力すること。
  ただしこのスクリプトは`config_fwd`シートや最大バックフィル日数セルは作らないので、再生成後は
  本READMEのセットアップ内容を参考に手動で追加すること)
- マクロを使うには、Excelで開いて「名前を付けて保存」→「Excelマクロ有効ブック(.xlsm)」で保存し直す

### 3. VBAモジュールの組み込み

1. `.xlsm`を開き、VBEditor(Alt+F11)を起動
2. 「ファイル→ファイルのインポート」で以下を取り込む
   - `ExportToDb.bas`
   - `ExportHistLevels.bas`
   - `ForwardSpec.bas`
   - `ForwardHistory.bas`
   - `DailyUpdate.bas`
3. シート上にボタンを置き、`DailyUpdate.RunDailyUpdate_Manual`等を割り当てると手動実行できる

## 使い方

### A. スポットのヒストリカルレートを取得する(Control / JPY_TONA_Histシート)

1. `Control`シートで開始日・終了日・価格フィールド・ティッカーを設定する
2. `JPY_TONA_Hist`シートのB3以降にあるBDH数式が、Bloombergアドイン有効な端末で再計算されて値が入る
   (Bloombergが無い環境では、ダミー値を直接貼り付けて動作確認する運用でもよい)

### B. FastAPIサーバーを起動する

```
uvicorn main:app
```

デフォルトで`http://127.0.0.1:8000`に起動する。各`.bas`の`API_URL`定数がこのURLを前提にしているので、
ポートを変える場合はすべて合わせて変更すること。

### C. 手動でスポットをDBに取り込む

Excelでボタンを押して`ExportToDb.RunExportToDb_Manual`を実行する。
`JPY_TONA_Hist`シートの全データ(B列=Date, C列以降=テナー別レート)をJSONにして`/history`へPOSTし、
`swap_history.db`に取り込まれる(結果はMsgBoxで表示)。

### D. CSV監査ログを出力する(スポット)

Excelでボタンを押して`ExportHistLevels.ExportAllSheetsToCsv`を実行する。
`config_csv`シートのテーブル(`SheetName | SaveFolder | FileName | DataStartCell`)に従って、
指定したシートをCSVに書き出す。

- `SaveFolder`が絶対パス(`C:\...`や`\\server\...`)でなければ、**ブックが置かれているフォルダからの相対パス**として解決される(フォルダが無ければ自動作成)。PCが変わってもセルの書き換えが不要
- `FileName`に`{date}`を含めると実行日(`yyyy-mm-dd`)に置き換わる(例: `JPY_TONA_Hist_{date}.csv` → `JPY_TONA_Hist_2026-10-29.csv`)。日次で別ファイルとして残るので、監査ログを上書きしない

### E. フォワードスワップレートを追加/計算する

1. `config_fwd`シートに行を追加する(`StartTenor`, `ForwardTenor`, `Active`=TRUE)。例: `2Y`, `2Y`, TRUE
2. `ForwardSpec.RunForwardSpecSync_Manual`を実行 → DBの`forward_spec`がシートの内容で同期される
3. `ForwardHistory.RunForwardHistoryRefresh_Manual`を実行 → 新しく追加したスペックの過去分(`Control!最大バックフィル日数`まで)がまとめて計算され、`forward_history`に入る
4. 追跡を止めたいだけなら行を消さず`Active`を`FALSE`にする(過去データは残る)

計算対象の判定は「`swap_history`にあるが`forward_history`に無い日付」を毎回全部探すという1つのロジックなので、
初回構築・エラー後の欠損・新規スペック追加のどの場合でも同じボタンで対応できる。
通常運用(前日までの分は既に計算済み)では実質「当日分だけ」の計算になり、数秒以内に終わる。
**ただし初回の全履歴バックフィルや新規スペック追加時は、カーブ再構築が日数分積み重なるため
時間がかかる(目安: 200日で約1分)。** `ForwardHistory.bas`の受信タイムアウトは5分に設定している。

計算結果が前日比で±50bp(`forward_history.WARN_THRESHOLD`)を超えて動いた場合、取り込み自体は
継続しつつ警告メッセージを返す(`UpdateLog`シートやMsgBoxに出る)。入力(スポットカーブ)側の
取得ミスに気づく手がかりとして使う。

### F. 日次更新を1本で行う(手動)

Excelでボタンを押して`DailyUpdate.RunDailyUpdate_Manual`を実行すると、以下を順に行う。

1. 完全再計算を実行し、BDHの再計算完了を待つ(最大60秒)
2. `JPY_TONA_Hist`シートにエラー値が無いか確認(あれば中止。中途半端なデータを取り込まないため)
3. ワークブックを保存
4. CSV監査ログを出力(`ExportHistLevels`)
5. FastAPI経由でスポットをDBに取り込み(`ExportToDb`)
6. `config_fwd`シートをDBに同期(`ForwardSpec`)
7. フォワードスワップレートの欠損分を計算(`ForwardHistory`)。実質当日分のみ
8. 結果を`UpdateLog`シートに追記

手順5(スポット取り込み)が失敗した場合は6・7を実行せず中止する。6・7で問題があっても、
スポット取り込み自体は既に成功しているため全体を失敗扱いにはしない(警告として記録するのみ)。

### G. 16:30頃に自動実行する(タスクスケジューラ)

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

### H. 分析Excelにフォワードスワップレートを取り込む(Power Query)

分析用の別ワークブック(同一PC)で、「データ」タブ→「データの取得」→「Web」で
`http://127.0.0.1:8000/forwardhistory` を指定し、「テーブルとして読み込み」を選ぶ。

- ロング形式(1行=1日×1スペック: `as_of_date, start_tenor, forward_tenor, index_name, ccy, rate`)のまま
  専用シート(例: `RawData`)にテーブルとして取り込む
- そのテーブルを元に、別シートでピボットテーブル/ピボットグラフを作って分析する
  (新スペックを追加してもフィールドリストに自動で出てくるだけで済む。Power Query側の変換ステップを
  都度編集する必要が無い)
- 更新は「データ→すべて更新」。全行が最新の内容に置き換わる(差分追記ではない)
- 必要に応じて`GET /forwardhistory?since=YYYY-MM-DD`で絞り込みも可能(将来、履歴が増えた時用)

## DBテーブル構造

### `swap_history`(スポットレート)

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

### `forward_spec`(追跡するフォワードスワップの一覧。`config_fwd`シートのミラー)

| 列名 | 型 | 内容 |
|---|---|---|
| `start_tenor` / `forward_tenor` | TEXT | 例: `2Y` x `1Y` = 「2年後スタート・1年物」(2Y1Y) |
| `index_name` / `ccy` | TEXT | `TONA` / `JPY` |
| `active` | INTEGER | 1=計算対象、0=対象外(過去データは残る) |

`sync_forward_spec()`が呼ばれるたびに全件削除→全件挿入で置き換わる(差分更新ではない)。

### `forward_history`(計算済みフォワードスワップレート)

| 列名 | 型 | 内容 |
|---|---|---|
| `as_of_date` | TEXT | 観測日 |
| `start_tenor` / `forward_tenor` | TEXT | フォワードスワップの仕様 |
| `index_name` / `ccy` | TEXT | `TONA` / `JPY` |
| `rate` | REAL | フォワードスワップのパーレート(**小数**) |
| `source` | TEXT | `COMPUTED`(サーバー側で計算した値であることを示す) |
| `inserted_at` | TEXT | 計算日時(UTC ISO8601) |

`UNIQUE(as_of_date, index_name, ccy, start_tenor, forward_tenor)` — 同じキーはUPSERT(上書き)される。
`forward_spec`を`Active=FALSE`にしても、ここの既存行は削除されない。

## 既知の注意点

- **レートの単位**: Excel/Bloomberg側は%表記(例`0.751`=0.751%)、SQLite側は小数(`0.00751`)。
  `ExportToDb.bas`はセルの値をそのまま(%表記のまま)送り、`history_store.upsert_rows`が小数に変換して保存する
- **フォワードレート ≠ フォワードスワップレート**: `curves.forward_rate()`は2点間の割引率だけから出す
  単一区間のポイントフォワードレートで、`forward_tenor`が1年超(年次複数回払い)だと本来の
  パーレートとは一致しない。フォワード**スワップ**レートが欲しい場合は必ず`curves.forward_swap_rate()`
  (`swap.fairRate()`を使う)を使うこと
- **カーブ構築は1回あたり約0.3秒かかる**: `build_curve()`はKruger法によるグローバル補間(LogCubic)を
  使っており、ローカル補間より計算コストが高い(旧プロジェクトの設計をそのまま踏襲)。日次運用では
  新規1日分だけなので問題にならないが、初回の全履歴バックフィルは日数分の時間がかかる
- **ポート番号を変える場合**: `main.py`を`uvicorn main:app --port XXXX`で起動したら、
  各`.bas`の`API_URL`定数も同じポートに変更すること
- **Excelファイルがロックされてエラーになる場合**: `.xlsx`/`.xlsm`をExcelで開いたまま
  Python側から書き込もうとすると`PermissionError`になる。Excelを閉じてから実行する
- **`swap_history.db`はgit管理外**: `history_store.py`が接続時に自動でテーブルを作成するため、
  ファイルが無い状態から始めても問題ない
