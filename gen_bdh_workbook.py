"""
Bloomberg BDH関数でJPY TONA OISスワップのヒストリカルレートを取得する
土台ワークブックを生成する。

レイアウトは旧プロジェクト(excel-fastapi/data/gen_mock_data.py)の規約を踏襲する:
  A列は未使用、B列がDate(1本のみ)、C列以降にテナーごとのレート。
  1行目=ティッカー名、2行目=Date/テナーラベル、3行目からデータ。

ダミー値をPythonで生成する代わりに、このスクリプトはセルにBDH数式そのものを
書き込む。実行後にExcel(Bloombergアドイン有効)で開いて再計算すれば、
実データが入る。

BDH数式の組み方:
  - 先頭テナー(ON)は通常のBDH(日付サプレスなし)にする。
    Excel365の動的配列で Date|Rate の2列(B:C)にスピルする。
  - 以降のテナー(1W, 1M, ...)は同じ日付範囲・カレンダーのはずなので、
    オーバーライド "Dts=S" で日付列を抑制し、レート1列だけをB:Cと同じ行に
    スピルさせる(D, E, F, ... 列)。

ティッカーは実在のBloombergコードではなく空欄にしてある。
Controlシートの表に、実際に使うBloombergティッカーを入力してから
Excelで開くこと。
"""
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUT_PATH = Path(__file__).parent / "JPY_TONA_SwapHistory.xlsx"
CONTROL_SHEET = "Control"
DATA_SHEET = "JPY_TONA_Hist"

TENORS = ["ON", "1W", "1M", "3M", "6M", "1Y", "2Y", "3Y", "5Y", "7Y", "10Y", "15Y", "20Y", "30Y"]

# Controlシート上のセル配置
START_DATE_CELL = "B2"
END_DATE_CELL = "B3"
FIELD_CELL = "B4"
PERIODICITY_CELL = "B5"
TICKER_TABLE_START_ROW = 8  # ヘッダ行。テナー行はこの次(9)から


def build_control_sheet(wb: Workbook) -> dict:
    """Controlシートを作り、テナー→ティッカーセルのアドレス辞書を返す。"""
    ws = wb.active
    ws.title = CONTROL_SHEET

    label_font = Font(bold=True, size=10)
    header_font = Font(bold=True, size=9, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4472C4")
    note_font = Font(size=9, italic=True, color="808080")

    ws["A1"] = "BDH取得設定"
    ws["A1"].font = Font(bold=True, size=12)

    ws["A2"] = "開始日"
    ws["A2"].font = label_font
    ws[START_DATE_CELL] = date.today() - timedelta(days=365 * 2)
    ws[START_DATE_CELL].number_format = "yyyy-mm-dd"

    ws["A3"] = "終了日"
    ws["A3"].font = label_font
    ws[END_DATE_CELL] = "=TODAY()"
    ws[END_DATE_CELL].number_format = "yyyy-mm-dd"

    ws["A4"] = "価格フィールド"
    ws["A4"].font = label_font
    ws[FIELD_CELL] = "PX_LAST"

    ws["A5"] = "周期(Per)"
    ws["A5"].font = label_font
    ws[PERIODICITY_CELL] = "D"
    ws["C5"] = "D=日次 W=週次 M=月次"
    ws["C5"].font = note_font

    ws["A7"] = "※ 下表のティッカー列に実際のBloombergティッカーを入力してください(現在は空欄)"
    ws["A7"].font = note_font

    header_row = TICKER_TABLE_START_ROW
    ws.cell(row=header_row, column=1, value="テナー").font = header_font
    ws.cell(row=header_row, column=1).fill = header_fill
    ws.cell(row=header_row, column=2, value="ティッカー").font = header_font
    ws.cell(row=header_row, column=2).fill = header_fill

    ticker_cells = {}
    for i, tenor in enumerate(TENORS):
        r = header_row + 1 + i
        ws.cell(row=r, column=1, value=tenor)
        ws.cell(row=r, column=2, value="")  # ここに実ティッカーを入力
        ticker_cells[tenor] = f"B{r}"

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 24

    return ticker_cells


def build_data_sheet(wb: Workbook, ticker_cells: dict):
    ws = wb.create_sheet(DATA_SHEET)

    ticker_font = Font(bold=True, size=10)
    subhdr_font = Font(bold=True, size=9, color="555555")
    ticker_fill = PatternFill("solid", fgColor="E7ECEF")
    date_fmt = "yyyy-mm-dd"
    rate_fmt = "0.000"

    DATE_COL = 2  # B

    date_hdr = ws.cell(row=2, column=DATE_COL, value="Date")
    date_hdr.font = subhdr_font
    ws.column_dimensions[get_column_letter(DATE_COL)].width = 12

    ctrl = f"'{CONTROL_SHEET}'!"
    start_ref = f"{ctrl}${START_DATE_CELL[0]}${START_DATE_CELL[1:]}"
    end_ref = f"{ctrl}${END_DATE_CELL[0]}${END_DATE_CELL[1:]}"
    field_ref = f"{ctrl}${FIELD_CELL[0]}${FIELD_CELL[1:]}"
    per_ref = f"{ctrl}${PERIODICITY_CELL[0]}${PERIODICITY_CELL[1:]}"

    for i, tenor in enumerate(TENORS):
        rate_col = DATE_COL + 1 + i  # C, D, E, ...
        col_letter = get_column_letter(rate_col)
        ticker_ref = f"{ctrl}${ticker_cells[tenor][0]}${ticker_cells[tenor][1:]}"

        ticker_cell = ws.cell(row=1, column=rate_col, value=f"={ticker_ref}")
        ticker_cell.font = ticker_font
        ticker_cell.fill = ticker_fill
        ticker_cell.alignment = Alignment(horizontal="center")

        label_cell = ws.cell(row=2, column=rate_col, value=tenor)
        label_cell.font = subhdr_font
        label_cell.alignment = Alignment(horizontal="center")

        ws.column_dimensions[col_letter].width = 10

        if i == 0:
            # 先頭テナー: 通常のBDH。Date|Rateの2列(B:C)にスピルする。
            formula = f'=BDH({ticker_ref},{field_ref},{start_ref},{end_ref},"Per="&{per_ref})'
            ws.cell(row=3, column=DATE_COL, value=formula)
        else:
            # 以降のテナー: Dts=Sで日付列を抑制し、レート1列だけをB列と同じ行にスピルさせる。
            formula = f'=BDH({ticker_ref},{field_ref},{start_ref},{end_ref},"Per="&{per_ref},"Dts=S")'
            ws.cell(row=3, column=rate_col, value=formula)

    ws.freeze_panes = "C3"
    ws.sheet_view.showGridLines = False


def main():
    wb = Workbook()
    ticker_cells = build_control_sheet(wb)
    build_data_sheet(wb, ticker_cells)
    wb.save(OUT_PATH)
    print(f"saved: {OUT_PATH}")
    print("次の手順:")
    print(f"  1. '{CONTROL_SHEET}'シートのティッカー列に実際のBloombergティッカーを入力")
    print("  2. Excel(Bloombergアドイン有効な端末)でファイルを開く")
    print("  3. 再計算(F9 / Ctrl+Alt+F9)してBDHの結果を反映させる")


if __name__ == "__main__":
    main()
