### readWriteXl.py
'''
セルの座標 (x, y) は 1 始まりで指定します。
openpyxl で読み書きできるのは .xlsx 形式のファイルのみです。
'''
# module
from openpyxl import load_workbook

import openpyxl

def excel_copy_sort_search(input_file="aaa.xlsx", output_file="bbb.xlsx", search_val="1234"):
    """
    Excelファイルのコピー、ソート、検索を行う関数

    Args:
        input_file (str): 入力Excelファイル名
        output_file (str): 出力Excelファイル名
        search_val (str): 検索する値
    """

    try:
        # Excelファイルの読み込み
        wb = openpyxl.load_workbook(input_file)
        ws = wb["test1"]

        # 新しいワークシートを作成し、データをコピー
        ws2 = wb.create_sheet(title="test2")
        for row in ws.iter_rows(values_only=True):
            ws2.append(row)

        # データのソート
        sort_data(ws2)

        # ソート結果を保存
        wb.save(output_file)

        # 検索
        col_no = 2  # B列
        n = 2  # 2行目から検索開始
        col, line = search(ws2, col_no, search_val, n)

        if col:
            print(f"値 '{search_val}' が行 {line} に見つかりました。")
            print(f"列データ: {col}")
        else:
            print(f"値 '{search_val}' は見つかりませんでした。")

    except FileNotFoundError:
        print(f"ファイル '{input_file}' が見つかりません。")
    except KeyError:
        print("指定されたシートが見つかりません。")
    except Exception as e:
        print(f"エラーが発生しました: {e}")


def sort_data(ws):
    """
    ワークシートのデータをソートする関数

    Args:
        ws (openpyxl.worksheet.worksheet.Worksheet): ワークシートオブジェクト
    """

    data = list(ws.iter_rows(min_row=2))  # ヘッダーを除いたデータを取得
    data.sort(key=lambda row: (row[1].value, row[9].value))  # B列と10列目でソート

    # ソートされたデータをワークシートに書き込む
    for i, row in enumerate(data, start=2):
        for j, cell in enumerate(row):
            ws.cell(row=i, column=j + 1, value=cell.value)


def search(ws, col_no, val, n):
    """
    ワークシート内で指定された値を検索する関数

    Args:
        ws (openpyxl.worksheet.worksheet.Worksheet): ワークシートオブジェクト
        col_no (int): 検索する列番号
        val (str): 検索する値
        n (int): 検索を開始する行番号

    Returns:
        tuple: 見つかった場合は列データのリストと行番号、見つからなかった場合はNone, None
    """

    for row in ws.iter_rows(min_row=n):
        if str(row[col_no - 1].value) == val:
            col = [cell.value for cell in row]
            return col, row[0].row
    return None, None


if __name__ == "__main__":
    excel_copy_sort_search()
exit
#
# ======================================================================================
#  Excel ファイルの読み込み

def openXl(fn, sheet_name="JROD"):
    """
    # Excel ファイルを開く
    Args:
        fn: file name
        sheet_name: sheet name

    Returns:
        workbook,
        worksheet,
        title: ヘッダーのカラム名 {1:'title1', 2:'title2, ... 'max_row':21, 'max_column':114}

    """
    workbook = load_workbook(fn)
    sheet = workbook[sheet_name]  # シート名を指定
    # info
    print(f"<openXl> fn: {fn}, sheet: {sheet_name}, max row={sheet.max_row}, col={sheet.max_column}")
    
    # セル (x, y) の値を取得 (x, y は 1 始まり)
    if 0:
        x = 2  # 例: 2 行目
        y = 3  # 例: 3 列目
        for y in range(1,10):
            print(f'{sheet.cell(row=1, column=y).value:16}:', end="")
            cell_value = sheet.cell(row=x, column=y).value
            print(f"セル({x}, {y}) の値: {cell_value}")

    # title set
    title = {}
    for y in range(1, sheet.max_column+1):
        title[y] = sheet.cell(row=1, column=y).value
    print(f"<openXl> title 1:{title[1]}, title max:{title[sheet.max_column]}")
    title['max_row'] = sheet.max_row
    title['max_column'] = sheet.max_column

    return workbook, sheet, title

    # すべてのセルを読み込む (ジェネレータ)
    for row in sheet.iter_rows():
        for cell in row:
            print(cell.value)

# ==============================================================
# Excel ファイルへの書き込み
'''
def writeXl():
    # Excel ファイルを開く
    workbook = load_workbook('your_excel_file.xlsx')

    # シートを取得
    sheet = workbook['Sheet1']

    # セル (x, y) に値を書き込む (x, y は 1 始まり)
    x = 2  # 例: 2 行目
    y = 3  # 例: 3 列目
    sheet.cell(row=x, column=y).value = '新しい値'

    # Excel ファイルを保存
    workbook.save('your_excel_file.xlsx')
'''
# ================================================================
# 特定のセル範囲の読み書き
'''
def readWrite():

    # Excel ファイルを開く
    workbook = load_workbook('your_excel_file.xlsx')

    # シートを取得
    sheet = workbook['Sheet1']

    # セル範囲 (x1, y1) から (x2, y2) の値を読み込む
    x1 = 2
    y1 = 3
    x2 = 4
    y2 = 5
    for row in sheet.iter_rows(min_row=x1, min_col=y1, max_row=x2, max_col=y2):
        for cell in row:
            print(cell.value)

    # セル範囲 (x1, y1) から (x2, y2) に値を書き込む
    x1 = 2
    y1 = 3
    x2 = 4
    y2 = 5
    values = [['a', 'b', 'c'], ['d', 'e', 'f']]  # 書き込む値のリスト
    for i, row in enumerate(sheet.iter_rows(min_row=x1, min_col=y1, max_row=x2, max_col=y2)):
        for j, cell in enumerate(row):
            cell.value = values[i][j]

    # Excel ファイルを保存
    workbook.save('your_excel_file.xlsx')
'''
# ================
# wsの指定した行のレコードを取得する。
def getRow(ws, x):
    """
    ワークシートの指定した行のすべての列の値を取得する関数

    Args:
        ws: openpyxlのワークシートオブジェクト
        x: 取得する行番号 (1始まり)

    Returns:
        指定した行のすべての列の値を含むリスト. d[0] = row_no
        行が存在しない場合は空のリストを返す
    """
    if x > ws.max_row: x = ws.max_row
    row_values = [x]
    if 1 <= x <= ws.max_row:  # 行が存在するか確認
        for cell in ws[x]:  # 指定した行のすべてのセルを反復処理
            if cell.value is None:
                v = ""
            else: v = cell.value
            row_values.append(v)  # セルの値をリストに追加
    return row_values

def setRow(ws, x, data):
    """
    1行分のデータを指定した行のワークシートに書き込む関数

    Args:
        ws: openpyxlのワークシートオブジェクト
        x: 書き込む行番号 (1始まり)
        data: 書き込むデータを含むリスト
    """
    if data[0] != x:
        print(f"ERROR: list[0] mismatch x:{x}, list[0]={data[0]}")
        return
    if 1 <= x <= ws.max_row + 1:  # 行が存在するか、または次の行に書き込むか確認
        for y, value in enumerate(data):  # データを列方向に書き込む
            if y == 0: continue
            ws.cell(row=x, column=y).value = value  # セルに値を書き込む
    else:
        print(f"Error: Row {x} is beyond the worksheet boundaries.")  # エラーメッセージを表示
    


if __name__ == "__main1__":
    print("このスクリプトは直接実行されました")
    fn = 'JRODe_SRC_Sample.xlsx'
    print(f'excel filename: {fn}')
    wb, ws, title = openXl(fn)
    #print(title)
    
    rows = getRow(ws, 1)
    print("\nROW:1, データを表示(y:1 ... 10)")
    for i in range(1, 10):
        print(f'{i}, {title[i]:10}: {rows[i]}')
        
    #print("rows=", rows)    
    