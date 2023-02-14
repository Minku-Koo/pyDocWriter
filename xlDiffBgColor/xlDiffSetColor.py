import xlwings as xw
import pandas as pd
from pyxlsb import open_workbook as open_xlsb
import openpyxl as xl
from openpyxl.styles import PatternFill

'''
def _open_excel_wlwings(filepath, sheetname):
    # 엑셀 인스턴스 생성
    app = xw.App(visible=True)
    # 파일 상장법인목록
    book = xw.Book(filepath)
    # 첫번째 시트 읽어오기
    sheet = book.sheets[sheetname]
    # 데이터프레임 형태로 엑셀 시트 읽어오기
    # df = sheet.options(index=False, expand='table').value
    # 인스턴스 종료
    app.kill()
    return book, sheet

def get_df_pandas(filepath, sheetname):
    df = []

    with open_xlsb(filepath) as wb:
        with wb.get_sheet(sheetname) as sheet:
            for row in sheet.rows():
                df.append([item.v for item in row])

    df = pd.DataFrame(df[1:], columns=df[0])
    return df

def _compare_df(df1, df2):
    result = []
    for x, _ in enumerate(df1.iterrows()):
        for y, _ in enumerate(df1.iloc[x]):
            if df1.iloc[x][y] == df2.iloc[x][y]:
                continue
            print(y, x)
            result.append((y, x))
    return result

if __name__ == "__main__":
    s = "./sample/sample.xlsb"
    t = "./sample/sample2.xlsb"

    sheet_name = 'sh2'

    df_s = get_df_pandas(s, sheet_name)
    df_t = get_df_pandas(t, sheet_name)

    diff_tuple = _compare_df(df_s, df_t)

    color = '#ff0000'

    t_book, t_sheet = _open_excel_wlwings(t, sheet_name)

    for y, x in diff_tuple:
        t_sheet.range(y, x).color = color
'''

def _open_excel_openpyxl(filepath, sheetname):
    # Load workbook
    book = xl.load_workbook(filepath)
    # Access to a worksheet named 'no_header'
    sheet = book[sheetname]
    return book, sheet

def get_df_pandas(filepath, sheetname):
    # Load workbook
    book = xl.load_workbook(filepath)
    # Access to a worksheet named 'no_header'
    sheet = book[sheetname]
    # Convert to DataFrame
    df = pd.DataFrame(sheet.values)
    return df

def _compare_df(df1, df2):
    result = []
    for x, _ in enumerate(df1.iterrows()):
        for y, _ in enumerate(df1.iloc[x]):
            if df1.iloc[x][y] == df2.iloc[x][y]:
                continue
            print(y, x)
            result.append((y, x))
    return result

if __name__ == "__main__":
    s = "./sample/1.xlsx"
    t = "./sample/3.xlsx"

    sheet_name = 'sh2'

    sb, ss = _open_excel_openpyxl(s, sheet_name)
    tb, ts = _open_excel_openpyxl(t, sheet_name)

    

    diff_tuple = _compare_df(pd.DataFrame(ss.values), pd.DataFrame(ts.values))

    print(pd.DataFrame(ts.values))

    color = 'FAAC58'
    bfill = PatternFill(start_color=color, end_color=color, fill_type='solid')

    for y, x in diff_tuple:
        ts.cell(row=y, column=x).fill = bfill

    # tb.template = False
    tb.save(t)