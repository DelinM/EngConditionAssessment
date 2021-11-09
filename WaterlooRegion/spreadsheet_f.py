from openpyxl import load_workbook
import pandas as pd

def get_pd_sheet(filename,sheetname):
    excel_name = filename + '.xlsx'
    excel_file = pd.ExcelFile(excel_name)
    df = excel_file.parse(sheetname)
    return df

def get_oxl_sheet(filename, sheetname):
    excel_name = filename + '.xlsx'
    workbook = load_workbook(filename=excel_name)
    worksheet = workbook[sheetname]
    data = worksheet.values

    cols = next(data)[1:]
    index = [r[0] for r in data]






def create_one_excel(excel_name, tab, content):
    excel_name = excel_name + '.xlsx'
    with pd.ExcelWriter(excel_name,engine='xlsxwriter') as writer:
        content.to_excel(writer,sheet_name=tab)
