import WaterlooRegion.spreadsheet_f as spreadsheet


path = r'C:\Users\raymond.mu\Documents\GitHub\EngConditionAssessment\WaterlooRegion'
excel_name = 'WaterData.xlsx'
sheet_name = 'sheet1'
#
# spreadsheet.combine_excels(path, excel_name, sheet_name)

df = spreadsheet.get_oxl_sheet(excel_name, sheet_name)
spreadsheet.update_sheet(df)