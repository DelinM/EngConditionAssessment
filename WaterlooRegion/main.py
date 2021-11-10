import WaterlooRegion.spreadsheet_f as ss


path = r'/Users/RayDelinmu/Documents' \
       r'/EngConditionAssessment/WaterlooRegion'
excel_name = 'WaterData'
sheet_name = 'sheet1'

ss.combine_excels(path, excel_name, sheet_name)