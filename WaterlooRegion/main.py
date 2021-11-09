import CAFunctions.CAFunctions as ca
import WaterlooRegion.spreadsheet_f as ss


df = ss.get_oxl_sheet('01 Preston - PPE','01 Preston - PPE')
ss.create_one_excel('Testing','tab',df)