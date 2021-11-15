from openpyxl import load_workbook
import pandas as pd
from itertools import islice
import WaterlooRegion.system_f as system
import CAFunctions.CAFunctions as ca

def get_pd_sheet(filename, sheetname):
    excel_name = filename
    excel_file = pd.ExcelFile(excel_name)
    df = excel_file.parse(sheetname)
    return df


def get_oxl_sheet(filename, sheet_name):
    excel_name = filename
    workbook = load_workbook(filename=excel_name)
    worksheet = workbook[sheet_name]
    data = worksheet.values
    cols = next(data)[0:]
    data = list(data)
    data = (islice(r, 0, None) for r in data)
    df = pd.DataFrame(data, columns=cols).set_index('Barcode')
    return df


def create_one_excel(excel_name, tab, content):
    excel_name = excel_name + '.xlsx'
    with pd.ExcelWriter(excel_name, engine='xlsxwriter') as writer:
        content.to_excel(writer, sheet_name=tab)

def combine_excels(path, workbook_name, tab_name):
    list_excel = system.get_excelname(path)
    list_data = []

    for name_excel in list_excel:
        name_sheet = name_excel.replace('.xlsx', '')
        df = get_oxl_sheet(name_excel, name_sheet)
        list_data.append(df)

    data_mega = pd.concat(list_data, sort=True)
    create_one_excel(workbook_name, tab_name, data_mega)



def update_sheet(df):
    # Factors for smooth function running: Asset Category, Asset Condition Rating
    # Asset installation year

    for index, asset_input in df.iterrows():

        ''' Extract spreadsheet information '''
        AssetIndex_Input = index
        AssetDescription_Input = asset_input['Asset Description']
        AssetBuilding_Input = asset_input['Building Name']
        AssetCategory_Input = asset_input['Category']
        AssetConditionRating_Input = asset_input['Condition Rating']
        AssetInspectDate_Input = asset_input['Date']
        AssetInspectionCom_Input = asset_input['Inspector Comments']
        AssetLocation_Input = asset_input['Location']
        AssetPhyLocation_Input = asset_input['Physical Location']
        AssetInstallYear_Input = asset_input['Year of Installation']


        '''Condition(s) will allow for a pass'''
        # note: if asset category is missing, pass
        if AssetCategory_Input is None:
            continue

        if AssetConditionRating_Input is None:
            continue


        '''Installation Year Processing'''
        # note: if the installation is empty, it will not process.
        # output: AssetInstallYear_Output
        if AssetInstallYear_Input is not None:
            AssetInstallYear_Input = str(AssetInstallYear_Input)
            AssetInstallYear_Output = AssetInstallYear_Input.split('-')[0]

        '''Code Violation and Concern'''
        AssetCodeViolation_tup = system.get_codeconcern(AssetInspectionCom_Input)
        AssetCodeBoolean_Output = AssetCodeViolation_tup[0]
        AssetCodeComment_Output = AssetCodeViolation_tup[1]

        '''Condition Processing'''
        # AG: Physical Condition Rating
        AssetConditionRating_Input = AssetConditionRating_Input.lower()
        AssetConditionRating_AG_Output = ca.Converter_UniAssetConditionConversion(AssetConditionRating_Input)
        AssetConditionSentence_Mid = 'The asset was in {} condition.'.format(AssetConditionRating_Input)

        '''Condition Sentence Processing'''


        '''Performance Rating'''
        # AI: Performance Condition Rating
        # AJ: Performance Rating Comment
        if AssetConditionRating_AG_Output == 1 or 2 or 3:
            AssetPerformanceRating_Output = AssetConditionRating_AG_Output
            AssetPerformanceSentence_Output = 'The asset did not have any performance issues.'
        if AssetConditionRating_AG_Output == 4 or 5 and 'performance' not in AssetInspectionCom_Input:
            AssetPerformanceRating_Output = 3
            AssetPerformanceSentence_Output = 'The asset did not have any performance issues.'
        else:
            AssetPerformanceRating_Output = 4
            AssetPerformanceSentence_Output = 'The asset had performance issue.'


        #
        # '''Inspection sentence processing'''
        # AssetInspectionCom_Input = AssetInspectionCom_Input.split('\n')




# filename = r'C:\Users\raymond.mu\Documents\GitHub\EngConditionAssessment\WaterlooRegion\Data\WaterData.xlsx'
# sheet_name = 'sheet1'
# df = get_oxl_sheet(filename,sheet_name)
# update_sheet(df)