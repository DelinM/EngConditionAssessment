import pyautogui as pyg
import pyperclip
import time
import pandas as pd
import xlrd
import xlsxwriter
import timeit
import numpy as np
import ConditionAssessment_Functions.CAFunctions as CA



# The purpose for Sudbury P



start = timeit.default_timer()

# Ask user to input which facility, excel sheet to work on, planning period and starting year
# facility_name = input('Type Facility Name: ')
excel_filename = 'Raw_SudburyPhase1_Levack'
excel_sheetname = 'Sheet1'
space = ' '


# Access information in the excel spreadsheet, and set up a dataframe for access.
excel_filename = excel_filename + '.xlsx'
AssetExcelFile = pd.ExcelFile(excel_filename)
df = AssetExcelFile.parse(excel_sheetname)
df_cleaned = CA.Database_Cleaning(df)
df_cleaned = df_cleaned.set_index('AssetID')

''' iterate each assets from the asset inventory to get asset information '''
for index, AssetInput in df_cleaned.iterrows():
    AssetIndex_Input = index
    #AssetProcess_Input = AssetInput['TempProcess']
    AssetCategory_Input = AssetInput['AssetCategory']
    #AssetSubProcess_Input = AssetInput['TempSubProcess']
    #AssetSubprocessInstance_Input = str(AssetInput['TempSubProcessInstance'])
    AssetName_Input = AssetInput['AssetName']
    AssetDescription_Input = AssetInput['AssetDescription']
    AssetFacilityName_Input = AssetInput['FacilityName']
    AssetLocation_Input = AssetInput['LocationName']

    ''' Produce description, manufacturer, model, serial number, size '''
    AssetDescription_Interal = CA.Analysis_Sudbury_AssetDescription(AssetDescription_Input)
    asset_description_output = AssetDescription_Interal[0]
    asset_manufacturer_output = AssetDescription_Interal[1]
    asset_model_output = AssetDescription_Interal[2]
    asset_serno_output = AssetDescription_Interal[3]
    asset_size_output = AssetDescription_Interal[4]

    df_cleaned.at[index, 'AssetDescription'] = asset_description_output
    df_cleaned.at[index, 'Manufacturer'] = asset_manufacturer_output
    df_cleaned.at[index, 'Model'] = asset_model_output
    df_cleaned.at[index, 'SerialNumber'] = asset_serno_output
    df_cleaned.at[index, 'SizeCapacity'] = asset_size_output

    ''' If the asset name contains (Removed), then skip the entire analysis '''
    AssetInput_Split_Internal = AssetName_Input.split(' ')
    if AssetInput_Split_Internal[0] == '(Removed)' or AssetInput_Split_Internal[0] == '(New)' or \
            AssetInput_Split_Internal[0] == '(Missing)':
        continue

    if AssetName_Input == '':
        continue

    ''' Code below to produce the correct asset name '''
    if AssetCategory_Input == 'Building Structural' or AssetCategory_Input == 'Building Architectural' or \
        AssetCategory_Input == 'Building Electrical' or AssetCategory_Input =='Building Mechanical':
        print('AssetLocation_Input:',AssetLocation_Input)
        if AssetFacilityName_Input == AssetLocation_Input:
            AssetName_Output = AssetLocation_Input + space + AssetName_Input
            # print('asset name:', AssetName_Output)
        else:
            AssetName_Output = AssetFacilityName_Input + space + AssetLocation_Input + space + AssetName_Input
            # print('asset name:',AssetName_Output)
    else:
        AssetName_Output = AssetName_Input
    # As per instruction, the renaming is not required for the process related assets.
    '''elif AssetCategory_Input == 'Process Structural':
        if AssetSubProcess_Input == '':
            AssetName_Output = AssetProcess_Input + space + AssetName_Input
        else:
            AssetSubProcessList_Internal = AssetSubProcess_Input.split()
            AssetNameList_Internal = AssetName_Input.split()
            if AssetSubProcessList_Internal[0] == AssetNameList_Internal[0]:
                AssetName_Output = AssetName_Input.lstrip(' ')
            else:
                if len(AssetSubProcessList_Internal) == 1:
                    AssetName_Output = AssetProcess_Input + space + AssetName_Input
                else:
                    AssetName_Output = AssetSubProcess_Input + space + AssetName_Input
    elif AssetCategory_Input == 'Process Instrumentation' or 'Process Electrical':
        AssetName_Output = AssetName_Input
    elif AssetCategory_Input == 'Process Mechanical':
        if AssetSubprocessInstance_Input != '':
            AssetSubprocessInstance_Input = AssetSubprocessInstance_Input[0]

        if AssetProcess_Input == 'Raw Sewage Pumping':
            AssetName_Output = AssetProcess_Input + space + AssetSubprocessInstance_Input + space + AssetName_Input
        elif AssetProcess_Input != 'Raw Sewage Pumping':
            AssetSubProcessList_Internal = AssetSubProcess_Input.split()
            AssetNameList_Internal = AssetName_Input.split()
            if AssetSubProcessList_Internal[0] == AssetNameList_Internal[0]:
                name = ''
                for string in AssetNameList_Internal[0:]:
                    name += string
                AssetName_Output = AssetSubProcess_Input + space + AssetSubprocessInstance_Input + space + name
            else:
                AssetName_Output = AssetSubProcess_Input + space + AssetSubprocessInstance_Input + space + \
                                   AssetName_Input '''
    print('Asset Index:', index)
    print('Asset Name:', AssetName_Output)

    print('\n')

    '''update the asset name for exporting purpose'''
    df_cleaned.at[index, 'AssetName'] = AssetName_Output


'''
writer is used to overwrite exisitng data
'''

writer = pd.ExcelWriter('Processed_SudburyPhase1_Levack1.xlsx', engine = 'xlsxwriter')
df_cleaned.to_excel(writer, sheet_name = 'Sheet1')

writer.save()

stop = timeit.default_timer()

time = stop - start

print('The total time used:', time)