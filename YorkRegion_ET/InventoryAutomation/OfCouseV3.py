import pyautogui as pyg
import pyperclip
import time
import pandas as pd
import xlrd
import xlsxwriter
import timeit
import numpy as np

# Import
import ConditionAssessment_Functions.CAFunctions as CA


'''
Version: Version 0.3
Date: 05/27/2020
Update Person: Delin Mu
Update Note: 1. modification of code to make it clear
             2. some revision on CAFunctions
'''

# Call for time function to record program running time
start = timeit.default_timer()

# User input information

file_name = 'Waterloo.xlsx'
sheet_name = 'Sheet1'
outputname_spreadsheet = '20200209_SuttonET.xlsx'

# Set program parameters
Remove_Keyword = '(Removed)'
PlanningPeriod = 10
AssessmentYear = 2021

# Create dataframe to store data
assetexcel_file = pd.ExcelFile(file_name)
df = assetexcel_file.parse(sheet_name)
df_Cleaned = CA.Database_Cleaning(df)
df_Cleaned = df_Cleaned.set_index('AssetID')

# For loop to analyze asset one by one
for index, AssetInput in df_Cleaned.iterrows():
    # Extract information from dataframe

    AssetInstallYear_Input = AssetInput['InstallYear']
    AssetCOF_Input = AssetInput['CoF']
    AssetPoF_Iput = AssetInput['PoF']
    AssetESL_Input = AssetInput['AvgESL']

    AssetAge_Internal = 2021 - AssetInstallYear_Input
    # Calculate asset estimated remaining service life

    AssetRisk_int_internal = AssetPoF_Iput * AssetCOF_Input
    AssetRealRL_Internal = CA.Calculator_AssetRealRemainingLife(AssetESL_Input, AssetAge_Internal)

    AssetERSL_Output =CA.Analysis_AssetEstimatedRemainingServiceLife(PlanningPeriod, AssetRealRL_Internal,
                                                                   AssetESL_Input
                                                                   , AssetPoF_Iput, AssetRisk_int_internal)



    df_Cleaned.at[index, 'remainingESL'] = AssetERSL_Output





writer = pd.ExcelWriter(outputname_spreadsheet, engine='xlsxwriter')
df_Cleaned.to_excel(writer, sheet_name='Sheet1')

writer.save()

stop = timeit.default_timer()

print('Number of Assets assessed:', index)
print('Program Time used: ', stop - start)
