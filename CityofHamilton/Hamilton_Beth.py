import pyautogui as pyg
import pyperclip
import time
import pandas as pd
import xlrd
import xlsxwriter
import timeit
import numpy as np
import ConditionAssessment_Functions.CAFunctions as CA


start = timeit.default_timer()

# Ask user to input which facility, excel sheet to work on, planning period and starting year
# facility_name = input('Type Facility Name: ')
excel_filename = 'BethComments'
excel_sheetname = 'ESL'
space = ' '
PlanningPeriod = 25
AssessmentYear = 2020

excel_filename = excel_filename + '.xlsx'
AssetExcelFile = pd.ExcelFile(excel_filename)
df = AssetExcelFile.parse(excel_sheetname)
df_cleaned = CA.Database_Cleaning(df)
df_cleaned = df_cleaned.set_index('AssetID')

for index, AssetInput in df_cleaned.iterrows():
    category = AssetInput['AssetCategory']
    ESL = CA.Calculator_Hamilton_AssetServiceLife(category)
    df_cleaned.at[index, 'VisualCondition'] = ESL[0]

with pd.ExcelWriter('Beth2.xlsx', engine = 'xlsxwriter') as writer:
    df_cleaned.to_excel(writer, sheet_name= 'Sheet1')