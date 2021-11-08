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
excel_filename = 'Phase2'
excel_sheetname = 'Sheet1'
space = ' '
PlanningPeriod = 30
AssessmentYear = 2020

# Access information in the excel spreadsheet, and set up a dataframe for access.
excel_filename = excel_filename + '.xlsx'
AssetExcelFile = pd.ExcelFile(excel_filename)
df = AssetExcelFile.parse(excel_sheetname)
df_cleaned = CA.Database_Cleaning(df)
df_cleaned = df_cleaned.set_index('AssetID')

# List to store asset comments for tblComments
tblcomment = [['ParentID', 'AssetName', 'Comment Type', 'Comment']]

# List to store asset replacement information for tblCapitalWorks
tblcapitalworks = [['ParentID', 'AssetName', 'WorkYear', 'Recurring', 'WorkFrequency']]

# List to store asset rehab information for tblDefects
tbldefects = [['ParentID', 'AssetName', 'RehabComment', 'Rehabyear', 'RehabUnitMaterialCost']]

# Import tblCapitalWorks table
excel_sheetname_capital = 'tblCapitalWorks'
df_capitalworks = AssetExcelFile.parse(excel_sheetname_capital)
df_capitalworks_cleaned = CA.Database_Cleaning(df_capitalworks)
df_capitalworks_cleaned = df_capitalworks_cleaned.set_index('WorkID')

# Update main asset spreadsheet using tblCapitalWorks
for index, AssetInput in df_cleaned.iterrows():
    print(df_capitalworks[index])








