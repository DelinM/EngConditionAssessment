import pyautogui as pyg
import pyperclip
import time
import pandas as pd
import xlrd
import xlsxwriter
import timeit
import numpy as np
'''
When use the code, copy paste the python inventory keywords file under the folder.
use:

import CA.Functions to solve 
'''
from ConditionAssessment_Functions.CAFunctions import Calculator_YorkRegion_AssetServiceLife
from ConditionAssessment_Functions.CAFunctions import Calculator_AssetAge
from ConditionAssessment_Functions.CAFunctions import Calculator_AssetRealRemainingLife
from ConditionAssessment_Functions.CAFunctions import Calculator_AssetRisk
from ConditionAssessment_Functions.CAFunctions import Converter_AssetConditionConversion
from ConditionAssessment_Functions.CAFunctions import Converter_UniAssetConditionConversion
from ConditionAssessment_Functions.CAFunctions import Converter_TimingtoYear
from ConditionAssessment_Functions.CAFunctions import Analysis_ConditionAssessment_AgeBased_Conservative
from ConditionAssessment_Functions.CAFunctions import Analysis_ConditionAssessment_AgedBased_LinearRegression
from ConditionAssessment_Functions.CAFunctions import Analysis_ConditionAssessment_ObservationBased
from ConditionAssessment_Functions.CAFunctions import Analysis_AssetEstimatedRemainingServiceLife
from ConditionAssessment_Functions.CAFunctions import Analysis_AssetRehabTiming
from ConditionAssessment_Functions.CAFunctions import SentenceGenerator_AssetCondition
from ConditionAssessment_Functions.CAFunctions import SentenceGenerator_AssetObservation
from ConditionAssessment_Functions.CAFunctions import SentenceGenerator_AssetReplacement
from ConditionAssessment_Functions.CAFunctions import SentenceGenetator_ObservationandRecommendation
from ConditionAssessment_Functions.CAFunctions import Database_Cleaning
from ConditionAssessment_Functions.CAFunctions import Databbase_Cleaning_RemoveZero
from ConditionAssessment_Functions.CAFunctions import converter_FillinNumber


start = timeit.default_timer()
facility_name = input("Type Facility Name: ")

'''
Whole function Testing
'''
Remove_Keyword = '(Removed)'
fileName = 'tbl_assets.xlsx'
Sheet_Name = 'tbl_assets'
PlanningPeriod = 20
AssessmentYear = 2019

AssetExcelFile = pd.ExcelFile(fileName)

df = AssetExcelFile.parse(Sheet_Name)
df_Cleaned = Database_Cleaning(df)
df_Cleaned = df_Cleaned.set_index('seqNum')

df_final = pd.DataFrame()

for index,AssetInput in df_Cleaned.iterrows():

    '''use the following for York, code below to extract all useful information from the inventory spreadsheet'''
    AssetName_Input = AssetInput['assetName']
    AssetCategory_Input = AssetInput['CategoryCode']
    AssetInstallYear_Input = AssetInput['installYear']
    AssetCOF_Input = AssetInput['COF']
    AssetDefect1_Input = AssetInput['defect1Input']
    AssetDefect2_Input = AssetInput['defect2Input']
    AssetDefect3_Input = AssetInput['defect3Input']
    AssetRecommendation1_Input = AssetInput['rehabComment1']
    AssetRecommendation2_Input = AssetInput['rehabComment2']
    AssetRehab1Cost_Output = AssetInput['rehabCost1']
    AssetRehab2Cost_Output = AssetInput['rehabCost2']
    AssetTAG_Input = AssetInput['assetTag']
    AssetLocation_Input = AssetInput['LocationTag']

    '''Convert Asset Tag number and Location number'''
    AssetTAG_Output = converter_FillinNumber(AssetTAG_Input,8)
    AssetLocation_Output = converter_FillinNumber(AssetLocation_Input,8)
    df_Cleaned.at[index,'assetTag'] = AssetTAG_Output
    df_Cleaned.at[index,'LocationTag'] = AssetLocation_Output

    '''
    remove the asset 'installationDate' information
    '''
    df_Cleaned.at[index, 'installationDate'] = ''

    '''
    if the asset name contains (Removed), then skip the entire analysis
    '''
    AssetInput_Split_Internal = AssetName_Input.split(' ')
    if AssetInput_Split_Internal[0] == '(Removed)' or AssetInput_Split_Internal[0] == '(New)' or AssetInput_Split_Internal[0] == '(Missing)':
        continue


    AssetCondition_int_Output = int(Analysis_ConditionAssessment_ObservationBased(AssetDefect1_Input,AssetDefect2_Input,AssetDefect3_Input,
                                                                                AssetCategory_Input))
    AssetRisk_int_internal = Calculator_AssetRisk(AssetCondition_int_Output,AssetCOF_Input)

    '''
    Form asset defects sentences
    '''
    AssetDefect1_Output = SentenceGenerator_AssetObservation(AssetDefect1_Input,AssetCategory_Input)
    AssetDefect2_Output = SentenceGenerator_AssetObservation(AssetDefect2_Input,AssetCategory_Input)
    AssetDefect3_Output = SentenceGenerator_AssetObservation(AssetDefect3_Input,AssetCategory_Input)

    '''
    calculate asset estimated remaining service life
    '''
    AssetAge_Internal = Calculator_AssetAge(AssessmentYear,AssetInstallYear_Input)
    AssetESL_Internal = Calculator_YorkRegion_AssetServiceLife(AssetCategory_Input)
    print('AssetESL_Internal', AssetCategory_Input)
    AssetRealRL_Internal = Calculator_AssetRealRemainingLife(AssetESL_Internal,AssetAge_Internal)

    AssetERSL_Output = Analysis_AssetEstimatedRemainingServiceLife(PlanningPeriod,AssetRealRL_Internal,AssetESL_Internal
                                                                   ,AssetCondition_int_Output,AssetRisk_int_internal)

    '''
    Asset Rehabiltation.

    '''
    AssetRehbab_Internal = Analysis_AssetRehabTiming(AssetName_Input,AssetCondition_int_Output,AssetESL_Internal,AssetERSL_Output, PlanningPeriod, AssetRisk_int_internal)
    AssetRehabSentence_Internal = AssetRehbab_Internal[0]
    AssetRehabTiming1_Output = Converter_TimingtoYear(AssetRehbab_Internal[1],AssessmentYear)
    AssetRehabTiming2_Output = Converter_TimingtoYear(AssetRehbab_Internal[2],AssessmentYear)

    if AssetRecommendation1_Input == '':
        AssetRehabSentence_Internal = AssetRecommendation1_Input + AssetRehabSentence_Internal
    elif AssetRecommendation1_Input != '':
        AssetRehabSentence_Internal = AssetRecommendation1_Input + ' ' +AssetRehabSentence_Internal

    '''
    observation sentence
    '''
    AssetConditionSentence_Internal = SentenceGenerator_AssetCondition(AssetName_Input,Converter_AssetConditionConversion(AssetCondition_int_Output),
                                                                      AssetDefect1_Input,AssetDefect2_Input,AssetDefect3_Input)

    AssetObservationsandRec_Output = SentenceGenetator_ObservationandRecommendation(AssetConditionSentence_Internal, AssetRehabSentence_Internal,
                                                                                    AssetName_Input, AssetERSL_Output)

    '''
    determine which project
    
    '''

    '''Capital Project'''
    if 0 <= AssetERSL_Output <= 10:
        Capital_Project_Output = facility_name + ' ' + 'Capital Project 1'
    elif 10 < AssetERSL_Output <= 20:
        Capital_Project_Output = facility_name + ' ' + 'Capital Project 2'
    else:
        Capital_Project_Output = ''

    '''Rehab Project'''
    Mid_Time = int(AssessmentYear + PlanningPeriod/2)
    Final_Time = int(AssessmentYear + PlanningPeriod)
    print (type(Mid_Time))
    # AssetRehabTiming1_Output = int(AssetRehabTiming1_Output)
    # AssetRehabTiming2_Output = int(AssetRehabTiming2_Output)
    if AssetRehabTiming1_Output == '':
        RehabProject1_Output = ''
    elif AssessmentYear <= AssetRehabTiming1_Output <= Mid_Time:
        RehabProject1_Output = facility_name + ' ' + 'Capital Project 1'
    elif Mid_Time < AssetRehabTiming1_Output <= Final_Time:
        RehabProject1_Output = facility_name + ' ' + 'Capital Project 2'
    else:
        RehabProject1_Output = ''

    if AssetRehabTiming2_Output =='':
        RehabProject2_Output = ''
    elif AssessmentYear <= AssetRehabTiming2_Output <= Mid_Time:
        RehabProject2_Output = facility_name + ' ' + 'Capital Project 1'
    elif Mid_Time < AssetRehabTiming2_Output <= Final_Time:
        RehabProject2_Output = facility_name + ' ' + 'Capital Project 2'
    else:
        RehabProject2_Output = ''


    '''
    data_Output reassigned to data_Input
    '''
    # AssetInput['condition'] = AssetCondition_int_Output
    # AssetInput['defect1Comment'] = AssetDefect1_Output
    # AssetInput['defect2Comment'] = AssetDefect2_Output
    # AssetInput['defect3Comment'] = AssetDefect3_Output
    # AssetInput['remainingESL'] = AssetERSL_Output
    # AssetInput['observations'] = AssetObservationsandRec_Output
    # AssetInput['RehabRepairYear'] = AssetRehabTiming1_Output
    # AssetInput['RehabRepairYear2'] = AssetRehabTiming2_Output

    df_Cleaned.at[index,'condition'] = AssetCondition_int_Output
    df_Cleaned.at[index,'defect1'] = AssetDefect1_Output
    df_Cleaned.at[index,'defect2'] = AssetDefect2_Output
    df_Cleaned.at[index,'defect3'] = AssetDefect3_Output
    df_Cleaned.at[index,'remainingESL'] = AssetERSL_Output
    df_Cleaned.at[index,'observations'] = AssetObservationsandRec_Output
    df_Cleaned.at[index,'RehabRepairYear'] = AssetRehabTiming1_Output
    df_Cleaned.at[index,'RehabRepairYear2'] = AssetRehabTiming2_Output
    df_Cleaned.at[index,'repProjectName'] = Capital_Project_Output
    df_Cleaned.at[index,'rehabProjectName'] = RehabProject1_Output
    df_Cleaned.at[index,'rehab2ProjectName'] = RehabProject2_Output
    df_Cleaned.at[index,'RehabRepairCost'] = AssetRehab1Cost_Output
    df_Cleaned.at[index,'RehabRepairCost2'] = AssetRehab2Cost_Output
    df_Cleaned.at[index,'RehabRepairCost2'] = AssetRehab2Cost_Output




    print(index)
    print('Asset Name', AssetInput['assetName'])
    print('Asset Condition', df_Cleaned.at[index,'condition'])
    print('Asset Risk',AssetRisk_int_internal)
    print('ARRL: ', AssetRealRL_Internal)
    print('Asset ESL:',  df_Cleaned.at[index,'remainingESL'])
    print('Asset rehab sentence:', df_Cleaned.at[index,'observations'])
    print('Asset rehab timing 1:', df_Cleaned.at[index,'RehabRepairYear'])
    print('Asset rehab timing 2:', df_Cleaned.at[index,'RehabRepairYear2'])
    print('Asset Tag Original', AssetTAG_Input)
    print('Asset Tag Original', type(AssetTAG_Input))
    print('Asset Tag:', AssetTAG_Output)
    print(type(AssetInput))
    df_final.append(AssetInput, ignore_index=True)


    print('\n')

'''

writer is used to overwrite exisitng data

'''

writer = pd.ExcelWriter('testing5.xlsx', engine = 'xlsxwriter')
df_Cleaned.to_excel(writer, sheet_name = 'Sheet1')

writer.save()

stop = timeit.default_timer()

print ('Number of Assets assessed:', index)
print('Program Time used: ', stop - start)
