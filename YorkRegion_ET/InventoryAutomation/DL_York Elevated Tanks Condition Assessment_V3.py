
import pyperclip
import time
import pandas as pd
import xlrd
import xlsxwriter
import timeit
import numpy as np

# Import
import CAFunctions.CAFunctions as CA


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
facility_name = 'North Richmond Hill Coons Road ET'
file_name = 'RawData\_NRH_RAW'

file_name += '.xlsx'
sheet_name = 'Sheet1'
outputname_spreadsheet = 'NRH_Processed.xlsx'

# Set program parameters
Remove_Keyword = '(Removed)'
PlanningPeriod = 20
AssessmentYear = 2021

# Create dataframe to store data
assetexcel_file = pd.ExcelFile(file_name)
df = assetexcel_file.parse(sheet_name)
df_Cleaned = CA.Database_Cleaning(df)
df_Cleaned = df_Cleaned.set_index('seqNum')

# For loop to analyze asset one by one
for index, AssetInput in df_Cleaned.iterrows():
    # Extract information from dataframe
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

    # Make sure that the Asset Tag and Location Tag has 8 digits, overwrite input data
    AssetTAG_Output = CA.converter_FillinNumber(AssetTAG_Input, 8)
    AssetLocation_Output = CA.converter_FillinNumber(AssetLocation_Input, 8)
    df_Cleaned.at[index, 'assetTag'] = AssetTAG_Output
    df_Cleaned.at[index, 'LocationTag'] = AssetLocation_Output

    # Remove asset 'installationDate' information
    df_Cleaned.at[index, 'installationDate'] = ''

    # Continue if key words appeared
    AssetInput_Split_Internal = AssetName_Input.split(' ')
    if AssetInput_Split_Internal[0] == '(Removed)' or AssetInput_Split_Internal[0] == '(New)' or \
            AssetInput_Split_Internal[0] == '(Missing)':
        continue

    # Calculate asset condition adn asset risk
    AssetCondition_int_Output = int(
        CA.Analysis_ConditionAssessment_ObservationBased(AssetDefect1_Input, AssetDefect2_Input, AssetDefect3_Input,
                                                      AssetCategory_Input))
    AssetRisk_int_internal = CA.Calculator_AssetRisk(AssetCondition_int_Output, AssetCOF_Input)

    # Generate asset defect sentence
    AssetDefect1_Output = CA.SentenceGenerator_AssetObservation(AssetDefect1_Input, AssetCategory_Input)
    AssetDefect2_Output = CA.SentenceGenerator_AssetObservation(AssetDefect2_Input, AssetCategory_Input)
    AssetDefect3_Output = CA.SentenceGenerator_AssetObservation(AssetDefect3_Input, AssetCategory_Input)

    # Calculate asset estimated remaining service life
    AssetAge_Internal = CA.Calculator_AssetAge(AssessmentYear, AssetInstallYear_Input)
    AssetESL_Internal = CA.Calculator_YorkRegion_AssetServiceLife(AssetCategory_Input)
    AssetRealRL_Internal = CA.Calculator_AssetRealRemainingLife(AssetESL_Internal, AssetAge_Internal)

    AssetERSL_Output =CA.Analysis_AssetEstimatedRemainingServiceLife(PlanningPeriod, AssetRealRL_Internal,
                                                                   AssetESL_Internal
                                                                   , AssetCondition_int_Output, AssetRisk_int_internal)

    # Generate rehabilitation sentence
    AssetRehbab_Internal = CA.Analysis_AssetRehabTiming(AssetName_Input, AssetCondition_int_Output, AssetESL_Internal,
                                                     AssetERSL_Output, PlanningPeriod, AssetRisk_int_internal)
    AssetRehabSentence_Internal = AssetRehbab_Internal[0]
    AssetRehabTiming1_Output = CA.Converter_TimingtoYear(AssetRehbab_Internal[1], AssessmentYear)
    AssetRehabTiming2_Output = CA.Converter_TimingtoYear(AssetRehbab_Internal[2], AssessmentYear)

    if AssetRecommendation1_Input == '':
        AssetRehabSentence_Internal = AssetRecommendation1_Input + AssetRehabSentence_Internal
    elif AssetRecommendation1_Input != '':
        AssetRehabSentence_Internal = AssetRecommendation1_Input + ' ' + AssetRehabSentence_Internal

    # Generate condition adn observation sentence
    AssetConditionSentence_Internal = CA.SentenceGenerator_AssetCondition('asset',
                                                                       CA.Converter_AssetConditionConversion(
                                                                           AssetCondition_int_Output),
                                                                       AssetDefect1_Input, AssetDefect2_Input,
                                                                       AssetDefect3_Input)

    AssetObservationsandRec_Output = CA.SentenceGenetator_ObservationandRecommendation(AssetConditionSentence_Internal,
                                                                                    AssetRehabSentence_Internal,
                                                                                    'asset', AssetERSL_Output)



    # Capital Project
    if 0 <= AssetERSL_Output <= 10:
        Capital_Project_Output = facility_name + ' ' + 'Capital Project 1'
    elif 10 < AssetERSL_Output <= 20:
        Capital_Project_Output = facility_name + ' ' + 'Capital Project 2'
    else:
        Capital_Project_Output = ''

    # Rehab Project
    Mid_Time = int(AssessmentYear + PlanningPeriod / 2)
    Final_Time = int(AssessmentYear + PlanningPeriod)
    print(type(Mid_Time))
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

    if AssetRehabTiming2_Output == '':
        RehabProject2_Output = ''
    elif AssessmentYear <= AssetRehabTiming2_Output <= Mid_Time:
        RehabProject2_Output = facility_name + ' ' + 'Capital Project 1'
    elif Mid_Time < AssetRehabTiming2_Output <= Final_Time:
        RehabProject2_Output = facility_name + ' ' + 'Capital Project 2'
    else:
        RehabProject2_Output = ''


    # Update dataframe result
    # AssetInput['condition'] = AssetCondition_int_Output
    # AssetInput['defect1Comment'] = AssetDefect1_Output
    # AssetInput['defect2Comment'] = AssetDefect2_Output
    # AssetInput['defect3Comment'] = AssetDefect3_Output
    # AssetInput['remainingESL'] = AssetERSL_Output
    # AssetInput['observations'] = AssetObservationsandRec_Output
    # AssetInput['RehabRepairYear'] = AssetRehabTiming1_Output
    # AssetInput['RehabRepairYear2'] = AssetRehabTiming2_Output

    df_Cleaned.at[index, 'condition'] = AssetCondition_int_Output
    df_Cleaned.at[index, 'defect1'] = AssetDefect1_Output
    df_Cleaned.at[index, 'defect2'] = AssetDefect2_Output
    df_Cleaned.at[index, 'defect3'] = AssetDefect3_Output
    df_Cleaned.at[index, 'remainingESL'] = AssetERSL_Output
    df_Cleaned.at[index, 'observations'] = AssetObservationsandRec_Output
    df_Cleaned.at[index, 'RehabRepairYear'] = AssetRehabTiming1_Output
    df_Cleaned.at[index, 'RehabRepairYear2'] = AssetRehabTiming2_Output
    df_Cleaned.at[index, 'repProjectName'] = Capital_Project_Output
    df_Cleaned.at[index, 'rehabProjectName'] = RehabProject1_Output
    df_Cleaned.at[index, 'rehab2ProjectName'] = RehabProject2_Output
    df_Cleaned.at[index, 'RehabRepairCost'] = AssetRehab1Cost_Output
    df_Cleaned.at[index, 'RehabRepairCost2'] = AssetRehab2Cost_Output
    df_Cleaned.at[index, 'RehabRepairCost2'] = AssetRehab2Cost_Output

    print(index)
    print('Asset Name', AssetInput['assetName'])
    print('Asset Condition', df_Cleaned.at[index, 'condition'])
    print('Asset Risk', AssetRisk_int_internal)
    print('ARRL: ', AssetRealRL_Internal)
    print('Asset ESL:', df_Cleaned.at[index, 'remainingESL'])
    print('Asset rehab sentence:', df_Cleaned.at[index, 'observations'])
    print('Asset rehab timing 1:', df_Cleaned.at[index, 'RehabRepairYear'])
    print('Asset rehab timing 2:', df_Cleaned.at[index, 'RehabRepairYear2'])
    print('Asset Tag Original', AssetTAG_Input)
    print('Asset Tag Original', type(AssetTAG_Input))
    print('Asset Tag:', AssetTAG_Output)
    print(type(AssetInput))
    print('\n')




writer = pd.ExcelWriter(outputname_spreadsheet, engine='xlsxwriter')
df_Cleaned.to_excel(writer, sheet_name='Sheet1')

writer.save()

stop = timeit.default_timer()

print('Number of Assets assessed:', index)
print('Program Time used: ', stop - start)
