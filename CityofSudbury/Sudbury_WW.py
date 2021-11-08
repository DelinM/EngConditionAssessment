import pandas as pd
import timeit

import ConditionAssessment_Functions as CA

start = timeit.default_timer()

# Ask user to input which facility, excel sheet to work on, planning period and starting year
# facility_name = input('Type Facility Name: ')
excel_filename = 'sudbury_ww'
excel_sheetname = 'Sheet1'
PlanningPeriod = 30
AssessmentYear = 2021

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

# iterate each assets from the asset inventory to get asset information
for index, AssetInput in df_cleaned.iterrows():
    AssetIndex_Input = index
    AssetCategory_Input = AssetInput['AssetCategory']
    AssetName_Input = AssetInput['AssetName']
    AssetDescription_Input = AssetInput['AssetDescription']
    SiteName_Input = AssetInput['SiteName']
    AssetFacilityName_Input = AssetInput['FacilityName']
    AssetInstallYear_Input = AssetInput['InstallYear']
    AssetCoF_Input = AssetInput['CoF']


    # receive asset comments
    AssetComment1_Input = AssetInput['TempComments1']
    AssetComment2_Input = AssetInput['TempComments2']
    AssetComment3_Input = AssetInput['TempComments3']

    # if asset name is missing, there is no need to go through this analysis
    if AssetName_Input == '':
        continue

    # Convert lifecycle category
    AssetCategory_Internal = CA.Calculator_Sudbury_AssetServiceLife(AssetCategory_Input)[1]

    # Generate Assessment Year
    AssetAssetYear_Output = AssessmentYear

    # Generate asset condition and PoF
    AssetCondition_Output = CA.Analysis_ConditionAssessment_ObservationBased(AssetComment1_Input, AssetComment2_Input,
                                                                          AssetComment2_Input,AssetCategory_Internal)
    AssetPoF_Output = AssetCondition_Output

    # Generate asset average ESL
    AssetESL_Output = CA.Calculator_Sudbury_AssetServiceLife(AssetCategory_Input)[0]

    # Generate asset real remaining service life (ARRL)
    if AssetInstallYear_Input == '':
        AssetInstallYear_Input = 1990
    AssetAge_Internal = CA.Calculator_AssetAge(AssessmentYear, AssetInstallYear_Input)
    ARRL_Internal = CA.Calculator_AssetRealRemainingLife(AssetESL_Output, AssetAge_Internal)

    # Generate asset risk
    if AssetCoF_Input == '':
        AssetCoF_Input = 0
    AssetRisk_Internal = AssetPoF_Output * AssetCoF_Input

    # Generate Asset remaining service life

    AssetERSL_Internal = CA.Analysis_AssetEstimatedRemainingServiceLife(PlanningPeriod, ARRL_Internal,
                                                                        AssetESL_Output, AssetCondition_Output,
                                                                        AssetRisk_Internal)



    # Section below is to produce the result for tblComments
    # Generate asset condition sentence
    AssetConditionWord_Output = CA.Converter_AssetConditionConversion(AssetCondition_Output)
    AssetConditionSentence_Internal = CA.SentenceGenerator_AssetCondition('asset', AssetConditionWord_Output,
                                                                          AssetComment1_Input,AssetComment2_Input,
                                                                          AssetComment3_Input)
    # Generate asset observation sentence
    Assetobs1_Internal = CA.SentenceGenerator_AssetObservation(AssetComment1_Input, AssetCategory_Internal)
    Assetobs2_Internal = CA.SentenceGenerator_AssetObservation(AssetComment2_Input, AssetCategory_Internal)
    Assetobs3_Internal = CA.SentenceGenerator_AssetObservation(AssetComment3_Input, AssetCategory_Internal)
    AssetOBSentencelist_internal = CA.SentenceGenerator_Sudbury_AssetObservationSummary(AssetConditionSentence_Internal,
                                                                                         Assetobs1_Internal,
                                                                                         Assetobs2_Internal,
                                                                                         Assetobs3_Internal)
    AssetFullObservation_Internal = AssetOBSentencelist_internal[0]
    AssetCC_Internal = AssetOBSentencelist_internal[1]
    AssetHS_Internal = AssetOBSentencelist_internal[2]
    AssetOM_Internal = AssetOBSentencelist_internal[3]

    # code below to append asset comments to Table tblcomment
    commenttypelist = ['Condition Comment', 'Code Concern Comment', 'H&S Concern Comment', 'Client Comment']
    i = 0
    while i < 4:
        if AssetFullObservation_Internal != '':
            tblcomment.append([index, AssetName_Input,commenttypelist[i],AssetOBSentencelist_internal[i]])
        i += 1


    # Generate replacement information
    Assetrepyear_Internal = CA.Analysis_ReplacementYears(AssetERSL_Internal, AssetESL_Output, AssessmentYear,
                                                       PlanningPeriod)
    tblcapitalworks.append([index,AssetName_Input, Assetrepyear_Internal[0], 'True', AssetESL_Output])

    print('Asset ID:',index)
    print('Asset Name:', AssetName_Input)
    # print('Asset Condition:', AssetCondition_Output)
    # print('condition comments:', AssetFullObservation_Internal)
    # print('code concern:',AssetCC_Internal)
    # print('Health and safety:', AssetHS_Internal)
    # print('OM:', AssetOM_Internal)
    print(tbldefects)
    print('\n\n')

    # Code below to override data into the df_cleaned dataframe
    df_cleaned.at[index, 'VisualCondition'] = AssetCondition_Output
    df_cleaned.at[index, 'PoF'] = AssetPoF_Output
    df_cleaned.at[index, 'AvgESL'] = AssetESL_Output


# print(tblcomment)
# print(tblcapitalworks)


# convert list to dataframe
df_tblcomment = pd.DataFrame(tblcomment)
df_tblcapitalworks = pd.DataFrame(tblcapitalworks)
df_tbldefects = pd.DataFrame(tbldefects)


# export dataframe to a new excel spreadsheet
with pd.ExcelWriter('Sudbury Phase 2 Result_RecommendationUpdate.xlsx', engine = 'xlsxwriter') as writer:
    df_cleaned.to_excel(writer, sheet_name= 'Sheet1')
    df_tblcomment.to_excel(writer, sheet_name= 'comment')
    df_tblcapitalworks.to_excel(writer, sheet_name= 'replacement')
    df_tbldefects.to_excel(writer, sheet_name= 'rehab')



stop = timeit.default_timer()

time = stop - start

print('The total time used:', time)