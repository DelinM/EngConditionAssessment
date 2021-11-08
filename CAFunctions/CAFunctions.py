import pyperclip

import time
import pandas as pd
import xlrd
import xlsxwriter
import timeit
import numpy as np

'''
Function Description:
First Line: Introduction: what does the function do?
Second Line: Input parameters
Third Line: Output parameters
Fourth Line: Note   1
Fifth Line: Testing Status.
'''


def Calculator_YorkRegion_AssetServiceLife(LifeCycleCategory):
    '''
    Introduction: Function to return theoretical service life based on Input 'Life Cycle Category'
    Input: York Region Life Cycle Category (string)
    Output: Asset Service Life (int)
    Note: If the input parameter cannot find a match in the function, the function will return 30.
    (Testing Status: Pass)
    '''

    if LifeCycleCategory == 'Architectural Components' or LifeCycleCategory == 'AC':
        return 15
    elif LifeCycleCategory == 'Building Mechanical' or LifeCycleCategory == 'BM':
        return 30
    elif LifeCycleCategory == 'Electrical System' or LifeCycleCategory == 'ES':
        return 25
    elif LifeCycleCategory == 'Health and Safety System' or LifeCycleCategory == 'HSS':
        return 15
    elif LifeCycleCategory == 'Process Mechanical' or LifeCycleCategory == 'PM':
        return 30
    elif LifeCycleCategory == 'Structural Components' or LifeCycleCategory == 'SC':
        return 60
    elif LifeCycleCategory == 'SCADA/Instrumentation/Control' or LifeCycleCategory == 'SIC':
        return 15
    elif LifeCycleCategory == 'Site Works' or LifeCycleCategory == 'SW':
        return 20
    else:
        return 30

def Calculator_Hamilton_AssetServiceLife(LifeCycleCategory):

    if LifeCycleCategory == 'Building Structural' or LifeCycleCategory ==\
            'Process Structural' or LifeCycleCategory == 'BS' or LifeCycleCategory == 'PS':
        return [60, 'SC']
    elif LifeCycleCategory == 'Building Architectural' or LifeCycleCategory == 'BA':
        return [30, 'AC']
    elif LifeCycleCategory == 'Building Electrical' or LifeCycleCategory == 'BE' or \
            LifeCycleCategory == 'Process Electrical' or LifeCycleCategory == 'PE':
        return [25, 'ES']
    elif LifeCycleCategory == 'Building Mechanical' or LifeCycleCategory == 'BM':
        return [30, 'BM']
    elif LifeCycleCategory == 'Process Mechanical' or LifeCycleCategory == 'PM':
        return [30, 'PM']
    elif LifeCycleCategory == 'Process Instrumentation' or LifeCycleCategory == 'PI':
        return [15, 'SIC']
    elif LifeCycleCategory == 'Site Works' or LifeCycleCategory == 'SW':
        return [20, 'SW']
    else:
        return [30, 'AC']

def Calculator_Halton_AssetServiceLife(LifeCycleCategory):

    if LifeCycleCategory == 'Building and Process Structural':
        return [60, 'SC']
    elif LifeCycleCategory == 'Building Architectural':
        return [30, 'AC']
    elif LifeCycleCategory == 'Process Electrical' or LifeCycleCategory == 'PE':
        return [25, 'ES']
    elif LifeCycleCategory == 'Building Services' or LifeCycleCategory == 'BS':
        return [30, 'BM']
    elif LifeCycleCategory == 'Process Mechanical Equipment' or LifeCycleCategory == 'PM':
        return [30, 'PM']
    elif LifeCycleCategory == 'Process Piping' or LifeCycleCategory == 'PI':
        return [60, 'PM']
    elif LifeCycleCategory == 'Process Instrumentation' or LifeCycleCategory == 'SIC':
        return [15, 'SIC']
    elif LifeCycleCategory == 'Site Works' or LifeCycleCategory == 'SW':
        return [20, 'SW']

def Calculator_Sudbury_AssetServiceLife(LifeCycleCategory):
    '''
    :param LifeCycleCategory:
    :return:
    '''
    if LifeCycleCategory == 'Building Structural' or LifeCycleCategory ==\
            'Process Structural' or LifeCycleCategory == 'BS' or LifeCycleCategory == 'PS':
        return [60, 'SC']
    elif LifeCycleCategory == 'Building Architectural' or LifeCycleCategory == 'BA':
        return [30, 'AC']
    elif LifeCycleCategory == 'Building Electrical' or LifeCycleCategory == 'BE' or \
            LifeCycleCategory == 'Process Electrical' or LifeCycleCategory == 'PE':
        return [25, 'ES']
    elif LifeCycleCategory == 'Building Mechanical' or LifeCycleCategory == 'BM':
        return [30, 'BM']
    elif LifeCycleCategory == 'Process Piping and Equipment' or LifeCycleCategory == 'PPE':
        return [30, 'PM']
    elif LifeCycleCategory == 'Process Instrumentation' or LifeCycleCategory == 'PI':
        return [15, 'SIC']
    elif LifeCycleCategory == 'Site Works' or LifeCycleCategory == 'SW':
        return [20, 'SW']
    else:
        return [30, 'AC']

def Calculator_AssetAge(AssessmentYear, InstallationYear):
    '''
    Introduction: Function to calculate asset age
    Input: AssessmentYear (int), InstallationYear (int)
    Output: Asset Age`
    Note: None
    (Testing Status: Pass)
    '''
    a = InstallationYear
    if a == None:
        a = 0
    age = AssessmentYear - a
    # Inputs:condition assessment year, asset installation year
    return AssessmentYear - InstallationYear

def Calculator_AssetRealRemainingLife(AssetESL, AssetAge):
    # Function to calculate asset real remaining service life
    # Inputs: asset expected service year, asset age
    Real_RemainingLife = AssetESL - AssetAge
    if Real_RemainingLife >= 0:
        return Real_RemainingLife
    if Real_RemainingLife < 0:
        return 0

def Calculator_AssetRisk(AssetCondition, AssetPofF):
    # Function to calculate asset risk
    # input: asset condition rating, asset probability of failure
    # Note: For York Region, the risk will be used to determine the timing of the replacement project,
    # lets say if the calculate risk value is 3.1, then the risk will be roll up to 4.
    asset_risk = AssetCondition * AssetPofF
    asset_risk_int = int(AssetCondition * AssetPofF)
    difference = asset_risk - asset_risk_int
    if difference > 0.2:
        asset_risk_int = asset_risk_int + 1
    return asset_risk_int

def Calculator_AssetReplacementYear(project_year, assetestimatedremaininglife):
    return project_year + assetestimatedremaininglife

def Converter_AssetConditionConversion(Number_AssetConditionGrade):
    # Converter to convert a asset condition number grading to a description grading.
    # Input: Number Asset Condition Grade
    # Note: The converter works for most of the Regions, including York Region, Halton Region,etc..
    if 0 <= Number_AssetConditionGrade < 1.5:
        return 'Very Good'
    elif 1.5 <= Number_AssetConditionGrade < 2.5:
        return 'Good'
    elif 2.5 <= Number_AssetConditionGrade < 3.5:
        return 'Fair'
    elif 3.5 <= Number_AssetConditionGrade < 4.5:
        return 'Poor'
    elif 4.5 <= Number_AssetConditionGrade <= 5.0:
        return 'Very Poor'

def Converter_StringtoNumber(string):
    number = 0
    if string == '':
        return number
    else:
        return float(string)

def Converter_UniAssetConditionConversion(Word_AssetConditionGrading):
    # Function to make sure the condition grading is an integer among: 1, 2, 3 ,4 ,5
    # Input: the description of the asset condition, i.e. "fair"
    # Note: for reporting purpose,  this rating system works for York Region, Halton Region, and Hamilton
    grading = Word_AssetConditionGrading.lower()
    assetcondition = int()

    if grading == 'very good':
        assetcondition = 1
        return assetcondition
    elif grading == 'good':
        assetcondition = 2
        return assetcondition
    elif grading == 'fair':
        assetcondition = 3
        return assetcondition
    elif grading == 'poor':
        assetcondition = 4
        return assetcondition
    elif grading == 'very poor':
        assetcondition = 5
        return assetcondition

def Converter_TimingtoYear(timing, assessmentyear):
    if timing == '':
        return ''
    else:
        return (timing + assessmentyear)

def Analysis_Sudbury_AssetDescription(assetdescription):
    '''
    :param assetdescription:
    :return: a list that contains asset description, manufacturer, model, serial number, size/capacity
    '''

    assetdescription_list = assetdescription.split(';')
    assetdescription_listlen = len(assetdescription_list)

    if assetdescription_listlen < 5:
        num_empty = 5 - assetdescription_listlen
        iter = 0
    else:
        num_empty = 5

    for iter in range(num_empty):
        assetdescription_list.append('')

    return assetdescription_list

def Analysis_Sudbury_AssetRehabTiming(AssetName, Asset_Condition, AssetESL, PlanningPeriod, Asset_Risk,
                                      Asset_rehabcoment1, Asset_rehabyear1, conditionassessmentyear):
    '''Function Description: Function to determine the timing(s) for asset rehabilitation and generate rehabilitaion
    sentence if applicable.
    Inputs: Asset Condition (int), AssetESL_Estimated Service Life (int), AssetERSL_estimated remaining service life (int)
    Asset Planning Planning Period (int), Asset Risk (int)
    Outputs: tuple (Rehabilitation Sentence, Rehabilitation 1 Timing, Rehabilitation 2 Timing)
    Note:
    (Testing Status:)
    '''

    # identify all output parameters.
    rehab_sentence1 = ''
    rehab_sentence2 = ''
    rehab_sentence3 = ''

    rehab_timing1 = int()
    rehab_timing2 = int()
    rehab_timing3 = int()

    rehab_sentence = ' Provision for structural repairs every 10 years.'
    rehab_defect_sentence = 'Structural repairs to address asset defects.'

    if Asset_rehabcoment1 != '':
        rehab_sentence1 = Asset_rehabcoment1
        rehab_timing1 = Asset_rehabyear1

    AssetName = AssetName.lower()
    AssetName = list(AssetName.split(' '))
    key_words = ['wall', 'concrete', 'masonry', 'footing', 'Cell', 'Chamber', 'tank', 'channel', 'roof']
    i = 0
    for name in AssetName:
        for key in key_words:
            if key == name:
                i = 1

    if i == 1:
        if rehab_sentence1 == '':
            if Asset_Condition < 2.5 and AssetESL >= 50:
                rehab_sentence1 = rehab_sentence
                rehab_sentence2 = rehab_sentence
                rehab_sentence3 = rehab_sentence
                rehab_timing1 = conditionassessmentyear + 10
                rehab_timing2 = rehab_timing1 + 10
                rehab_timing3 = rehab_timing2 + 10
            elif Asset_Condition >= 2.5 and AssetESL >= 50:
                rehab_sentence1 = rehab_defect_sentence
                rehab_sentence2 = rehab_sentence
                rehab_sentence3 = rehab_sentence
                if Asset_Risk >= 10:
                    rehab_timing1 = conditionassessmentyear + 4
                    rehab_timing2 = rehab_timing1 + 10
                    rehab_timing3 = rehab_timing2 + 10
                elif Asset_Risk < 10:
                    rehab_timing1 = conditionassessmentyear + 6
                    rehab_timing2 = rehab_timing1 + 10
                    rehab_timing3 = rehab_timing2 + 10
        elif rehab_sentence1 != '':
            rehab_sentence2 = rehab_sentence
            rehab_sentence3 = rehab_sentence
            rehab_timing2 = rehab_timing1 + 10
            rehab_timing3 = rehab_timing2 + 10

    i = 0

    result_loop = [rehab_timing1, rehab_timing2, rehab_timing3, rehab_sentence1, rehab_sentence2, rehab_sentence3]

    while i < 3:
        if result_loop[i] >= conditionassessmentyear + PlanningPeriod + 1 or result_loop[i] == 0:
            result_loop[i] = ''
        i += 1
    a = 3
    while a < 6:
        if result_loop[a-3] == '':
            result_loop[a] = ''
        a += 1

    return result_loop

def Analysis_ConditionAssessment_AgeBased_Conservative(AssetAge, ESL):
    # Function to do condition assessment when there's no information available
    # Input: asset age, asset expected service life or theoretical service life
    # Note: if inspectors do not see the asset, the worse condition will be assumed to be fair, as per "conservative"

    difference = ESL - AssetAge
    if difference <= 0:
        return 3.0
    elif difference > 0:
        ratio = AssetAge / ESL
        if 0 <= ratio <= 0.1:
            return 1.0
        elif 0.1 < ratio < 0.5:
            return 2.0
        else:
            return 3.0

def Analysis_ConditionAssessment_AgedBased_LinearRegression(AssetAge, ESL):
    # Function to do condition assessment when  there's no information available
    ratio = AssetAge / ESL * 5
    if ratio >= 1:
        return 1.0
    else:
        return ratio

def Analysis_ConditionAssessment_ObservationBased(observation1, observation2, observation3, lifecyclecategory):
    file = 'Python Inventory_Final.xlsx'
    data = pd.ExcelFile(file)
    df = data.parse(lifecyclecategory)
    df1 = df.set_index("Keyword")
    observation1_input = observation1.capitalize()
    observation2_input = observation2.capitalize()
    observation3_input = observation3.capitalize()

    # for first observation

    if observation1_input != '':
        i = 0
        a = len(observation1_input.split())
        testing1 = observation1_input
        for i in range(a):
            if testing1 in df1.Condition:
                condition1_output = float(df1.Condition[testing1])
                break
            condition1_output = 3.0
            testing1 = str(observation1_input.rsplit(' ', i)[0])
            i = i + 1
    elif observation1_input == '':
        condition1_output = 3.0
    else:
        condition1_output = 3.0

    # for second observation
    if observation2_input != '':
        o = 0
        b = len(observation2_input.split())
        testing2 = observation2_input
        for o in range(b):
            if testing2 in df1.Condition:
                condition2_output = float(df1.Condition[testing2])
                break
            condition2_output = 0
            testing2 = str(observation2_input.rsplit(' ', o)[0])
            o = o + 1
    elif observation2_input == '':
        condition2_output = 0
    else:
        condition2_output = 3.0

    # for third observation
    if observation3_input != '':
        u = 0
        c = len(observation3_input.split())
        testing3 = observation3_input
        for u in range(c):
            if testing3 in df1.Condition:
                condition3_output = float(df1.Condition[testing3])
                break
            condition3_output = 0
            testing3 = str(observation3_input.rsplit(' ', u)[0])
            u = u + 1
    elif observation3_input == '':
        condition3_output = 0
    else:
        condition3_output = 3.0

    condition = max([condition1_output, condition2_output, condition3_output])

    # Override Condition conditions
    # 1. if 'hs' or 'cc' (code concern) is included in the observation comments, then condition will be assigned to 4.
    # 2. if 'om' (operation & maintenance) is included in the observation comments, then condition will be assigne to 3.
    observation_list = [observation1_input, observation2_input, observation3_input]
    override_condition = []

    for ob_sentence in observation_list:
        ob_sentence = ob_sentence.lower()
        if 'no arc flash label' in ob_sentence:
            override_condition.append(3.0)
        elif 'hs' in ob_sentence or 'cc' in ob_sentence:
            override_condition.append(4.0)
        elif 'om' in ob_sentence:
            override_condition.append(3.0)
        else:
            override_condition.append(condition)

    #print('The condition override list is: ', override_condition)

    return max(override_condition)

def Analysis_AssetEstimatedRemainingServiceLife(PlanningScope, ARRL, AssetESL, num_AssetCondition, riskrating):
    '''
    Function to calculate Asset remaining life based on asset condition
    Inputs: the length of capital plan, ARRL(asset real remaining service life), AssetESL, asset condition grade
    Note: * this analysis determine the timing of the capital planning!!
    if asset is in very good or good condition,

    if the asset is in good condition, even the asset is old, the replacement of the asset will not happen inc das a
    short term. If the asset has a risk higher 8 (including 8) and the asset is not in a very good or good condition,
    the remaining service life will be estimated to be in the first 5 years.
    '''
    year_upbound = 6
    risk_upbound = 9
    risk_trigger = 14
    # print('ARRL Value:', ARRL)
    # print('Asset ESL', AssetESL)
    if num_AssetCondition <= 2.5:
        if AssetESL > 40:
            # If the asset is structural asset
            if ARRL >= PlanningScope * 1.25:
                return ARRL
            else:
                ARRL = PlanningScope + 2
            return ARRL
        if AssetESL <= 50:
            # If the asset is other than structural asset
            if ARRL <= year_upbound:
                ARRL = ARRL + year_upbound
                return ARRL
            else:
                return ARRL
    elif num_AssetCondition > 2.5:
        if AssetESL > 50:
            if ARRL >= PlanningScope * 1.25:
                return ARRL
            else:
                '''if the asset is structural asset '''
                ARRL = PlanningScope + 2
                return ARRL
        if AssetESL <= 50:
            # if the asset is other than structural asset
            if ARRL <= year_upbound:
                # if the asset is really needed to be replaced within a short term.
                if riskrating >= risk_upbound:
                    if riskrating > risk_trigger:
                        return 2
                    elif riskrating <= risk_trigger:
                        if ARRL == 0:
                            ARRL = ARRL + 1
                        if ARRL == 1:
                            ARRL = ARRL + 1
                            return ARRL
                        else:
                            return ARRL
                    # if the asset has a risk higher than 9 (including 9), meaning the asset actually has a high risk,
                    # then the asset is recommended to be replaced within 5 years.
                    return ARRL
                elif riskrating < risk_upbound:
                    ARRL = ARRL + year_upbound
                    return ARRL
            elif ARRL > year_upbound:
                return ARRL

def Analysis_AssetRehabTiming(AssetName, Asset_Condition, AssetESL, AssetERSL, PlanningPeriod, Asset_Risk):
    '''Function Description: Function to determine the timing(s) for asset rehabilitation and generate rehabilitaion
    sentence if applicable.
    Inputs: Asset Condition (int), AssetESL_Estimated Service Life (int), AssetERSL_estimated remaining service life (int)
    Asset Planning Planning Period (int), Asset Risk (int)
    Outputs: tuple (Rehabilitation Sentence, Rehabilitation 1 Timing, Rehabilitation 2 Timing)
    Note:
    (Testing Status:)
    '''

    rehab_sentence = ''
    rehab_Timing1 = ''
    rehab_Timing2 = ''

    AssetName = AssetName.lower()
    AssetName = list(AssetName.split(' '))
    key_words = ['wall', 'concrete']
    i = 0
    for name in AssetName:
        for key in key_words:
            if key == name:
                i = 1
    if i == 1:
        if Asset_Condition < 2.5 and AssetESL >= 50:
            rehab_sentence = ' Provision for structural repairs every 10 years.'
            rehab_Timing1 = int(PlanningPeriod / 2)
            return rehab_sentence, rehab_Timing1, rehab_Timing2
        elif Asset_Condition < 2.5 and AssetESL < 50:
            return rehab_sentence, rehab_Timing1, rehab_Timing2
        elif Asset_Condition >= 2.5:
            if AssetESL >= 50 and PlanningPeriod > 10:
                rehab_sentence = ' Provisions for structural repairs.'
                if Asset_Risk >= 10:
                    rehab_Timing1 = 2
                    rehab_Timing2 = 17
                elif Asset_Risk < 10:
                    rehab_Timing1 = 6
                    rehab_Timing2 = 17
                return rehab_sentence, rehab_Timing1, rehab_Timing2
            elif AssetESL >= 50 and PlanningPeriod <= 10:
                rehab_sentence = ' Provision for structural repairs.'
                if Asset_Risk >= 10:
                    rehab_Timing1 = 2
                if Asset_Risk < 9:
                    rehab_Timing1 = 2
                rehab_Timing2 = ''
                return rehab_sentence, rehab_Timing1, rehab_Timing2
            elif AssetESL < 50 and AssetERSL > 20:
                rehab_sentence = ''
            return rehab_sentence, rehab_Timing1, rehab_Timing2
    else:
        return rehab_sentence, rehab_Timing1, rehab_Timing2

def Analysis_ReplacementYears(AssetERSL,AssetESL, conditionassessmentyear, plannerperiod):
    '''
    :param AssetERSL:
    :return: replacement year in a type
    '''

    replacement_year1 = conditionassessmentyear + AssetERSL
    replacement_year2 = replacement_year1 + AssetESL
    replacement_year3 = replacement_year2 + AssetESL

    replacementyears_list = [replacement_year1, replacement_year2, replacement_year3]

    # for i, year in enumerate(replacementyears_list):
    #     if replacementyears_list[i] >= conditionassessmentyear + plannerperiod + 1:
    #         replacementyears_list[i] = ''

    return replacementyears_list

def SentenceGenerator_AssetCondition(assetname, assetcondition, observation1, observation2, observation3):
    assetcondition = assetcondition.lower()
    assetname = assetname.lower()
    InspectionType_visual = 'observed to be in'
    InspectionType_assume = 'assumed to be in'
    AssetConditionSentence = ''
    if observation1 == observation2 == observation3 == '':
        AssetConditionSentence = 'The {} was {} fair condition.'.format(assetname, InspectionType_assume)
    else:
        AssetConditionSentence = 'The {} was {} {} condition.'.format(assetname, InspectionType_visual, assetcondition)
    return AssetConditionSentence

def SentenceGenerator_AssetObservation(observation, LifeCycleCategory):
    # Function to take inspectors notes, and convert them into a sentence based on keywords match and keywords unmatch

    # standardize input string
    observation = observation.lower()
    observation = observation.capitalize()

    if 'Hs: ' in observation:
        return observation
    if 'Cc: ' in observation:
        return observation
    if 'Om: ' in observation:
        return observation


    if observation == '':
        return ''
    elif observation == 'Good':
        return ''
    elif observation == 'New':
        return ''
    elif observation == 'Fair':
        return ''

    # it's possible that inspector add '.' at the end of the sentence code below to remove '.' in observation
    observation = observation.rstrip('.')

    # import adj. inventory
    File_ObservationADJ = 'Python Inventory_Final.xlsx'
    data = pd.ExcelFile(File_ObservationADJ)
    df = data.parse('ADJ')
    df1 = df.set_index('Keyword')

    # import keyword inventory
    File_SearchKeywords = 'Python Inventory_Final.xlsx'
    dataK = pd.ExcelFile(File_SearchKeywords)
    dfK = dataK.parse(LifeCycleCategory)
    df2 = dfK.set_index('Keyword')

    # parameters for finding adj and the associated loop
    i = 0
    observation_length = len(observation.split())
    LoopTesting = observation

    for i in range(observation_length):
        if LoopTesting in df1.Condition:
            LoopTesting = LoopTesting.lower()
            Observation_Sentence = 'The asset was {}.'.format(LoopTesting)
            return Observation_Sentence
        # The following code is not necessary for this for loop, because if we expect an adj, then the 'observation'
        # must be one(1) word
        # LoopTesting = str(observation.rsplit(' ', i)[0])
        i = i + 1

    # parameters for finding keywords and the associated loop
    u = 0
    observation_list = observation.split()

    # The following for loop is to identify few things:
    # 1. if observation is keyword(s)
    # 2. how long is this keyword
    # 3. use the length of the keywords to determine length of defect location
    # 4. based on the information above, form a observation sentence

    # if observation is a keyword(s), return observation sentence
    if observation in df2.Condition:
        Observation_Sentence = '{} was observed.'.format(observation)
        return Observation_Sentence

    # determine how long is the keyword
    word_counter = 0
    keyword_counter = 0
    keyword_appending = ''
    for word in observation_list:
        word_counter += 1
        if keyword_appending == '':
            keyword_appending += word
        else:
            keyword_appending += ' ' + word
        if keyword_appending in df2.Condition:
            keyword_counter = word_counter
    if keyword_counter == 0 or keyword_counter == observation_length:
        Observation_Sentence = '{} was observed.'.format(observation)
        return Observation_Sentence
    elif 0 < keyword_counter < observation_length:
        keyword_string = ' '.join(str(key) for key in observation_list[0:keyword_counter])
        location_string = ' '.join(str(a) for a in observation_list[keyword_counter:])
        Observation_Sentence = '{} was observed {}.'.format(keyword_string, location_string)
        return Observation_Sentence

def SentenceGenerator_AssetObservationSummary(conditionsentence, obsentence1, obsentence2, obsentence3):
    # Function to merge all four sentences together

    # store input sentences in a list
    sentence_list = [obsentence1, obsentence2, obsentence3]

    # parameter required for this operation
    observationsentence_list = []
    keywords = []
    # store key value list, all keywords (without location) saved in keywords list, and if keywords has location, the
    # fully sentence is saved in observationsentence_list.

    for sentence in sentence_list:
        sentenceword_list = sentence.split(' ')
        if sentenceword_list[-1] == "observed.":
            keywords.append(' '.join(str(word).lower() for word in sentenceword_list[:-2]))
        else:
            observationsentence_list.append(sentence.capitalize())

    final_sentence = ''
    if len(keywords) == 1:
        final_sentence = '{} was observed.'.format(keywords[0])
    elif len(keywords) == 2:
        final_sentence = '{} were observed.'.format(' and '.join(str(word1) for word1 in keywords))
    else:
        final_sentence = ''

    if len(observationsentence_list) == 0:
        final_sentence += ''
    elif len(observationsentence_list) > 0:
        final_sentence = final_sentence + ' {}'.format(' '.join(str(word2) for word2 in observationsentence_list))

    final_sentence = final_sentence.lstrip(' ')

    final_sentence = final_sentence.split(' ')
    final_sentence[0] = final_sentence[0].capitalize()
    final_sentence = ' '.join(str(word) for word in final_sentence)

    final_sentence = conditionsentence + ' ' + final_sentence
    return final_sentence

def SentenceGenerator_Sudbury_AssetObservationSummary(conditionsentence, obsentence1, obsentence2, obsentence3):
    # Function to merge all four sentences together
    # Special requirement for Sudbury, this function will analyze all inputs and split out information into a list of
    # [Condition Comment, Code Concern, Health and Safety Issue, Operation and Maintenance Comments]

    # hided
    # no_codeconcern = 'No code concerns were identified during the assessment.'
    # no_healthsafty = 'No H&S concerns were identified during the assessment.'

    # for sudbury, if there is no issue, we leave them empty
    no_codeconcern = ''
    no_healthsafty = ''


    # identify
    comments_output = ['', no_codeconcern, no_healthsafty, '']

    # store input sentences in a list
    sentence_list = [obsentence1, obsentence2, obsentence3]

    # the code below to identify if the observation has H&S, code concern, operation maintenance comments; if it has,
    # then store this observation into the corresponding list, and removed this observation from sentence_list

    for num, sentence in enumerate(sentence_list):
        sentence = sentence.lower()
        if 'cc:' in sentence:
            comments_output[1] = sentence[4:].capitalize()
            sentence_list[num] = ''
        elif 'hs:' in sentence:
            comments_output[2] = sentence[4:].capitalize()
            sentence_list[num] = ''
        elif 'om:' in sentence:
            comments_output[3] = sentence[4:].capitalize()
            sentence_list[num] = ''

    # parameter required for this operation
    observationsentence_list = []
    keywords = []
    # store key value list, all keywords (without location) saved in keywords list, and if keywords has location, the
    # fully sentence is saved in observationsentence_list.

    for sentence in sentence_list:
        sentenceword_list = sentence.split(' ')
        if sentenceword_list[-1] == "observed.":
            keywords.append(' '.join(str(word).lower() for word in sentenceword_list[:-2]))
        else:
            observationsentence_list.append(sentence.capitalize())

    final_sentence = ''
    if len(keywords) == 1:
        final_sentence = '{} was observed.'.format(keywords[0])
    elif len(keywords) == 2:
        final_sentence = '{} were observed.'.format(' and '.join(str(word1) for word1 in keywords))
    else:
        final_sentence = ''

    if len(observationsentence_list) == 0:
        final_sentence += ''
    elif len(observationsentence_list) > 0:
        final_sentence = final_sentence + ' {}'.format(' '.join(str(word2) for word2 in observationsentence_list))

    final_sentence = final_sentence.lstrip(' ')

    final_sentence = final_sentence.split(' ')
    final_sentence[0] = final_sentence[0].capitalize()
    final_sentence = ' '.join(str(word) for word in final_sentence if word != '')

    final_sentence = conditionsentence + ' ' + final_sentence
    comments_output[0] = final_sentence

    return comments_output

def SentenceGenerator_AssetReplacement(AssetRemainingServiceLife):
    response = ''
    if AssetRemainingServiceLife == 0:
        response = ' Asset to be replaced in the immediate term.'
    if 1 <= AssetRemainingServiceLife <= 5:
        response = ' Asset to be replaced in the short term.'
    if 5 < AssetRemainingServiceLife <= 10:
        response = ' Asset to be replaced in the intermediate term.'
    if 10 < AssetRemainingServiceLife <= 20:
        response = ' Asset to be replaced in the long term.'
    return response

def SentenceGenetator_ObservationandRecommendation(AssetConditionComment, AssetRehabComment, AssetName, AssetESL):
    AssetName = AssetName.lower()
    AssetName = AssetName.capitalize()
    replacement_sentence = AssetName + ' to be replaced at the end of remaining service life.'

    if AssetESL == 0:
        replacement_sentence = 'Asset to be replaced in the immediate term.'
    if 1 <= AssetESL <= 5:
        replacement_sentence = 'Asset to be replaced in the short term.'
    if 5 < AssetESL <= 10:
        replacement_sentence = 'Asset to be replaced in the intermediate term.'
    if 10 < AssetESL <= 20:
        replacement_sentence = 'Asset to be replaced in the long term.'

    Combined_Sentence = AssetConditionComment + AssetRehabComment + ' ' + replacement_sentence
    return Combined_Sentence

def Database_Cleaning(dataframe):
    dataframe = dataframe.replace(np.nan, '', regex=True)
    return dataframe

def Databbase_Cleaning_RemoveZero(Input):
    if input == 0:
        return ''

def converter_FillinNumber(input, digit):
    if input == '':
        return ''
    input = int(input)
    input = str(input)
    input = input.zfill(digit)
    return input

def converter_ListtoDataframe(list,IndexName):
    # to convert list to dataframe and assign header and index
    df_list = pd.DataFrame(list)

    # first row to become header
    header = df_list.ilock[0]
    df_list = df_list[1:]
    df_list.column = header

    # set first column index as index of dataframe
    df_list.set_index([IndexName])

    return df_list

def converter_Barrie_ConditionRating(condition_rating):
    # Function to convert condition rating to Barrie's standard
    if condition_rating == 1 or condition_rating == 2:
        return 1
    elif condition_rating == 3:
        return 2
    elif condition_rating == 4:
        return 3
    elif condition_rating == 5:
        return 4
    else:
        return condition_rating

#def Analysis_NewDataBase_CommentIdentification():
    # Function to analyze raw comment data from George's database


#
# def Database_SkipRemovedAsset(AssetName):
#     AssetName = AssetName.split(' ')
#     keyword = '(Removed)'
#     if AssetName[0] == '(Removed)':
#         return continue
#     else:
#         return pass