import pandas as pd
from pathlib import Path
import numpy as np

# Author: Delin Mu
# Update note:
def yorkregion_summarytable(file_name, file_location, sheet_name, result_name, result_location):
    # Purpose:
    # Analyze a raw inventory to generate data summary for reporting purpose
    #
    # Inputs:
    # file_name: spreadsheet that contains raw data
    # file_location: the full path (recommended) of the file is to be processed
    # sheet_name: sheet where data is saved at
    # result_name: user to create a name for the spreadsheet that stores the processed data
    # result_location: user to provide input for location of the processed files
    #
    # Constraints: the inventory shall have the following for the name of the columns:
    # 'CategoryCode', 'condition', 'repCost', 'COF', 'remainingESL'


    def rep_result(df, name, finalname):
        result_df = df.groupby(['CategoryCode'])[name].sum().to_frame()
        result_df[finalname] = (result_df[name] / df_rep_cost['Total Replacement Cost']).round(decimals=2)
        result_df = result_df.drop([name], axis=1)
        return result_df

    def df_merge(df_list, key):
        i = 0
        for df in df_list:
            if i == 0:
                df_result = df
            else:
                df_result = pd.merge(df_result, df, on=key)
            i += 1
        return df_result

    def round_half_up(number, decimals=0):
        multiplier = 10 ** decimals
        new_number = number * multiplier + 0.5
        return np.floor(new_number) / multiplier

    # file path: locate spreadsheet path
    excelraw_path = Path(file_location)

    file_name = excelraw_path / file_name

    # prepare df
    excel_file = pd.ExcelFile(file_name)
    df = excel_file.parse(sheet_name).set_index('seqNum')

    # add new columns to the df for analysis
    df['cost_condition'] = df['condition'] * df['repCost']
    df['cost_COF'] = df['COF'] * df['repCost']
    df['cost_risk'] = df['COF'] * df['condition'] * df['repCost']
    df['cost_esl'] = df['remainingESL'] * df['repCost']

    # prepare an empty df to store results
    df_result = pd.DataFrame()

    # calculate replacement cost based on asset category
    df_rep_cost = df.groupby(['CategoryCode'])['repCost'].sum().to_frame()
    df_rep_cost = df_rep_cost.rename(columns={'repCost': 'Total Replacement Cost'})

    # calculate and round half-up COF (replacement based)
    df_COF_rep = round_half_up(rep_result(df, 'cost_COF', 'Average COF'), 1)

    # calculate and round half-up Condition (replacement based)
    df_condition_rep = round_half_up(rep_result(df, 'cost_condition', 'Average Condition'), 1)

    # calculate and round half-up Risk (replacement based)
    df_risk_rep = round_half_up(rep_result(df, 'cost_risk', 'Average Risk'), 1)

    # calculate and round down ESL (replacement based)
    df_esl_rep = np.floor(rep_result(df, 'cost_esl', 'Average ESL'))

    # add therotical service life column
    esl_stand = {'AC': 15, 'BM': 30, 'ES': 25, 'HSS': 15, 'PM': 30, 'SC': 60, 'SCS': 15, 'SIC': 15, 'SW': 20}
    df_esl_stand = df.from_dict(esl_stand, orient='index', columns=['Theoretical Service Life'])
    df_esl_stand.index.name = 'CategoryCode'

    # append all required dataframe
    df_list = [df_esl_stand, df_esl_rep, df_condition_rep, df_COF_rep, df_risk_rep, df_rep_cost]

    # merge all results together
    df_result = df_merge(df_list, 'CategoryCode')

    # Generate total average data
    TotalReplacementCost = int(df_rep_cost.sum())

    average_condition = (df_result['Average Condition'] * df_result['Total Replacement Cost']).sum() / \
                        TotalReplacementCost
    average_condition = round_half_up(average_condition, 1)

    average_cof = (df_result['Average COF'] * df_result['Total Replacement Cost']).sum() / \
                  TotalReplacementCost
    average_cof = round_half_up(average_cof, 1)

    average_risk = (df_result['Average Risk'] * df_result['Total Replacement Cost']).sum() / \
                   TotalReplacementCost
    average_risk = round_half_up(average_risk, 1)

    # Create a new data entry to store all average values
    df_totalaverage = pd.DataFrame({'Total Replacement Cost': [TotalReplacementCost],
                                    'Average Condition': [average_condition],
                                    'Average COF': [average_cof],
                                    'Average Risk': [average_risk], })
    df_totalaverage.index.name = 'CategoryCode'
    df_totalaverage.rename(index={0: 'Total'}, inplace=True)

    # practice mode
    df_result = pd.concat([df_result, df_totalaverage], sort=False).fillna('')

    # create path name for the processed data
    processed_path = Path(result_location)
    result_name = processed_path / result_name

    # data cleaning for report (saving decimal points)
    df_result = df_result.to_excel(result_name, sheet_name='Sheet 2')
