import pandas as pd
import openpyxl

def result(file_name, sheet_name, result_name):
    def rep_result(df, name, finalname):
        result_df = df.groupby(['CategoryCode'])[name].sum().to_frame()
        result_df[finalname] = (result_df[name] / df_rep_cost['Total Replacement Cost']).round(decimals=1)
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

    # calculate COF (replacement based)
    df_COF_rep = rep_result(df, 'cost_COF', 'Average COF')

    # calculate condition (replacement based)
    df_condition_rep = rep_result(df, 'cost_condition', 'Average Condition')

    # calculate risk (replacement based)
    df_risk_rep = rep_result(df, 'cost_risk', 'Average Risk')

    # calculate esl (replacement based)
    df_esl_rep = rep_result(df, 'cost_esl', 'Average ESL').round(decimals=0)

    # append all required dataframe
    df_list = [df_esl_rep, df_condition_rep, df_COF_rep,  df_risk_rep, df_rep_cost]


    df_result = df_merge(df_list, 'CategoryCode')




    print(df_result)

    df_result = df_result.to_excel(result_name, sheet_name= 'Sheet 2')


action = result('NNE Database Result.xlsx', 'Sheet1','Final.xlsx')
