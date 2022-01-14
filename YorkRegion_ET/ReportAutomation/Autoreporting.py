import docxtpl
import openpyxl
import os.path
import ReportingFunctions.ReportingFunctions as rf

from docxtpl import DocxTemplate
import jinja2
import pandas as pd

# pass the word template to variable doc
doc = DocxTemplate('report_template.docx')

# pass the input form data into a dataframe

raw_data = pd.read_excel('input_form.xlsx', 'Raw_Data')

# Below for Stage 2 data analysis. Note that input_form shall be opened and click manual update.
# create table 5-1
table5_1 = rf.yorkregion_summarytable('input_form.xlsx',
                   'E:/DelinM_IOS/Operation YorkET/ReportAutomation'
                   , 'Raw_Data',
                   'Stage_2.xlsx',
                   'E:\DelinM_IOS\Operation YorkET\ReportAutomation')


#
# Stage 1 Operation
input_data_stage1 = pd.read_excel('input_form.xlsx','Stage_1')

# dataframe will need to ignore the var name has n/a
# if n/a is not removed, it will create issue in word template
input_data_stage1 = input_data_stage1[input_data_stage1['var_name'].notna()]

# zip dataframe into a dict which is acceptable for template
context_stage1 = dict(zip(input_data_stage1['var_name'], input_data_stage1['parameter']))

# Render template using context dic
doc.render(context_stage1)

# Stage 2 Operation
input_data_stage2 = pd.read_excel('input_form.xlsx', 'Stage_2')
input_data_stage2 = input_data_stage2[input_data_stage2['var_name'].notna()]
context_stage2 = dict(zip(input_data_stage2['var_name'], input_data_stage2['parameter']))
print(context_stage2)
doc.render(context_stage2)



# save the processed word to a specific path
file_name = 'Processed Doc.docx'
path = 'E:/DelinM_IOS/Operation YorkET/ReportAutomation/Updated Report/'
doc.save(path + file_name)


