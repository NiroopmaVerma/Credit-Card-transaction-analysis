# Niroopma Verma
# 2021csb1115@iitrpr.ac.in
# Data Analytics

# Importing the required module
import pandas as pd

# Loading the excel file
xls = pd.ExcelFile("Credit Banking_Project - 1.xls")

# Loading each sheet into separate dataframes
df_acquisition = pd.read_excel(xls, 'Customer Acqusition')
df_spend = pd.read_excel(xls, 'Spend')
df_repayment = pd.read_excel(xls, 'Repayment')

# Sanity Check : P rovide a meaningful treatment to all values where age is less than 18.
df_acquisition.loc[df_acquisition['Age'] < 18, 'Age'] = 18

# Sanity Check : Is there any customer who have spent more than his/her Credit Limit for any particular month.
spend_with_limit = df_spend.merge(df_acquisition[['Customer', 'Limit']], left_on='Customer', right_on='Customer')
spend_with_limit_per_month = spend_with_limit.groupby(['Customer', 'Month'])['Amount'].sum().reset_index()
spend_with_limit_per_month = spend_with_limit_per_month.merge(df_acquisition[['Customer', 'Limit']], left_on='Customer', right_on='Customer')
over_limit_spend = spend_with_limit_per_month[spend_with_limit_per_month['Amount'] > spend_with_limit_per_month['Limit']]
# print(over_limit_spend)

#  Combining all the sheets into a new file named 'CreditBanking_Project-1_UPDATED.xlsx' after performing the sanity checks 
#  and saving it for performing the tasks given
with pd.ExcelWriter('CreditBanking_Project-1_UPDATED.xlsx', engine = 'openpyxl') as writer:
    df_acquisition.to_excel(writer, sheet_name = 'Customer Acqusition', index = False)
    df_spend.to_excel(writer, sheet_name = 'Spend', index = False)
    df_repayment.to_excel(writer, sheet_name = 'Repayment', index = False)
    over_limit_spend.to_excel(writer, sheet_name = 'Over Limit Spend', index = False)
