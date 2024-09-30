# Niroopma Verma
# 2021csb1115@iitrpr.ac.in
# Data Analytics

#  Importing the required modules
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Loading the excel file
xls = pd.ExcelFile("CreditBanking_Project-1_UPDATED.xlsx")

# Loading each sheet into separate dataframes
df_acquisition = pd.read_excel(xls, 'Customer Acqusition')
df_spend = pd.read_excel(xls, 'Spend')
df_repayment = pd.read_excel(xls, 'Repayment')


# TASK_1 : Monthly spend of each customer
monthly_spend = df_spend.groupby(['Customer', 'Month'], observed = True)['Amount'].sum().reset_index()
# print(monthly_spend)



# TASK_2 : Monthly repayment of each customer.
monthly_repayment = df_repayment.groupby(['Customer', 'Month'], observed=True)['Amount'].sum().reset_index()
# print(monthly_repayment)



# TASK_3 : Highest paying 10 customers.
total_spend_per_customer = df_spend.groupby('Customer', observed=True)['Amount'].sum().reset_index()
highest_paying_customers = total_spend_per_customer.nlargest(10, 'Amount')
# print(highest_paying_customers)



# TASK_4 : People in which segment are spending more money.
spend_with_segment = df_spend.merge(df_acquisition[['Customer', 'Segment']], left_on = 'Customer', right_on = 'Customer')
spend_per_segment = spend_with_segment.groupby('Segment', observed = True)['Amount'].sum().reset_index().sort_values(by = 'Amount', ascending = False)
# print(spend_per_segment)
# Plotting the pie chart
plt.figure(figsize=(8, 8))
plt.pie(spend_per_segment['Amount'], labels=spend_per_segment['Segment'], autopct='%1.1f%%', startangle=140)
plt.title('Segment Spending More Money')
plt.axis('equal')
plt.savefig('segment_spending.png') # Saving the pie chart as a PNG file
# plt.show()



# TASK_5 :  Which age group is spending more money?
df_acquisition['age_group'] = pd.cut(df_acquisition['Age'], bins = [18, 25, 35, 45, 55, 65, 100], labels = ['18-25', '26-35', '36-45', '46-55', '56-65', '65+'])
spend_with_age_group = df_spend.merge(df_acquisition[['Customer', 'age_group']], left_on = 'Customer', right_on = 'Customer')
spend_per_age_group = spend_with_age_group.groupby('age_group', observed=True)['Amount'].sum().reset_index().sort_values(by = 'Amount', ascending = False)
# print(spend_per_age_group)
# Plotting the histogram
plt.figure(figsize=(5, 6))
sns.histplot(data=spend_per_age_group, x='age_group', weights='Amount', kde=False, bins=6)
plt.title('Spending by Age Group')
plt.xlabel('Age Group')
plt.ylabel('Total Amount Spent')
plt.xticks(rotation=25)
plt.tight_layout()
plt.savefig('spending_by_age_group_histogram.png') # Saving the histogram as a PNG file
# plt.show()



# TASK_6 :  Which is the most profitable segment?
# Assuming profit is calculated as spend - repayment 
spend_with_repayment = df_spend.merge(df_repayment, on = ['Customer', 'Month'], suffixes=('_spend', '_repay'))
profit_per_customer = spend_with_repayment.groupby('Customer', observed = True).apply(lambda x: x['Amount_spend'].sum() - x['Amount_repay'].sum(), include_groups = False).reset_index(name = 'Profit')
profit_with_segment = profit_per_customer.merge(df_acquisition[['Customer', 'Segment']], left_on = 'Customer', right_on = 'Customer')
profit_per_segment = profit_with_segment.groupby('Segment', observed = True)['Profit'].sum().reset_index().sort_values(by = 'Profit', ascending=False)
# print(profit_per_segment)
#  Plotting the bar graph
plt.figure(figsize=(5, 6))
plt.bar(profit_per_segment['Segment'], profit_per_segment['Profit'], color='skyblue')
plt.xlabel('Segment')
plt.ylabel('Total Profit')
plt.title('Most Profitable Segment')
plt.xticks(rotation=25)
plt.tight_layout()
plt.savefig('most_profitable_segment.png')  # Saving the bar graph as a PNG file
# plt.show()



# TASK_7 : In which category the customers are spending more money?
spend_per_category = df_spend.groupby('Type', observed=True)['Amount'].sum().reset_index().sort_values(by = 'Amount', ascending= False)
# print(spend_per_category)
# Plotting the pie chart
plt.figure(figsize=(8, 8))
plt.pie(spend_per_category['Amount'], labels=spend_per_category['Type'], autopct='%1.1f%%', startangle=140)
plt.title('Category Spending More Money')
plt.axis('equal')
plt.savefig('category_spending.png') # Saving the pie chart as a PNG file
# plt.show()



# TASK_8 : Impose an interest rate of 2.9% for each customer for any due amount.
spend_with_repayment['due_amount'] = spend_with_repayment['Amount_spend'] - spend_with_repayment['Amount_repay']
spend_with_repayment['interest_due'] = spend_with_repayment['due_amount'] * 0.029
# print(spend_with_repayment)



# TASK_9 : Monthly profit for the bank.
# Assuming profit is calculated as (spend - repayment) + interest
spend_with_repayment['monthly_profit'] = spend_with_repayment['Amount_spend'] - spend_with_repayment['Amount_repay'] + spend_with_repayment['interest_due']
monthly_profit = spend_with_repayment.groupby('Month', observed=True)['monthly_profit'].sum().reset_index()
# print(monthly_profit)



# Combining all the outputs into a single excel file (Outputs.xlsx)
with pd.ExcelWriter('Outputs.xlsx', engine = 'openpyxl') as writer:
    monthly_spend.to_excel(writer, sheet_name = 'monthly spend', index = False)
    monthly_repayment.to_excel(writer, sheet_name = 'monthly repayment', index = False)
    highest_paying_customers.to_excel(writer, sheet_name= 'highest paying customers', index = False)
    spend_per_segment.to_excel(writer, sheet_name= 'spend per segmennt', index = False)
    spend_per_age_group.to_excel(writer, sheet_name= 'spend per age group', index = False)
    profit_per_segment.to_excel(writer, sheet_name= 'profit per segment', index = False)
    spend_per_category.to_excel(writer, sheet_name= 'spend per category', index = False)
    spend_with_repayment.to_excel(writer, sheet_name= 'spend with repayment', index = False)
    monthly_profit.to_excel(writer, sheet_name= 'monthly profit', index = False)
