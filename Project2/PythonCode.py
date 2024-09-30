# Niroopma Verma
# 2021csb1115@iitrpr.ac.in
# Data Analytics
# Project 2

# Importing the requires libraries. Pandas for data analysis and matplotlib for data visualisation
import pandas as pd
import matplotlib.pyplot as plt

# Loading data from the given excel file 'Credit_card_transactions_Project2'
df = pd.read_excel('Credit_card_transactions_Project2.xlsx')


# Task 1: Top 5 cities with highest spends and their percentage contribution of total credit card spends
total_spends = df['Amount'].sum()       # taking sum of the spends
city_spends = df.groupby('City')['Amount'].sum().sort_values(ascending=False).head(5)       # grouping by City column, summing the amount and sorting the data and selecting top 5 values        
city_percentage_contribution = (city_spends / total_spends) * 100       # calculating the percentage of spends per city
task1_result = pd.DataFrame({'City': city_spends.index, 'Total Spends': city_spends.values, 'Percentage Contribution': city_percentage_contribution.values})        # saving results in the dataframe

# Bar graph for Task 1
plt.figure(figsize=(8, 5))
plt.bar(task1_result['City'], task1_result['Total Spends'], color='skyblue')        # creating the bar graph 
plt.xlabel('City')
plt.ylabel('Total Spends')
plt.title('Top 5 Cities with Highest Spends')
plt.xticks(rotation=45)
plt.savefig('Task1.png')        # saving the plot



# Task 2: Highest spend month and amount spent in that month for each card type
df['Month'] = df['Date'].dt.to_period('M')      # extracting the month from the date column
monthly_spends = df.groupby(['Card Type', 'Month'])['Amount'].sum().reset_index()       # grouping by card type and Month and summing the amount
task2_result = monthly_spends.loc[monthly_spends.groupby('Card Type')['Amount'].idxmax()]       # find the month with the highest spend for each card type

# Bar graph for Task 2
plt.figure(figsize=(8, 5))
for card_type in task2_result['Card Type'].unique():
    data = task2_result[task2_result['Card Type'] == card_type]
    plt.bar(data['Month'].astype(str), data['Amount'], label=card_type)
plt.xlabel('Month')
plt.ylabel('Amount Spent')
plt.title('Highest Spend Month for Each Card Type')
plt.legend(title='Card Type')
plt.xticks(rotation=45)
plt.savefig('Task2.png')



# Task 3: Transaction details for each card type when it reaches a cumulative of 1000000 total spends
cumulative_spends = df.groupby('Card Type')['Amount'].cumsum()      #grouping by card type and taking the cumulative sum for each type
task3_result = df.loc[cumulative_spends[cumulative_spends >= 1000000].groupby(df['Card Type']).idxmin()]        # extracing data for the card type when the cumulatiive sum reaches >= 1000000



# Task 4: City with lowest percentage spend for gold card type
gold_spends = df[df['Card Type'] == 'Gold']        # selecting only Gold Card type data
gold_city_spends = gold_spends.groupby('City')['Amount'].sum()      # gouping by City and summing tha amount
gold_total_spends = gold_city_spends.sum()
gold_city_percentage = (gold_city_spends / gold_total_spends) * 100     # calculating the percentage of spends for gold type
lowest_gold_city = gold_city_percentage.idxmin()        # selecting the city with lowest spends
task4_result = pd.DataFrame({'City': [lowest_gold_city], 'Percentage': [gold_city_percentage[lowest_gold_city]]})



# Task 5: Highest and lowest expense type for each city
expense_type_spends = df.groupby(['City', 'Exp Type'])['Amount'].sum().unstack().fillna(0)      # grouping by city and exp type, summing the amount and unstacking the result to create a pivot table
task5_result = pd.DataFrame({
    'City': expense_type_spends.index,
    'Highest Expense Type': expense_type_spends.idxmax(axis=1),
    'Lowest Expense Type': expense_type_spends.idxmin(axis=1)
}).reset_index(drop=True)       # saving the results in a dataframe



# Task 6: Percentage contribution of spends by females for each expense type
gender_expense_spends = df.groupby(['Exp Type', 'Gender'])['Amount'].sum().unstack().fillna(0)      # grouping by gender and exp type, summing the amount and unstacking the result to create a pivot table
female_expense_percentage = (gender_expense_spends['F'] / gender_expense_spends.sum(axis=1)) * 100      # calculating the percentage contribution in the expense by women
task6_result = pd.DataFrame({'Expense Type': female_expense_percentage.index, 'Female Spend Percentage': female_expense_percentage.values})

# Bar graph for Task 6
plt.figure(figsize=(8, 5))
plt.bar(task6_result['Expense Type'], task6_result['Female Spend Percentage'], color='lightcoral')
plt.xlabel('Expense Type')
plt.ylabel('Female Spend Percentage')
plt.title('Percentage Contribution of Spends by Females for Each Expense Type')
plt.xticks(rotation=45)
plt.savefig('Task6.png')



# Task 7: Card and expense type combination with highest month over month growth in Jan-2014

# Filter the data for December 2013 and January 2014
dec_data = df[df['Date'].dt.to_period('M') == '2013-12']
jan_data = df[df['Date'].dt.to_period('M') == '2014-01']

# Group by Card Type and Expense Type and sum the Amount for each group
dec_grouped = dec_data.groupby(['Card Type', 'Exp Type'])['Amount'].sum().reset_index()
jan_grouped = jan_data.groupby(['Card Type', 'Exp Type'])['Amount'].sum().reset_index()

merged_data = pd.merge(dec_grouped, jan_grouped, on=['Card Type', 'Exp Type'], suffixes=('_Dec', '_Jan'))       # Merging the two DataFrames on Card Type and Exp Type

merged_data['MoM_Growth'] = ((merged_data['Amount_Jan'] - merged_data['Amount_Dec']) / merged_data['Amount_Dec']) * 100     # Calculating the month-over-month growth

max_growth = merged_data.loc[merged_data['MoM_Growth'].idxmax()]        # Finding the combination with the highest month-over-month growth

task7_result = pd.DataFrame({
    'Card Type': [max_growth['Card Type']],
    'Expense Type': [max_growth['Exp Type']],
    'Growth': [max_growth['MoM_Growth']]
})      # saving the resultin a dataframe



# Task 8: City with highest total spend to total number of transactions ratio during weekends
df['DayOfWeek'] = df['Date'].dt.dayofweek       # creating a dataframe for the days 
weekends = df[df['DayOfWeek'] >= 5]     # selecting the weekends
weekend_city_spends = weekends.groupby('City')['Amount'].sum()      # grouping by city and taking sum of the amount
weekend_city_transactions = weekends.groupby('City').size()   
spend_to_transaction_ratio = weekend_city_spends / weekend_city_transactions        # taking the ratio of total number of transactions on weekends to the total number of transactions
highest_ratio_city = spend_to_transaction_ratio.idxmax()        # selecting the city with highest value of transactions on weekends
task8_result = pd.DataFrame({
    'City': [highest_ratio_city],
    'Spend to Transaction Ratio': [spend_to_transaction_ratio[highest_ratio_city]]
})      # saving the results



# Task 9: City that took least number of days to reach its 500th transaction after first transaction in that city
first_transaction_dates = df.groupby('City')['Date'].min()      # grouping by city and selecting the minimum date of transation
df['Days Since First'] = df['Date'] - df['City'].map(first_transaction_dates)       # calculating the number of days since the first transaction
transactions_count = df.groupby(['City', 'Days Since First']).size().groupby(level=0).cumsum()      # grouping by  city and days since first transaction and calculating the cumulative sum
city_500th_transaction_day = transactions_count[transactions_count >= 500].groupby(level=0).idxmin()        # calculating the day on which each city reaches its 500th transaction
days_to_500_transactions = city_500th_transaction_day.map(lambda x: x[1].days)      # calculating the number of days it took for each city to reach 500 transactions
fastest_city_to_500_transactions = days_to_500_transactions.idxmin()        # selecting the city that reached the 500 transactions the fastest
task9_result = pd.DataFrame({
    'City': [fastest_city_to_500_transactions],
    'Days to 500 Transactions': [days_to_500_transactions[fastest_city_to_500_transactions]]
})      # saving the results in a dataframe



# Saving all results in a single Excel file (results.xlsx)
with pd.ExcelWriter('outputs.xlsx') as writer:
    task1_result.to_excel(writer, sheet_name='Task 1', index=False)
    task2_result.to_excel(writer, sheet_name='Task 2', index=False)
    task3_result.to_excel(writer, sheet_name='Task 3', index=False)
    task4_result.to_excel(writer, sheet_name='Task 4', index=False)
    task5_result.to_excel(writer, sheet_name='Task 5', index=False)
    task6_result.to_excel(writer, sheet_name='Task 6', index=False)
    task7_result.to_excel(writer, sheet_name='Task 7', index=False)
    task8_result.to_excel(writer, sheet_name='Task 8', index=False)
    task9_result.to_excel(writer, sheet_name='Task 9', index=False)
