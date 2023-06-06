import pandas as pd
# Read the XLSX file
df = pd.read_excel('IGNORE.xlsx')

# Convert 'Dates' column to datetime
df['Order Date & Time'] = pd.to_datetime(df['Order Date & Time'],format='%d/%m/%Y %H:%M:%S')

df_filtered = df[df['coupons'].notnull()]

# Find the first usage date of each coupon
first_usage_dates = df_filtered.groupby('coupons')['Order Date & Time'].min().reset_index()

# Calculate the revenue 10 days before/adter the first usage of each coupon
revenue_10_days_before = []
revenue_10_days_after = []
percentage_inc= []
for _, row in first_usage_dates.iterrows():
    coupon = row['coupons']
    first_usage_date = row['Order Date & Time']
    start_date = first_usage_date - pd.DateOffset(days=7)
    end_date2 = first_usage_date + pd.DateOffset(days=7)
    filtered_data = df[(df['Order Date & Time'] >= start_date) & (df['Order Date & Time'] < first_usage_date)]
    filtered_data2= df[(df['Order Date & Time'] >= first_usage_date) & (df['Order Date & Time'] < end_date2)]
    revenue = filtered_data['Total'].sum()
    revenue2=filtered_data2 ['Total'].sum()
    revenue_10_days_before.append(revenue)
    revenue_10_days_after.append(revenue2)
    percentage_increase=((revenue2-revenue)/revenue2)
    percentage_inc.append(percentage_increase)

# Add the revenue values to the first_usage_dates DataFrame
first_usage_dates['Revenue A Week Before First Usage'] = revenue_10_days_before
first_usage_dates['Revenue A Week After First Usage'] = revenue_10_days_after
first_usage_dates['Percentage Increase/Decrease'] = percentage_inc
first_usage_dates = first_usage_dates.sort_values('Percentage Increase/Decrease', ascending=False)
first_usage_dates['Percentage Increase/Decrease'] = first_usage_dates['Percentage Increase/Decrease'].apply(lambda x: '{:.2%}'.format(x))

first_usage_dates = first_usage_dates.rename(columns={'coupons':'Coupons','Order Date & Time':'Start Date'})

existing_file = 'OutputFile.xlsx'

with pd.ExcelWriter(existing_file, engine='openpyxl', mode='a') as writer:
    sheet_name = 'Comparison'

    # Write the DataFrame to the new sheet
    first_usage_dates.to_excel(writer, sheet_name=sheet_name, index=False)
