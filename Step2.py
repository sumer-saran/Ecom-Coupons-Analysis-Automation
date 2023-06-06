import pandas as pd

# Read the XLSX file
df = pd.read_excel('OutputFile.xlsx')
df['coupons'].fillna("No Coupon Used", inplace=True)

# Count the number of orders per coupon code bucket
coupon_counts = df.groupby('coupons')['Order Number'].nunique().reset_index()

coupon_counts = coupon_counts.rename(columns={'coupons':'Coupons','Order Number':'Total No. Of Orders'})
coupon_value = df.groupby('coupons').agg({'Total': 'sum'}).reset_index()

coupon_value = coupon_value.rename(columns={'coupons':'Coupons','Total':'Total Order Value'})
avg_order=df.groupby('coupons').agg({'Total':'mean'}).reset_index()

avg_order = avg_order.rename(columns={'coupons':'Coupons','Total':'Average Order Value'})

result = pd.merge(coupon_counts, coupon_value)
result=pd.merge(result,avg_order)

with pd.ExcelWriter('OutputFile.xlsx') as writer:
    result.to_excel(writer, sheet_name='Analysis',index=False)

df = pd.read_excel('OutputFile.xlsx')


total_revenue = df['Total Order Value'].astype(int).sum()
df['Revenue Percentage'] = df['Total Order Value'].astype(int) / total_revenue 

df = df.sort_values('Total No. Of Orders', ascending=False)

column_sums = df.sum(axis=0)
sum_row = pd.DataFrame(column_sums).T
sum_row.iloc[0, 0] = '**Total**'
df_with_sum = pd.concat([df, sum_row], ignore_index=True)

df_with_sum['Revenue Percentage'] = df_with_sum['Revenue Percentage'].apply(lambda x: '{:.2%}'.format(x))
df_with_sum['Total No. Of Orders'] = df_with_sum['Total No. Of Orders'].apply(lambda x: '{:,.0f}'.format(x))
df_with_sum['Total Order Value'] = df_with_sum['Total Order Value'].apply(lambda x: '{:,.0f}'.format(x))
df_with_sum['Average Order Value'] = df_with_sum['Average Order Value'].apply(lambda x: '{:,.0f}'.format(x))


with pd.ExcelWriter('OutputFile.xlsx') as writer:
    df_with_sum.to_excel(writer, sheet_name='Analysis',index=False)


