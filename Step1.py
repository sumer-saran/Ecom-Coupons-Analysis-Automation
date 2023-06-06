import pandas as pd
import openpyxl

# Read the source Excel file into a Pandas dataframe
df_source = pd.read_excel('Data.xlsx')

# Select the columns you want to copy
columns_to_copy = ['Order Number','coupons', 'Total']
columns_copy=['Order Date & Time','coupons','Total']
# Create a new dataframe containing only the selected columns
df_copy = df_source[columns_to_copy]
df_2=df_source[columns_copy]

# Write the copied columns to a new Excel file
df_copy.to_excel('OutputFile.xlsx', sheet_name='Analysis', index=False)
df_2.to_excel('IGNORE.xlsx',index=False)

workbook = openpyxl.load_workbook('OutputFile.xlsx')
worksheet = workbook['Analysis']
column = worksheet['A']
for cell in column:
    if cell.value is not None:
        split_values = cell.value.split('-')
        cell.value = split_values[0]     
workbook.save('OutputFile.xlsx')






