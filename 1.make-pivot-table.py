import pandas as pd

# Read Excel File
df = pd.read_excel('supermarket_sales.xlsx')


df = df[['Gender', 'Product line', 'Total']]  # Selecting required columns only                  


pivot_table = df.pivot_table(index='Gender', columns='Product line',            #making pivot tables,   index is gender and the columns will be the product line
                             values='Total', aggfunc='sum').round(0)            #values - specifies the column that has to be aggregated, here total column ke values are aggregated
                                                                                #aggfunc - specifies the function to be applied to the values, here sum is done during aggregation
print(pivot_table)

pivot_table.to_excel('pivot_table.xlsx', 'Report', startrow=4)      # Export pivot table to Excel file

#Report is the name of the sheet in excel  #starts from row number 4