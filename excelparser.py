import pandas as pd

df = pd.read_excel('finance_test.xlsx')

selected_revenue_cagr = df.iloc[4, 15]
selected_target_margin = df.iloc[5, 15]
selected_wacc = df.iloc[7, 15]
selected_sales_to_capital = df.iloc[10, 3]
selected_taxrate = df.iloc[11, 3]

company_name = df.iloc[15, 3]
company_ticker = df.iloc[16, 3]
company_stock_close = df.iloc[17, 3]
company_marketcap = df.iloc[18, 3]
company_n_shares_outstanding = df.iloc[21, 3]

company_description = df.iloc[14, 5]
company_targetvalue = df.iloc[48, 20]

# read dcf data
df_dcf = df.iloc[38:55, 1:16]
df_dcf.set_index("Unnamed: 1", inplace=True)

print(company_description)
print(company_targetvalue)