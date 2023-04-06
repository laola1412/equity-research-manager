import pandas as pd

def parse_excel_file(xlsx):
    df = pd.read_excel(xlsx)

    selected_revenue_cagr = df.iloc[4, 15]
    selected_target_margin = df.iloc[5, 15]
    selected_wacc = df.iloc[7, 15]
    selected_sales_to_capital = df.iloc[10, 3]
    selected_taxrate = df.iloc[11, 3]

    company_name = df.iloc[15, 3]
    company_ticker = df.iloc[16, 3]
    company_stock_close = df.iloc[17, 3]
    company_marketcap = df.iloc[20, 3]
    company_n_shares_outstanding = df.iloc[22, 3]

    company_description = df.iloc[14, 5]
    company_targetvalue = df.iloc[22, 3]

    stock_rating = df.iloc[0, 8]

    # read dcf data
    df_dcf = df.iloc[38:55, 1:16]
    df_dcf.set_index("Unnamed: 1", inplace=True)
    
    # clean company_name; remove comma if in company_name
    company_name = company_name.split(",")[0] if "," in company_name else company_name
    
    print(company_targetvalue)
    return company_name, company_ticker, company_stock_close, company_targetvalue, stock_rating, company_description