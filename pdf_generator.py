from fpdf import FPDF
import unicodedata
from datetime import date

from excelparser import company_name, company_ticker, company_stock_close, company_marketcap, company_n_shares_outstanding, df_dcf, company_description

# create a pdf document
pdf = FPDF(orientation="P", unit="pt", format="A4")

# add page to the pdf
pdf.add_page()

# set a font
pdf.set_font(family="Arial", size=24, style="B")

# create a title (w is the width of the cell, if 0 then it will be the width of the page)
# align = C means center, border = 1 adds a border, ln = 1 means go to the next line
pdf.cell(w=0, h=50, txt=f"{company_name} Report", align="C", border=1, ln=1)

# set a font
pdf.set_font(family="Arial", size=12)

pdf.cell(w=0, h=16, txt="", ln=1)
pdf.cell(w=0, h=16, txt=f"Report date: {date.today().strftime('%d/%m/%Y')}", ln=1)
pdf.cell(w=0, h=16, txt=f"Ticker: {company_ticker}", ln=1)
pdf.cell(w=0, h=16, txt=f"Close: {company_stock_close}$", ln=1)
pdf.cell(w=0, h=16, txt="", ln=1)

# about the company
pdf.set_font(family="Arial", size=18, style="B")
pdf.cell(w=0, h=16, txt=f"About", ln=1)

pdf.set_font(family="Arial", size=12)
# multi_cell is used for multiple lines of text
pdf.multi_cell(w=0, h=16, txt=f"{unicodedata.normalize('NFKD', company_description).encode('ASCII', 'ignore').decode('ASCII')}")

# output the pdf
pdf.output("test.pdf")