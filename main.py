from fpdf import FPDF

# create a pdf document
pdf = FPDF(orientation="P", unit="pt", format="A4")

# add page to the pdf
pdf.add_page()

# # add image to the pdf
# pdf.image("smallworld.png", w=210, h=297)

# set a font
pdf.set_font(family="Arial", size=24, style="B")

# create a title (w is the width of the cell, if 0 then it will be the width of the page)
# align = C means center, border = 1 adds a border, ln = 1 means go to the next line
pdf.cell(w=0, h=50, txt="Stock Report", align="C", border=1, ln=1)

# output the pdf
pdf.output("test.pdf")