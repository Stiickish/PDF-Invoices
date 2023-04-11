import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Getting all excel documents
filepaths = glob.glob("invoices/*.xlsx")

# Loop through the files
for filepath in filepaths:

    # Prepare the PDF format
    # Add page to PDF
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # Use .stem to get filename
    filename = Path(filepath).stem
    invoice_nr = filename.split("-")[0]
    invoice_date = filename.split("-")[1]

    # Set font
    # Make a cell
    pdf.set_font(family="Times", size=18, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}", ln=1)

    pdf.set_font(family="Times", size=18, style="B")
    pdf.cell(w=50, h=8, txt=f"Date {invoice_date}", ln=5)

    # Read the data from sheet 1
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Adding header
    # Replace underscore with space using list comprehension
    # Get every header from Excel document
    columns = df.columns
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # Adding rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]) + "$", border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]) + "$", border=1, ln=1)

    # Adding total sum at the end of table
    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum) + "$", border=1, ln=1)

    # Add total sum info
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt=f"The total price is {total_sum}" + "$", ln=1)

    # Add company name and logo
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=25, h=8, txt="PythonHow")
    pdf.image("pythonhow.png", w=10)

    pdf.output(f"pdfs/{filename}.pdf")
