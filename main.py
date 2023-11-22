import pandas
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")  # Extracting filepaths

for filepath in filepaths:
    total = 0  # total amount payed to calc in the end
    df = pandas.read_excel(filepath, sheet_name="Sheet 1")  # Open xlsx file
    filename = Path(filepath).stem
    split = filename.split('-')  # Splitting the invoice number and the date into list
    number = split[0]  # Extract num
    date = split[1] # Extract date
    pdf = FPDF(orientation="p", unit="mm", format="a4")  # A4 formatted pdf
    pdf.add_page()
    pdf.set_font(family="Times", style="B", size=24)  # Setting font
    pdf.set_text_color(0, 0, 0)  # Black
    pdf.cell(w=0, h=12, txt=f"Invoice Number: {number}", border=0, ln=1, align="l")
    pdf.cell(w=0, h=12, txt=f"Date: {date}", border=0, ln=1, align="l")
    headers = list(df.columns)  # Extracting from the xlsx file the columns headers
    headers = [item.replace("_", " ").title() for item in headers]  # Replacing underlines and capitalizing
    """
    Here we create the first row in the table which contins the headers so we will make them bold and slightly bigger
    """

    pdf.set_font(family="Times", style="B", size=12)
    pdf.cell(w=30, h=8, txt=str(headers[0]), border=1, align="C")
    pdf.cell(w=50, h=8, txt=str(headers[1]), border=1, align="C")
    pdf.cell(w=45, h=8, txt=str(headers[2]), border=1, align="C")
    pdf.cell(w=35, h=8, txt=str(headers[3]), border=1, align="C")
    pdf.cell(w=30, h=8, txt=str(headers[4]), border=1, align="C", ln=1)
    # Here we add ln=1 to make sure we have break line

    """
    Here we create the rest of the table, we will change size and colour 
    """

    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10, style="B")  # We want the product name will be bold
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1, align="C")
        pdf.set_font(family="Times", size=10)  # other columns not bold
        pdf.cell(w=50, h=8, txt=str(row["product_name"]), border=1, align="C")
        pdf.cell(w=45, h=8, txt=str(row["amount_purchased"]), border=1, align="C")
        pdf.cell(w=35, h=8, txt=str(row["price_per_unit"]), border=1, align="C")
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, align="C", ln=1)
        total += row["total_price"]

    """
    Here we create the last row in the table, which contains no information in its columns except from the total amount
    so we will leave those cells empty
    """

    pdf.cell(w=30, h=8, border=1, align="C")
    pdf.cell(w=50, h=8, border=1, align="C")
    pdf.cell(w=45, h=8, border=1, align="C")
    pdf.cell(w=35, h=8, border=1, align="C")
    pdf.cell(w=30, h=8, txt=str(total), border=1, align="C", ln=1)  # Here we use the total variable we calculated

    """
    Here we create a summary line and business signature
    """

    pdf.set_font(family="Times", style="B", size=16)
    pdf.set_text_color(0, 0, 0)  # Black
    pdf.cell(w=0, h=12, txt=f"Total due amount is {total} ILS", border=0, ln=1, align="l")

    pdf.set_font(family="Times", style="B", size=12)
    pdf.cell(w=32, h=13, txt=f"The Yarin Depot")
    pdf.image('depot.jpeg',w=12)
    pdf.output(f"PDF's/{filename}.pdf")
