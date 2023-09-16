import pandas as pd 
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob('invoices/*.xlsx')

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name='Sheet 1')
    pdf = FPDF(orientation='P', unit='mm', format='A4')

    filename = Path(filepath).stem

    pdf.add_page()

    # Added Date and Title

    pdf.set_font(family='Times', size=16, style ='B')
    pdf.cell(w=50, h=8, txt=f'Invoice nr. {filename.split("-")[0]} ', ln=1)

    pdf.set_font(family='Times', size=16, style ='B')
    pdf.cell(w=50, h=8, txt=f'Date: {filename.split("-")[1]} ', ln=1)

    #Added Column names

    headers = list(df.columns)
    headers = [item.replace('_', ' ').title() for item in headers]

    pdf.set_font(family='Times', size=10, style='B')
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=30, h=8, txt= headers[0], border=1)
    pdf.cell(w=70, h=8, txt= headers[1], border=1)
    pdf.cell(w=35, h=8, txt= headers[2], border=1)
    pdf.cell(w=30, h=8, txt= headers[3], border=1)
    pdf.cell(w=30, h=8, txt= headers[4], border=1, ln=1)


#Added Items in Table

    for index, row in df.iterrows():
        pdf.set_font(family='Times', size=10, style='B')
        pdf.set_text_color(0, 0, 0)
        pdf.cell(w=30, h=8, txt= str(row['product_id']), border=1)
        pdf.cell(w=70, h=8, txt= str(row['product_name']), border=1)
        pdf.cell(w=35, h=8, txt= str(row['amount_purchased']), border=1)
        pdf.cell(w=30, h=8, txt= str(row['price_per_unit']), border=1)
        pdf.cell(w=30, h=8, txt= str(row['total_price']), border=1, ln=1)

#Added total sum    

    total_sum = df['total_price'].sum()
    pdf.set_font(family='Times', size=10, style='B')
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=30, h=8, txt= '', border=1)
    pdf.cell(w=70, h=8, txt= '', border=1)
    pdf.cell(w=35, h=8, txt= '', border=1)
    pdf.cell(w=30, h=8, txt= '', border=1)
    pdf.cell(w=30, h=8, txt= str(total_sum), border=1, ln=1)

#Added Total sum line and company name

    pdf.set_font(family='Times', size=12, style='B')
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=90, h=8, txt=f'The total price is {total_sum} pounds', align='L', ln=1)

    pdf.set_font(family='Times', size=12, style='B')
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=90, h=8, txt=f'ABC Company', align='L')

    pdf.output(f'Outputs/{filename}.pdf')


