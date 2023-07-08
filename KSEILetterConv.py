import requests
import pdfplumber
import pandas as pd
import numpy as np

# Function to extract tables from PDF
def extract_tables_from_pdf(pdf_link, output_file):
    # Download the PDF file
    response = requests.get(pdf_link)
    with open('temp.pdf', 'wb') as f:
        f.write(response.content)
    
    # Open the PDF
    with pdfplumber.open("temp.pdf") as pdf:
        # Extract tables from each page
        data_frames = []
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                df = pd.DataFrame(table[1:], columns=table[0])
                # Clean up specific columns for the first two tables
                if len(data_frames) < 2:
                    if df.columns[1] in df:
                        df[df.columns[1]] = df[df.columns[1]].str.replace("Rp", "", regex=True).str.replace(" ", "", regex=True).str.replace(",-", "", regex=True).str.replace(".", "", regex=True)
                        df[df.columns[1]] = pd.to_numeric(df[df.columns[1]], errors='coerce')
                    if df.columns[2] in df:
                        df[df.columns[2]] = df[df.columns[2]].str.replace("%", "", regex=True).str.replace(" ", "", regex=True).str.replace("p.a", "", regex=True)
                        df[df.columns[2]] = pd.to_numeric(df[df.columns[2]], errors='coerce')
                data_frames.append(df)
    
    # Write data frames to Excel file
    with pd.ExcelWriter(output_file) as writer:
        data_frames[0].to_excel(writer, sheet_name='Obligasi', index=False)
        data_frames[1].to_excel(writer, sheet_name='Sukuk', index=False)
        data_frames[2].to_excel(writer, sheet_name='Agenda Penerbitan', index=False)

# Use the function
pdf_link = "https://www.ksei.co.id/Announcement/Files/156611_ksei_1859_dir_0723_202307031943.pdf"
output_file = "D:/datapublikasi.xlsx"
extract_tables_from_pdf(pdf_link, output_file)
