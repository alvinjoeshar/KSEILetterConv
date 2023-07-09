import requests
import pdfplumber
import pandas as pd
import re
from tqdm import tqdm

def print_ascii(ascii_art):
    for line in ascii_art.splitlines():
        print('\033[92m' + ''.join([chr(c) if c.isalpha() else c for c in line]) + '\033[0m')

if __name__ == '__main__':
    ascii_art = """
    .##....##..######..########.####..######...#######..##....##.##.....##
    .##...##..##....##.##........##..##....##.##.....##.###...##.##.....##
    .##..##...##.......##........##..##.......##.....##.####..##.##.....##
    .#####.....######..######....##..##.......##.....##.##.##.##.##.....##
    .##..##.........##.##........##..##.......##.....##.##..####..##...##.
    .##...##..##....##.##........##..##....##.##.....##.##...###...##.##..
    .##....##..######..########.####..######...#######..##....##....###...
    """
   
    print_ascii(ascii_art)
    print('\033[92m' + "v.01 by AlvinLee" + '\033[0m')
    print('\033[92m' + "https://github.com/alvinjoeshar" + '\033[0m')
    print("")
    print("")

import time

def extract_tables_from_pdfs(pdf_links, output_file):
    obligasi_dfs = []
    sukuk_dfs = []
    timeline_dfs = []

    pdf_tqdm = tqdm(pdf_links, desc="\033[92mProcessing PDFs\033[0m")
    for i, pdf_link in enumerate(pdf_tqdm):
        response = requests.get(pdf_link)
        with open('temp.pdf', 'wb') as f:
            f.write(response.content)

        with pdfplumber.open("temp.pdf") as pdf:
            page_tqdm = tqdm(pdf.pages, desc=f"\033[92mProcessing pages of PDF {i+1}\033[0m")
            data_frames = []
            for page in page_tqdm:
                tables = page.extract_tables()
                for table in tables:
                    df = pd.DataFrame(table[1:], columns=table[0])
                    if len(data_frames) < 2:
                        if 'Jumlah PokokSebesar' in df.columns:
                            df['Jumlah PokokSebesar'] = df['Jumlah PokokSebesar'].str.replace("Rp", "", regex=True).str.replace(" ", "", regex=True).str.replace(",-", "", regex=True).str.replace(".", "", regex=True)
                            df['Jumlah PokokSebesar'] = pd.to_numeric(df['Jumlah PokokSebesar'], errors='coerce')
                        if 'Jenis &Tingkat Bunga(Tetap)' in df.columns:
                            df['Jenis &Tingkat Bunga(Tetap)'] = df['Jenis &Tingkat Bunga(Tetap)'].str.replace("%", "", regex=True).str.replace(" ", "", regex=True).str.replace("p.a", "", regex=True)
                            df['Jenis &Tingkat Bunga(Tetap)'] = pd.to_numeric(df['Jenis &Tingkat Bunga(Tetap)'], errors='coerce')
                        if 'Jumlah Dana Sukuksebesar' in df.columns:
                            df['Jumlah Dana Sukuksebesar'] = df['Jumlah Dana Sukuksebesar'].str.replace("Rp", "", regex=True).str.replace(" ", "", regex=True).str.replace(",-", "", regex=True).str.replace(".", "", regex=True)
                            df['Jumlah Dana Sukuksebesar'] = pd.to_numeric(df['Jumlah Dana Sukuksebesar'], errors='coerce')
                    data_frames.append(df)
            page_tqdm.close()  # Manually close the progress bar

            first_page = pdf.pages[0]
            text = first_page.extract_text()

            issuer_pattern = r"PT\s(.*?)(\n|$)"
            bond_pattern = r"Obligasi\s(.*?)(\n|$)"
            sukuk_pattern = r"Sukuk\s(.*?)(\n|$)"

            issuer_name = re.search(issuer_pattern, text).group(1) if re.search(issuer_pattern, text) else None
            bond_name = re.search(bond_pattern, text).group(1) if re.search(bond_pattern, text) else None
            sukuk_name = re.search(sukuk_pattern, text).group(1) if re.search(sukuk_pattern, text) else None

            if issuer_name is not None:
                for df in data_frames:
                    df.insert(0, 'Issuer Name', issuer_name)

            if bond_name is not None and sukuk_name is not None:
                if len(data_frames) >= 1:
                    data_frames[0].insert(1, 'Bond Name', bond_name)
                    obligasi_dfs.append(data_frames[0])
                if len(data_frames) >= 2:
                    data_frames[1].insert(1, 'Sukuk Name', sukuk_name)
                    sukuk_dfs.append(data_frames[1])
                if len(data_frames) >= 3:
                    timeline_dfs.append(data_frames[2])
            elif bond_name is not None:
                if len(data_frames) >= 1:
                    data_frames[0].insert(1, 'Bond Name', bond_name)
                    obligasi_dfs.append(data_frames[0])
                if len(data_frames) >= 2:
                    timeline_dfs.append(data_frames[1])
            elif sukuk_name is not None:
                if len(data_frames) >= 1:
                    data_frames[0].insert(1, 'Sukuk Name', sukuk_name)
                    sukuk_dfs.append(data_frames[0])
                if len(data_frames) >= 2:
                    timeline_dfs.append(data_frames[1])
    pdf_tqdm.close()  # Manually close the progress bar

    for df in obligasi_dfs:
        df.columns = obligasi_dfs[0].columns

    for df in timeline_dfs:
        if len(df.columns) >= 4:
            df.columns = ["Issuer Name", "Keterangan", ":", "Tanggal"] + list(df.columns[4:])
        else:
            print("Warning: Timeline dataframe has less than 4 columns. Column names have not been changed.")

    with pd.ExcelWriter(output_file) as writer:
        if obligasi_dfs:
            pd.concat(obligasi_dfs, ignore_index=True).to_excel(writer, sheet_name='Obligasi', index=False)
        if sukuk_dfs:
            pd.concat(sukuk_dfs, ignore_index=True).to_excel(writer, sheet_name='Sukuk', index=False)
        if timeline_dfs:
            pd.concat(timeline_dfs, ignore_index=True).to_excel(writer, sheet_name='Timeline', index=False)

    print(" ")
    print('\033[92m' + "Compilation completed successfully." + '\033[0m')
    print(" ")

pdf_links = []
while True:
    pdf_link = input("Enter a PDF link, or 'done' to finish: ")
    if pdf_link.lower() == 'done':
        break
    pdf_links.append(pdf_link)

output_file = "D:/datapublikasi.xlsx"
extract_tables_from_pdfs(pdf_links, output_file)
