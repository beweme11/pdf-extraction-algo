import win32com.client
import os
from docx.api import Document
import pandas as pd
import csv

# INPUT/OUTPUT PATH
pdf_path = "pdfs\\Xometry_DesignGuide_InjectionMolding.pdf"
pdf_directory = "D:\\all codes\\code\\pdfs"
output_path = "D:\\all codes\\code\\outputs"
word = win32com.client.Dispatch("Word.Application")
word.visible = 0

def conversion(pdf_path):
    # Conversion
    filename = pdf_path.split('\\')[-1]
    in_file = os.path.abspath(pdf_path)
    wb = word.Documents.Open(in_file)
    out_file = os.path.abspath(output_path + '\\' + filename[0:-4] + ".docx")
    wb.SaveAs2(out_file, FileFormat=16)
    wb.Close()
    # Extraction
    print(f"Working on file {filename} ...")
    document = Document(out_file)
    tables = document.tables
    total_tables = 0
    for table_index, table in enumerate(tables, start=1):
        total_tables += 1

        table_data = []
        for row in table.rows:
            row_data = [cell.text.strip().replace('-', ' to ') for cell in row.cells] 
            table_data.append(row_data)

        df = pd.DataFrame(table_data)
        csv_file_path = f"D:\\all codes\\code\\outputs\\{filename}\\table_{table_index}.csv"
        os.makedirs(os.path.dirname(csv_file_path), exist_ok=True)  # Create directory if it doesn't exist
        df.to_csv(csv_file_path, index=False, header=False)

    print(f"Total number of tables extracted: {total_tables}")

#for filename in os.listdir(pdf_directory):
  #  if filename.endswith(".pdf"):
   #     pdf_path = os.path.join(pdf_directory, filename)
conversion(pdf_path)
