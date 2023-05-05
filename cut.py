# Read in the Excel file as a dataframe

import pandas as pd
import numpy as np
import openpyxl
import os
from openpyxl import load_workbook
from pdfrw import PdfWriter,PdfReader


excel_name="C://Users//star8//Desktop//OP-copy//OP//2881//2881_110_etr.xlsx"
filename="C://Users//star8//Desktop//OP-copy//IP//2881_110.pdf"

df = pd.read_excel(excel_name, engine='openpyxl')
df

# Loop through each row in the DataFrame
for i in range(df.shape[0]):
    # Split the pages string into individual page numbers
    pages = list(map(int, df.loc[i, 'pages'].split(',')))
    
    # Generate the corresponding PDF file
    output_file = os.path.join("outputfiles/", str(i+1) + '.pdf')
    writer = PdfWriter()
    for page in pages:
        writer.addpage(PdfReader(filename).pages[page-1])
    writer.write(output_file)
    
    # Update the pdf_path column in the DataFrame
    df.loc[i, 'pdf_path'] = output_file