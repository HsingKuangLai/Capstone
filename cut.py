def cut(company_code, year_code):
 # Read in the Excel file as a dataframe

 import pandas as pd
 import numpy as np
 import openpyxl
 import os
 import re
 from openpyxl import load_workbook
 from pdfrw import PdfReader, PdfWriter

 output_excel = 'outputfiles/' + company_code + '_' + year_code + '_etr.xlsx'
 if not os.path.exists(output_excel):
   output_excel = 'outputfiles/' + company_code + '_' + year_code + '_etr.xlsm'
 
 filename='inputfiles/'+company_code+'_'+year_code+'.pdf'

 df = pd.read_excel(output_excel, engine='openpyxl')
 df

 for i in range(df.shape[0]):
     # Check if the pages column contains a string
     #if isinstance(df.loc[i, 'pages'], str):
         # Extract digits from the pages string
         digits = re.findall(r'\d+', str(df.iloc[i, 1]))
         pages = list(map(int, digits))

         # Generate the corresponding PDF file
         output_file = os.path.join('outputfiles/'+company_code+'_'+year_code+'_'+str(df.iloc[i, 0]) + '.pdf')
         writer = PdfWriter()
         for page in pages:
             writer.addpage(PdfReader(filename).pages[page-1])
         writer.write(output_file)
         
         # Update the pdf_path column in the DataFrame
         df.loc[i, 'pdf_path'] = company_code+'_'+year_code+'_'+str(df.iloc[i, 0]) + '.pdf'
 output_excel = 'outputfiles/' + company_code + '_' + year_code + '_etr.xlsx'
 df.to_excel(output_excel,index=False,sheet_name=company_code)



