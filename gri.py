
import re 
import pandas as pd 
import pdfplumber as pr
import requests
# 設置 PDF 檔案的下載鏈接和檔案名稱
company_code = "1310"
year = "110"

url = "https://mops.twse.com.tw/server-java/FileDownLoad?step=9&filePath=/home/html/nas/protect/t100/&fileName=t100sa11_" + company_code + "_" + year + ".pdf"
filename = "t100sa11_" + company_code + "_" + year + ".pdf"

# 發送 GET 請求並下載檔案
response = requests.get(url)
with open(filename, "wb") as f:
    f.write(response.content)

pdf = pr.open(filename)
ps = pdf.pages
df_new=[]
 # 取得總頁數
total_pages = len(pdf.pages)
# 設定起始頁數
start_page = int(total_pages * 0.8)
    
# 迭代每頁，從設定的起始頁開始
for page in pdf.pages[start_page:]:
    tables = page.extract_tables()
    for df in tables:
        df = pd.DataFrame(df, columns=df[0])
        if df.iloc[0,0] != '': 
            #df=df.drop_duplicates()
            for i in df.columns:
                filtered_columns = [col for col in df.columns if col is not None]
                for col in filtered_columns:
                    if  ("GRI" ) in col:
                            df_new.append(df) 

merged_df = pd.concat(df_new, axis=0)
merged_df

excel_name=filename+".xlsx"
output_excel=merged_df.to_excel(excel_name, index=False)
