def gritable(company_code, year_code):
 import re 
 import pandas as pd 
 import pdfplumber as pr
 import requests
 #import requests
 # 設置 PDF 檔案的下載鏈接和檔案名稱
 
 url = "https://mops.twse.com.tw/server-java/FileDownLoad?step=9&filePath=/home/html/nas/protect/t100/&fileName=t100sa11_" + company_code + "_" + year_code + ".pdf"
 filename = "inputfiles/" + company_code + "_" + year_code + ".pdf"

  #發送 GET 請求並下載檔案
 response = requests.get(url)
 with open(filename, "wb") as f:
    f.write(response.content)


 ##Step2
 pdf = pr.open(filename)
 ps = pdf.pages
 df_new=[]
 new_page=[]
 columns_count = []
  # 取得總頁數
 total_pages = len(pdf.pages)
 # 設定起始頁數
 start_page = int(total_pages * 0.8)
    
 # 迭代每頁，從設定的起始頁開始if any(("頁") not in elem for elem in df[i]):
 for page in pdf.pages[start_page:]:
     text = page.extract_text()
     if ("GRI" and "頁" )in text:
         new_page.append(page)
 for page in new_page:
    tables = page.extract_tables()
    for df in tables:
     try:
      col=[]
      for i in range(2):
         concatenated_string = ''.join(item for item in df[i] if item is not None) 
         if ("頁") not in concatenated_string:
          pass
         else:
          df[i]=['空白' if item is None else item for item in df[i]]
          col=df[i]
          break
      if len(col)>0:      
       df1= pd.DataFrame(df, columns=col)
       df_new.append(df1)
      else:
       pass 
     except:
        pass 

 UR=[]
 final=[]
 pattern="\d{1,3}[\-－]\d{1,2}"

 for df in df_new:
    if df.shape[1] >= 4:
        final.append(df)
 for df in final:
    if df.apply(lambda x: x.str.contains(pattern)).any().any():
     UR.append(df)       


 ##Step4
 # 對於每個DataFrame，只保留包含指定模式的列
 new_df=pd.DataFrame()
 for df in UR:
     pattern = r"\d{1,3}[\-－]\d{1,2}"
     has_page_condition = lambda x: x.str.contains(r"頁碼|頁數|頁", flags=re.IGNORECASE)
     # 使用 str.contains() 方法和 `|` 操作符組合篩選條件
     filtered_df = df.loc[:, df.apply(lambda x: x.str.contains(pattern) | has_page_condition(x)).any()]
     new_df=new_df.append(filtered_df, ignore_index=True)

 new_df=new_df.reset_index(drop=True)
 new_df=new_df.drop_duplicates()
 new_df=new_df.rename(columns={new_df.columns[1]:"pages"})

 pattern = "\d{3}[\-－]\d{1,2}"

 ##Step5
 for i in range(new_df.shape[0]):
     # 檢查第一欄是否符合正則表達式(pattern)
     if re.search(pattern, str(new_df.iloc[i, 0]), flags=re.MULTILINE):
         # 第一欄符合條件，不作處理
         new_df.iloc[i, 0] = (re.findall(pattern, str(new_df.iloc[i, 0]), flags=re.MULTILINE))[0]
     else:
        # 第一欄不符合條件，清空該欄
        # 清空該列的其他欄位
         for j in range(new_df.shape[1]):
             new_df.iloc[i, j] = ''

 new_df=new_df.dropna(how="any")
 new_df=new_df.drop_duplicates()
 output_excel='outputfiles/'+company_code+'_'+year_code+'_etr.xlsx'
 new_df.to_excel(output_excel,index=False,header=False,sheet_name='工作表1')




test_elect=['1708','1709','1710','1711','1712','1713','1714','1717','1718','1721','1722','1723','1725','1726','1727','1730','1732','1735','1742','1773','1776','3430','3708']
test_elect2=['4702','4706','4707','4711','4714','4716','4720','4721','4722','4739','4741','4754','4755','4763','4764','4766','4767','4768','4770','6509']
test_elect3=['2101','2102','2103','2104','2105','2106','2107','2108','2109','2114','6582']
test_elect4=['1101','1102','1103','1104','1108','1109','1110']


for firm in test_elect2 :
  try:
     gritable(firm,'110')
  except:
     pass

#
# gritable('2881','110')