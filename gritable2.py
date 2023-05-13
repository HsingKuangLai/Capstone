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


 ##Step2table_settings={"horizontal_strategy": "text"}
 pdf = pr.open(filename)
 ps = pdf.pages
 df_new=pd.DataFrame()
 df_fil=pd.DataFrame()
 pure_df=[]
 new_page=[]
 columns_count = []
 total_pages = len(pdf.pages)
 start_page = int(total_pages * 0.8)
    

#  for page in pdf.pages[start_page:]:
#      text = page.extract_text()
#       if ("GRI" and "頁" )in text:
#           new_page.append(page)
 for page in pdf.pages[start_page:]:
    tables = page.extract_tables()
    for df in tables:
      pure_df=[]
      pattern="\d{3}[\-－]\d{1,2}"
      check=False
      for row in df:
         if any(item is not None and re.search(pattern, str(item)) for item in row):
          colcount=len(row)
          pure_df.append(row)
          check=True
         else:
          continue
      if check is True:   
        df1= pd.DataFrame(pure_df, columns=range(colcount))
        df_fil=pd.DataFrame(page_col(df1))
      if df_fil.shape[1]==2: 
        df_new=pd.concat([df_new,df_fil], ignore_index=True) ##全GRI
        df_new=df_new.drop_duplicates()

 pattern = "\d{3}[\-－]\d{1,2}"

 ##Step5
 for i in range(df_new.shape[0]):
     if re.search(pattern, str(df_new.iloc[i, 0]), flags=re.MULTILINE):
         df_new.iloc[i, 0] = (re.findall(pattern, str(df_new.iloc[i, 0]), flags=re.MULTILINE))[0]
     else:
         for j in range(df_new.shape[1]):
             df_new.iloc[i, j] = ''

 df_new=df_new.dropna(how="all")
 output_excel='outputfiles/'+company_code+'_'+year_code+'_etr.xlsx'
 df_new.to_excel(output_excel,index=False,header=False,sheet_name='工作表1')




def page_col(df):
    import re
    import pandas as pd
    page_columns = []
    original_column_names=[]
    pattern_all = r'^[^A-Za-z\u4e00-\u9fa5]*$|^頁數$|^頁碼$|^頁次$' 
    pattern_gri = r"\d{3}[\-－]\d{1,2}"
    df = df.fillna('')
    for column in df.columns:
        column_values = df[column].astype(str)
        if (column_values.str.match(pattern_all) | column_values.str.findall(pattern_gri).apply(lambda x: len(x) > 0)).all() and (not column_values.str.strip().eq('').all() and not column_values.eq('None').all()):
          page_columns.append(column)
          original_column_names.append(column)      
    subset_df = df[page_columns]
    if len(subset_df.columns)>2:
       subset_df = subset_df.iloc[:, :2]
    renamed_df = subset_df.rename(columns=dict(zip(original_column_names, ['GRI', 'Pages'])))
    return renamed_df





gritable('1712','110')
test_elect=['1708','1709','1710','1711','1712','1713','1714','1717','1718','1721','1722','1723','1725','1726','1727','1730','1732','1735','1742','1773','1776','3430','3708']
test_elect2=['4702','4706','4707','4711','4714','4716','4720','4721','4722','4739','4741','4754','4755','4763','4764','4766','4767','4768','4770','6509']
test_elect3=['2101','2102','2103','2104','2105','2106','2107','2108','2109','2114','6582']
test_elect4=['1101','1102','1103','1104','1108','1109','1110']


# for firm in test_elect :
#   try:
#      gritable(firm,'110')
#   except:
#      pass

