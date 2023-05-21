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
 df_new=pd.DataFrame()
 df_fil=pd.DataFrame()
 pure_df=[]
 new_page=[]
 columns_count = []
 total_pages = len(pdf.pages)
 start_page = int(total_pages * 0.8)


#Step3table_settings={"horizontal_strategy": "text"}'[^\x00-\x7F]|(?![Pp])[A-Za-z]'
 for page in pdf.pages[start_page:]:
    tables = page.extract_tables()
    for df in tables:
      pure_df=[]
      pattern="\d{3}[\-－]\d{1,2}"
      pattern_nascll=r'[^\x00-\x7F](、)'
      check=False
      for row in df:
         if any(item is not None and re.search(pattern, str(item), flags=re.MULTILINE) for item in row):
          row = ['' if item is not None and re.search(pattern, str(item), flags=re.MULTILINE) is None and bool(re.search(pattern_nascll, str(item), flags=re.MULTILINE)) is True else item for item in row]
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

#r'^[^A-Za-z\u4e00-\u9fa5]*$|^頁數$|^頁碼$|^頁次$' 
def page_col(df):
    import re
    import pandas as pd
    page_columns = []
    pattern_all = r'^[^\u4e00-\u9fa5]*$|^頁數$|^頁碼$|^頁次$' 
    pattern_gri = r"\d{3}[\-－]\d{1,2}"
    df = df.fillna('')
    top_two_columns=[]
    top_two_columns_list=[]
    for column in df.columns:
        column_values = df[column].astype(str)
        if (column_values.str.match(pattern_all,flags=re.MULTILINE) | column_values.str.findall(pattern_gri).apply(lambda x: len(x) > 0)).all() and (not column_values.str.strip().eq('').all() and not column_values.eq('None').all()):
          page_columns.append(column)
          top_two_columns_list.append(column) 
    subset_df = df[page_columns]
    if len(subset_df.columns)>2:
      count_df = subset_df.apply(lambda x: x[x != ''].count())
      top_two_columns = count_df.nlargest(2).index
      # concatenated =subset_df.apply(lambda x: ''.join(x.astype(str)), axis=0)
      # top_two_columns = concatenated.str.len().nlargest(2).index
      subset_df = subset_df[top_two_columns]
      top_two_columns_list=top_two_columns.tolist()
    renamed_df = subset_df.rename(columns=dict(zip(top_two_columns_list, ['GRI', 'Pages'])))
    return renamed_df

# test_elect=['1708','1709','1710','1711','1712','1713','1714','1717','1718','1721','1722','1723','1725','1726','1727','1730','1732','1735','1742','1773','1776','3430','3708']
# #
# test_elect2=['4702','4706','4707','4711','4714','4716','4720','4721','4722','4739','4741','4754','4755','4763','4764','4766','4767','4768','4770','6509']
# test_elect3=['2101','2102','2103','2104','2105','2106','2107','2108','2109','2114','6582']
# test_elect4=['1101','1102','1103','1104','1108','1109','1110']

# for firm in test_elect4 :
#   try:
#      gritable(firm,'110')
#   except:
#      pass
# gritable('1201','110')