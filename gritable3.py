def gritable(company_code, year_code):
 import re 
 import pandas as pd 
 import pdfplumber as pr
 import requests
 import time
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
 start_page = int(total_pages * 0.85)


#Step3table_settings={"horizontal_strategy": "text"}'[^\x00-\x7F]|(?![Pp])[A-Za-z]'=r'[^\x00-\x7F](、)'r'[^\x00-\x7F、]+|\d+\.\d+'SS
 for page in pdf.pages[start_page:]:
    tables = page.extract_tables()
    for df in tables:
      pure_df=[]
      pattern="\d{3}[\-－]\d{1,2}"
      pattern_nascll=r'[^\x01-\x7F、]|(?![Pp])[A-Za-z]+|\d+\.\d+'

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

#電子零組
# test_elect=['1582','2059','2308','2313','2327','2328','2367','2368','2383','2385','2392','2420','2421','2457','2484','2492','3015','3023','3026','3037','3042','3044','3376','3533','3605','3645','3715','4915','4927','4958','5469','6153','6191','6213','6269','6282','6412','8039','8046','1336','3206','3236','3294','3357','3388','3624','5227','5309','5457','6173','6207','6208','6220','6274','6279','6284','6538','6664','8074','8121','8182']
#化工
# test_elect2=["1708","1709","1710","1711","1712","1713","1714","1717","1718","1721","1722","1723","1725","1726","1727","1730","1732","1735","1773","1776","3708","4720","1742","4702","4706","4707","4711","4714","4716","4721","4741","4754","4767","6509","4722","4739","4755","4763","4764","4766","4770"]
# 電腦"3706","3712","4938","6128","6166","6206","6277","6414","6579","6669","3088","3211","3272","3594","4931",
# test_elect3=["2103","2106","2108","2109","6582"]"1301","1303","1304","1305","1308","1310","1312","1313","1314","1326",
# "2323","2349","2409","2489","3049","3481","3576","2301","2324","2331","2352","2353","2356","2357","2362","2376","2377","2382","2395","3005","3231","3706",
# "2317","2354","1785","1103","1104","2329","2337","2344","2401","6147","6182","2102","1305","2349","2409","3576","6120","8069"
# test_elect4=["2344","2401","6147","6182","2102","1305","2349","2409","3576","6120","8069"]

# for firm in test_elect4 :
#   try:
#      gritable(firm,'109')
#   except:
#       pass
gritable('2883','110')