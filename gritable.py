import re 
import pandas as pd 
import pdfplumber as pr
import requests
# 設置 PDF 檔案的下載鏈接和檔案名稱
company_code = "2886"
year = "110"

url = "https://mops.twse.com.tw/server-java/FileDownLoad?step=9&filePath=/home/html/nas/protect/t100/&fileName=t100sa11_" + company_code + "_" + year + ".pdf"
filename = "t100sa11_" + company_code + "_" + year + ".pdf"

# 發送 GET 請求並下載檔案
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
    
# 迭代每頁，從設定的起始頁開始
for page in pdf.pages[start_page:]:
    text = page.extract_text()
    if ("GRI" and "頁" )in text:
        new_page.append(page)
        for page in new_page:
            tables = page.extract_tables()
            for df in tables:
                df1= pd.DataFrame(df, columns=df[0])
                df1=df1.drop_duplicates()
                df_new.append(df1) 


##step3
UR=[]
final=[]
pattern="\d{1,3}[\-－]\d{1,2}"

for df in df_new:
    if df.shape[1] >= 4:
        final.append(df)
        for df in final:
            # 检查当前DataFrame对象的前两列是否与正则表达式匹配
            
            col1 = df.iloc[:, 0]
            col2 = df.iloc[:, 1]
            
            matches = col1.str.contains(pattern) | col2.str.contains(pattern)

            # 如果至少有一个单元格与正则表达式匹配，则将该DataFrame添加到UR列表中
            if matches.any():
                UR.append(df)


##Step4
# 對於每個DataFrame，只保留包含指定模式的列
new_df=pd.DataFrame()
for df in UR:
    pattern = r"\d{1,3}[\-－]\d{1,2}"
    has_page_condition = lambda x: x.str.contains("頁")
    # 使用 str.contains() 方法和 `|` 操作符組合篩選條件
    filtered_df = df.loc[:, df.apply(lambda x: x.str.contains(pattern) | has_page_condition(x)).any()]
    new_df = pd.concat([new_df, filtered_df], axis=0)

new_df=new_df.reset_index(drop=True)
new_df=new_df.drop_duplicates()
new_df=new_df.rename(columns={new_df.columns[1]:"pages"})

pattern="\d{3}-\d{1,2}"

##Step5
for i in range(new_df.shape[0]):
    # 檢查第一欄是否符合正則表達式(pattern)
    if re.match(pattern, str(new_df.iloc[i, 0])):
        # 第一欄符合條件，不作處理
        new_df.iloc[i, 0] = (re.findall(pattern, str(new_df.iloc[i, 0]))[0])
    else:
        # 第一欄不符合條件，清空該欄
        # 清空該列的其他欄位
        for j in range(new_df.shape[1]):
            new_df.iloc[i, j] = ''

new_df=new_df.dropna(how="any")
new_df=new_df.drop_duplicates()
output_excel=filename+'.xlsx'
new_df.to_excel(output_excel,index=False,header=False,sheet_name=company_code)