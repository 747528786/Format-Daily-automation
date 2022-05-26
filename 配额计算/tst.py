import pandas as pd
import json
df = pd.read_excel(r'C:\Users\ASUS\Desktop\4月至今GT渠道已完结工作流.xlsx')
result = df.changed_data.apply(lambda x:json.loads(x)).apply(pd.Series)
df_result = pd.merge(df,result,left_index=True,right_index=True)
df_result.to_excel(r'C:\Users\ASUS\Desktop\4月至今GT渠道已完结工作流5.18.xlsx')