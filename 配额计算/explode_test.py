import pandas as pd

df = pd.read_clipboard()

list = ['AP1', 'AP2', 'AP3', 'AP4', 'NC1', 'NC2', 'NC3', 'NC4', 'AC1', 'AC2', 'AC3', 'AC4', 'AP1MINI', 'AP2MINI', 'AP3MINI']

sku = [list for i in range(df.shape[0])]


def ex():
    df['sku'] = pd.DataFrame({'sku': sku})
    df1 = df.explode('sku')
    print(df1)
    df1.to_excel(r'C:\Users\ASUS\Desktop\mdm_sku.xlsx', index=False)


ex()
