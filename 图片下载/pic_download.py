import pandas as pd
import requests


def cut_str(str):
    return str.split('/')[5]

df = pd.read_excel(r"C:\Users\ASUS\Desktop\无标题.xlsx")
print(df)
df['img'] = df['imgs'].str.split(',')
df1 = df.explode('img')
df1['pic_name'] = df1['img'].apply(lambda x:cut_str(x))
print(df1['name'])
df1['img_name'] = df1['code'] + '_' + df1['name'] + '_' + df1['pic_name']
img_list = df1['img'].tolist()
filename_list = df1['img_name'].tolist()


def download_imgs(img_url,filename):

    try:
        r = requests.get(img_url)
        if r.status_code == 200:
            with open(r"C:\Users\ASUS\Desktop\图片" + "\\" + filename,"wb") as f:
                f.write(r.content)
    except:
        print("第"+ str(i)  +"行下载失败")



for i in range(len(img_list)):
    print(img_list[i],filename_list[i])
    download_imgs(img_url=img_list[i],filename=filename_list[i])





