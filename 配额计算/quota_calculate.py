import pandas as pd
import pymysql
import tkinter as tk
from tkinter import filedialog
import math
import time


def get_conn():
    #建立连接
    conn = pymysql.connect(
        host='host',
        user='user',
        password='password',
        db='db',
        charset='utf8'
    )
    return conn

def get_path():
    #选择文件路径
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askopenfilename()
    return path

def df_allocation_mdm(conn=get_conn()):
    sql = "SELECT `Code`,`Name`,CustGroup_Name,Distcode_Code,Distcode_Name,CustomerClass,SupplyType_Name FROM `user_retailer_missing_mdm_copy`"
    df_mdm = pd.read_sql(sql, conn)
    return df_mdm

def df_allocation(path=get_path()):
    df_all = pd.read_excel(path)
    return df_all


def allocation_match(df_all, df_mdm):
    #MDM码列转为字符串并对7位的补0
    df_all['customer_code'] = df_all['customer_code'].astype(str)
    df_all['customer_code'] = df_all['customer_code'].apply(lambda x: str(x).zfill(8))

    #删除有效性不为1的数据
    df_all.drop(df_all[df_all.validity != 1].index, inplace=True)
    #TIN列向上取整
    df_all.tin = df_all.tin.apply(lambda x: math.ceil(x))
    #重新匹配字段
    df_rematch_allocation = pd.merge(df_all, df_mdm, left_on='customer_code', right_on='Code', how='left')
    list_drop = ['account_name', 'customer_name', 'distributor_code', 'distributor_name', 'customer_type', 'Code']
    #删除列
    for col in list_drop:
        df_rematch_allocation.drop([col], axis=1, inplace=True)
    #调整列的顺序
    col_order = ['data_source', 'year', 'month', 'channel_code', 'grade_group_code', 'grade_code', 'CustGroup_Name',
                 'customer_code', 'Name', 'Distcode_Code', 'Distcode_Name', 'CustomerClass', 'sku', 'tin', 'SupplyType_Name']
    df_rematch_allocation = df_rematch_allocation[col_order]
    return df_rematch_allocation


if __name__ == '__main__':
    start = time.time()
    allocation_match(df_all=df_allocation(), df_mdm=df_allocation_mdm()).to_excel(r'C:\Users\ASUS\Desktop\allocation.xlsx', index=False)
    end = time.time()
    print('运行时间: %s秒' % (end - start))
