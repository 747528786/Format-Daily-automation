import pandas as pd
import pymysql
import tkinter as tk
from tkinter import filedialog
import math
import time
import numpy as np


class QuotaError(Exception):
    pass


def get_conn():
    # 建立连接
    conn = pymysql.connect(
        host='rm-uf6lxxf35108k7uq19o.mysql.rds.aliyuncs.com',
        user='danego_ro',
        password='D@none20181024',
        db='danego',
        charset='utf8'
    )
    return conn


def get_path(file_name):
    # 选择文件路径
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askopenfilename(title=file_name)
    return path


def get_df_retailer():
    """
    门店列表
    :return: 门店列表.csv
    """
    path = get_path('请选择门店列表')
    df_retailer = pd.read_csv(path, usecols=['门店ID', '门店名称', '达能门店编码', '渠道', '门店等级', '门店状态', '父级门店ID', '关联经销商', '经销商编码',
                                             '注册地址-省', '是否是羊滋滋门店'])
    # 父级门店转为数值，如有文本格式无法转换则为Nan
    df_retailer['父级门店ID'] = pd.to_numeric(df_retailer['父级门店ID'], errors='coerce')
    # 门店编码列处理，去除=“ ”
    df_retailer['达能门店编码'] = df_retailer['达能门店编码'].str.split('="').str[1].str.split('"').str[0]
    # 去除MDM列的重复数据，保留第一个
    df_retailer.drop_duplicates(keep='first', subset='达能门店编码', inplace=True)
    return df_retailer


def get_df_guide():
    """
    共享导购列表
    :return: mdm_code & guide_phone
    """
    sql = '''
    SELECT t2.mdm_code,t1.guide_phone FROM `user_guide` t1
    left join user_retailer t2
    on t1.retailer_id = t2.id
    '''
    df_guide = pd.read_sql(sql, con=get_conn())
    return df_guide


def get_df_order():
    """
    订单表
    :return: mdm_code & sku & goods_num & retailer_type & order_time
    """
    try:
        path = get_path('请选择订单表')
        df_order = pd.read_excel(path)
        df_order['mdm_code'] = df_order['mdm_code'].str.upper()
        df_order['order_time'] = pd.to_datetime(df_order['order_time'])
        df_order['month'] = df_order['order_time'].dt.month
        return df_order
    except:
        raise QuotaError('打开订单表')


def match_main_store(df_order, df_retailer):
    """
    匹配父级
    :return:
    """
    df1 = df_order
    df2 = df_retailer[['门店ID', '达能门店编码', '父级门店ID']]
    df = pd.merge(df1, df2, left_on='mdm_code', right_on='达能门店编码', how='left')
    # 根据门店类型匹配父级
    d = (df['retailer_type'] == '总仓配货门店').values
    f = list(enumerate(d))
    for i in f:
        if i[1]:
            df.loc[i[0], '门店ID'] = df.loc[i[0], '父级门店ID']
    # 删除不需要的列
    col_order = ['门店ID', 'mdm_code', 'sku', 'goods_num', 'retailer_type', 'month']
    for col in list(df):
        if col not in col_order:
            df.drop(columns=col, inplace=True)
    # 调整列的顺序
    df = df[col_order]
    # df['门店ID'] = df['门店ID'].fillna('#NA')
    df_piv = pd.pivot_table(df, index=['门店ID'], columns='month', values='goods_num', aggfunc=np.sum)
    df_result = pd.merge(df_piv, df_retailer, on='门店ID', how='left', )
    return df_result


if __name__ == '__main__':
    match_main_store(df_order=get_df_order(), df_retailer=get_df_retailer()).to_excel(
        r'C:\Users\ASUS\Desktop\test.xlsx')
