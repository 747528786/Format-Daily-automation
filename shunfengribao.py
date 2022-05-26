import pandas as pd
import pymysql


def data_from_sql(first_date, last_date):
    conn = pymysql.connect(
        host='host',
        user='user',
        password='password',
        db='db',
        charset='utf8'
    )
    sql = 'SELECT *,sum( goods_num ) as "数量" FROM \
        ( SELECT t1.order_code,t1.retailer_id,t2.NAME,t2.mdm_code,t2.address "注册地址",t1.address "收货地址", \
        CASE t2.new_type WHEN 1 THEN "联盟总店" \
                WHEN 2 THEN "联盟分店" \
                WHEN 3 THEN "分销商总仓配货门店" \
                WHEN 4 THEN "分销商总仓和门店" \
                WHEN 5 THEN "分销商供货总仓" \
                WHEN 6 THEN "经销商" \
                WHEN 7 THEN "总仓和门店" \
                WHEN 8 THEN "总仓配货门店" \
                WHEN 9 THEN "总仓" \
                WHEN 10 THEN "直接订货门店" \
                WHEN 11 THEN "分销商供货门店" \
        END AS "门店类型", t3.goods_num ,t5.name as "省份",t6.name as "城市"\
        FROM `order_oms_info` t1\
            LEFT JOIN user_retailer t2 ON t1.retailer_id = t2.id \
            LEFT JOIN order_oms_goods t3 ON t1.order_code = t3.order_code \
            LEFT JOIN region t5 ON t2.province = t5.id \
		    LEFT JOIN region t6 ON t2.city = t6.id \
        WHERE DATE( t1.order_time ) BETWEEN "{}" and "{}") t4 \
    GROUP BY order_code'.format(first_date, last_date)
    df = pd.read_sql(sql, conn)
    conn.close()
    return df

def data_from_excel(excel_name):
    df1 = pd.read_excel(excel_name)
    df1.rename(columns={'订单号': 'order_code'}, inplace=True)
    return df1

def data_transform(df, df1):
    df2 = pd.merge(df, df1, on='order_code', how='right')
    df2.rename(columns={'order_code': '达能（EOG）订单号', 'retailer_id': '达易购ID', 'mdm_code': 'MDM码', 'NAME': '零售商名称'}, inplace=True)
    df2.drop(columns=['goods_num'], inplace=True)
    return df2


def main():
    first_date = '2022-2-1'
    last_date = '2022-4-12'
    excel_name = '顺丰已妥投数据导出2022-03-23(1)(1).xlsx'
    df = data_from_sql(first_date, last_date)
    df1 = data_from_excel(excel_name)
    df2 = data_transform(df=df, df1=df1)
    df2.to_excel("shunfen.xlsx", index=False)




if __name__ == '__main__':
    main()


