import sys
import pymysql
import pandas as pd
from datetime import datetime
from pandas import ExcelWriter

print('请输入开始时间，不输入默认为当前日期')
s = input()
if s == '':
    dt = datetime.today().strftime('%Y-%m-%d')
else:
    dt = s
fmd = datetime.today().strftime('%Y-%m-01')


def write_excel(writer, df, sheet_name):
    df.fillna('')
    df.to_excel(writer, sheet_name, index=False)
    workbook1 = writer.book
    worksheet1 = writer.sheets[sheet_name]
    # 设置写入excel表头格式
    title_format = workbook1.add_format({'font_name': u'微软雅黑', 'bg_color': '#95B3D7',
                                         'font_color': 'black', 'valign': 'vcenter',
                                         'align': 'center', 'font_size': 10, 'border': 0,
                                         'bold': 'True'})

    # 设置写入excel内容格式
    common_format = workbook1.add_format({'font_name': u'微软雅黑', 'font_size': 10,
                                          'valign': 'vcenter', 'align': 'center', 'border': 0})

    # 设置写入excel日期格式
    date_time_format = workbook1.add_format({'font_name': u'微软雅黑', 'font_size': 10,
                                             'valign': 'vcenter', 'align': 'center',
                                             'num_format': 'yyyy/mm/dd HH:mm', 'border': 0
                                             })
    # 写入标题
    for col_num, value in enumerate(list(df)):
        worksheet1.write(0, col_num, value, title_format)
    # 写入数据
    cols_total = list(df)
    df = df.reset_index()
    for index, value_serise in df.iterrows():
        col_num = -1
        try:
            for value in list(value_serise)[1:]:
                col_num += 1
                worksheet1.write(index + 1, col_num, value, common_format)
        except:
            pass
    # 设置列宽
    worksheet1.set_column(0, len(cols_total), 14)
    # 设置标题行高
    worksheet1.set_row(0, 30)
    # 设置日期格式
    date_format_list = ['create_time', '修改时间', '加入推送列表时间', '生效时间']
    for date_col in date_format_list:
        if date_col in cols_total:
            n = 0
            for value in df[date_col]:
                try:
                    n += 1
                    worksheet1.write(n, cols_total.index(date_col), value, date_time_format)
                except:
                    pass


def sync_solution(df_sync):
    if '经销商匹配失败' in str(df_sync['remark']):
        return '系统中未匹配到sold_to和ship_to'
    elif '已关店门店，信息不可修改' in str(df_sync['remark']):
        return '已关店门店，无需处理'
    elif '门店送达方编码为空' in str(df_sync['remark']):
        return 'ship_to_code为空，MDM进行维护'
    else:
        return ''


def store_to_erp_solution(df_store_to_erp):
    if '不能为空' in str(df_store_to_erp['回调报文']):
        return str(df_store_to_erp['回调报文']).split('":"')[1].split(' ')[0] + '不能为空，区域补全信息'
    elif '注册文书图片必传' in str(df_store_to_erp['回调报文']):
        return '营业执照照片不能为空，区域补全信息'
    elif '注册文书名称和营业执照图片中公司名称不一致，请检查' in str(df_store_to_erp['回调报文']):
        return '注册文书名称和营业执照图片中公司名称不一致，请检查。也可更换更清晰的图片再次尝试'
    else:
        return ''


def get_conn():
    conn = pymysql.connect(
        host='rm-uf6lxxf35108k7uq19o.mysql.rds.aliyuncs.com',
        user='danego_ro',
        password='D@none20181024',
        db='danego',
        charset='utf8'
    )
    return conn


def query(sql):
    conn = get_conn()
    df = pd.read_sql(sql, conn)
    return df


def get_sync():
    sql = '''
    SELECT
	t1.mdm_code,
CASE
		t1.STATUS 
		WHEN 1 THEN
		'同步成功' 
		WHEN 2 THEN
		'同步失败' 
		WHEN 3 THEN
		'无数据同步' ELSE "" 
	END '状态',
	t1.content,
	t1.remark,
	t1.create_time,
	t2.Distcode_Code,
	t2.Distcode_Name,
	t2.ShiptoCode,
	t2.ShiptoName,
	t2.Validity 
FROM
	store_info_sync t1
	LEFT JOIN user_retailer_missing_mdm_copy t2 ON t1.mdm_code = t2.`Code` 
WHERE
	date(t1.create_time) >= '{}'
    '''.format(dt)
    df_sync = query(sql)
    df_sync['解决方案'] = df_sync.apply(lambda x: sync_solution(x), axis=1)
    df_sync.sort_values(by='解决方案', ascending=False, inplace=True)
    return df_sync


def get_change_log():
    sql = '''
    SELECT t1.retailer_id '达易购ID',t2.name '零售商名称',t1.text '修改内容',t1.operator '操作人',t1.change_time '修改时间' FROM `user_retailer_change_log` t1
    left join user_retailer t2
    on t1.retailer_id = t2.id
    where t1.text like '%资金商%'
    and date(change_time) >= '{}'
    '''.format(dt)
    df_change_log = query(sql)
    return df_change_log


def get_store_to_erp():
    sql = '''
    SELECT
    case t1.source 
    when 0 THEN  '达易购'
    WHEN 1 THEN '海拍客'
    else '' end as '来源',
    t1.retailer_id '达易购ID',
    t3.mdm_code 'MDM码',
    t3.name '零售商名称',
    t4.name '资金商名称',
    case t1.type 
    when 1 then '新增'
    when 2 then '修改'
    else '' end as '类型',
    case t1.is_effect
    when 0 then '未生效'
    when 1 then '已生效'
    else '' end as '是否生效',
    case t1.status 
    when 0 then '待推送'
    when 1 then '推送成功'
    when 2 then '推送失败'
    when 3 then '终止推送'
    else '' end as '推送状态',
    t1.count '推送次数',
    t1.create_time '加入推送列表时间',
    t1.effect_time '生效时间',
    t1.success_time '推送成功时间',
    t2.push_content '推送报文',
    t2.callback '回调报文',
    t2.remark '备注'
    FROM
	( SELECT * FROM retailer_push_erp ORDER BY create_time DESC ) t1
	LEFT JOIN (select * from retailer_push_erp_log ORDER BY push_time DESC) t2 ON t1.id = t2.push_id 
	left join user_retailer t3 on t1.retailer_id=t3.id
	left join user_agent t4 on t1.agent_id=t4.id
    WHERE
	date(t1.create_time) >= '{}' 
	and t1.`status`=2
	GROUP BY t1.retailer_id
    '''.format(fmd)
    df_store_to_erp = query(sql)
    df_store_to_erp['解决方案'] = df_store_to_erp.apply(lambda x: store_to_erp_solution(x), axis=1)
    df_store_to_erp.sort_values(by='解决方案', ascending=False, inplace=True)
    return df_store_to_erp


def get_order_to_erp():
    sql = '''
    SELECT
CASE
		t1.source 
		WHEN 0 THEN
		'达易购' 
		WHEN 1 THEN
		'海拍客' ELSE '' 
	END AS '来源',
	t1.retailer_id '达易购ID',
	t3.mdm_code 'MDM码',
	t3.NAME '零售商名称',
	t5.name '省份',
	t1.order_code '订单号',
	t4.NAME '资金商名称',
CASE
		t1.is_effect 
		WHEN 0 THEN
		'未生效' 
		WHEN 1 THEN
		'已生效' ELSE '' 
	END AS '是否生效',
CASE
		t1.STATUS 
		WHEN 0 THEN
		'待推送' 
		WHEN 1 THEN
		'推送成功' 
		WHEN 2 THEN
		'推送失败' 
		WHEN 3 THEN
		'终止推送' ELSE '' 
	END AS '推送状态',
	t1.count '推送次数',
	t1.create_time '加入推送列表时间',
	t1.effect_time '生效时间',
	t1.success_time '推送成功时间',
	t2.push_content '推送报文',
	t2.callback '回调报文',
	t2.remark '备注' 
FROM
	( SELECT * FROM order_push_erp ORDER BY create_time DESC ) t1
	LEFT JOIN ( SELECT * FROM order_push_erp_log ORDER BY push_time DESC ) t2 ON t1.id = t2.push_id
	LEFT JOIN user_retailer t3 ON t1.retailer_id = t3.id
	LEFT JOIN user_agent t4 ON t1.agent_id = t4.id
left join region t5 on t3.province=t5.id	
where 
	t1.status=2 and date(t1.create_time) >= '{}'
GROUP BY
	t1.order_code
    '''.format(fmd)
    df_order_to_erp = query(sql)
    return df_order_to_erp


if __name__ == '__main__':
    print('正在新建文件')
    writer: ExcelWriter = pd.ExcelWriter(r'C:\Users\ASUS\Desktop\监测日志.xlsx', engine='xlsxwriter')
    print('正在写入第一个sheet')
    write_excel(writer, get_sync(), 'MDM同步至EGO')
    print('正在写入第二个sheet')
    write_excel(writer, get_change_log(), '供货关系变更')
    print('正在写入第三个sheet')
    write_excel(writer, get_store_to_erp(), '零售商推送至ERP')
    print('正在写入第四个sheet')
    write_excel(writer, get_order_to_erp(), '订单推送至ERP')
    print('写入完成，正在保存')
    writer.save()
