# -*- coding: utf-8 -*-
"""
Created on Fri May 19 16:18:22 2023

@author: fuwenyue
"""


import pandas as pd
import numpy as np
import time,os,glob,configparser



conf = configparser.ConfigParser()
conf.read('config.ini', encoding="utf-8-sig")
workfloder = conf.get('Path','workfloder')
print(f'检测到当前配置工作目录为:{workfloder}')
fileList = glob.glob(os.path.join(workfloder,'*.xlsx'))
fliePath = [file for file in fileList if os.path.basename(file).startswith('汇总表')][0]
print(f'检测到汇总表位于：{fliePath}')
print('*'*10)
# Calculate percentage
def calculate_percentage(a, b):
    return (a - b) / b
now = time.strftime('%Y-%m-%d')
#输出文件名称
output_file_name = workfloder+f'\\同环比分表_{now}.xlsx'
# Remove existing file if it exists
if os.path.exists(output_file_name):
    os.remove(output_file_name)
#获取对比期表头
def duibiqi_values(duibiqi):
    pivot_values = ['本期是否营业', f'{duibiqi}期是否营业', '本期实收金额', f'{duibiqi}期实收金额', '本期堂食实收', f'{duibiqi}期堂食实收',
              '本期外卖实收', f'{duibiqi}期外卖实收', '本期自提实收', f'{duibiqi}期自提实收']
    output_values = ['本期是否营业',f'{duibiqi}期是否营业','店铺增减','本期实收金额',f'{duibiqi}期实收金额',f'实收{duibiqi}','本期堂食实收',f'{duibiqi}期堂食实收',
              f'堂食实收{duibiqi}','本期外卖实收',f'{duibiqi}期外卖实收',f'外卖实收{duibiqi}','本期自提实收',f'{duibiqi}期自提实收',f'自提实收{duibiqi}']
    return pivot_values,output_values
#移除检测(暂时弃用)
def remove_check(person,output_df):
    if person == '省经理':
        output_df.drop('赵磊', inplace=True)
    return output_df
#生成汇总行
def total_row(df,duibiqi):
    total = df[['本期店铺数', f'{duibiqi}期店铺数','店铺增减','本期实收金额',f'{duibiqi}期实收金额',
                '本期堂食实收',f'{duibiqi}期堂食实收','本期外卖实收',f'{duibiqi}期外卖实收','本期自提实收',f'{duibiqi}期自提实收']].sum()
    total[f'实收{duibiqi}'] = calculate_percentage(total['本期实收金额'] , total[f'{duibiqi}期实收金额'])
    total[f'堂食实收{duibiqi}'] = calculate_percentage(total['本期堂食实收'] , total[f'{duibiqi}期堂食实收'])
    total[f'外卖实收{duibiqi}'] = calculate_percentage(total['本期外卖实收'] , total[f'{duibiqi}期外卖实收'])
    total[f'自提实收{duibiqi}'] = calculate_percentage(total['本期自提实收'] , total[f'{duibiqi}期自提实收'])
    total_df = pd.DataFrame(total).T.rename(index={0: '合计'})
    return total_df

def calculate_result(current_table,person,duibiqi):#需要透视的表、经理、对比期
    pivot_values,output_values = duibiqi_values(duibiqi)
    pivot_df = current_table.pivot_table(index=[f'{person}'],values = pivot_values,aggfunc=np.sum)
    pivot_df['店铺增减'] = pivot_df['本期是否营业']- pivot_df[f'{duibiqi}期是否营业']
    pivot_df[f'实收{duibiqi}'] = calculate_percentage(pivot_df['本期实收金额'] ,pivot_df[f'{duibiqi}期实收金额'])
    pivot_df.sort_values(f"实收{duibiqi}",inplace=True,ascending=False) #排序
    pivot_df[f'堂食实收{duibiqi}'] = calculate_percentage(pivot_df['本期堂食实收'] ,pivot_df[f'{duibiqi}期堂食实收'])
    pivot_df[f'外卖实收{duibiqi}'] = calculate_percentage(pivot_df['本期外卖实收'] ,pivot_df[f'{duibiqi}期外卖实收'])
    pivot_df[f'自提实收{duibiqi}'] = calculate_percentage(pivot_df['本期自提实收'] ,pivot_df[f'{duibiqi}期自提实收'])
    output_df = pivot_df.reindex(columns=output_values)
    output_df.rename(columns={'本期是否营业': '本期店铺数',f'{duibiqi}期是否营业': f'{duibiqi}期店铺数'}, inplace=True)
    total_df = total_row(output_df,duibiqi)
    output_df = pd.concat([output_df, total_df])
    #output_df = remove_check(person,output_df)
    if os.path.exists(output_file_name):
        with pd.ExcelWriter(output_file_name, mode='a',engine="openpyxl") as writer:
            output_df.to_excel(writer, sheet_name=f'{person}{liangji}店铺{duibiqi}')
    else:
        output_df.to_excel(output_file_name, sheet_name=f'{person}{liangji}店铺{duibiqi}')

df_zongbiao = pd.read_excel(fliePath,sheet_name="总表",header = 3)
#全量店铺
liangji = '全量'
df_zongbiao[["大区经理", "省经理","区域经理"]] = df_zongbiao[["大区经理", "省经理","区域经理"]].replace(np.nan, "错误，请提醒我重做")
current_table = df_zongbiao #需要透视的表-总表
for person in ['大区经理','省经理','区域经理']:
    for duibiqi in ['同比','环比']:
        print(f'正在生成{person}{liangji}店铺{duibiqi}表格')
        calculate_result(current_table,person,duibiqi)
        
#同比存量店铺
liangji = '存量'
df_tongbi_cunliang = df_zongbiao[(df_zongbiao['本期是否营业'] == 1) & (df_zongbiao['同比期是否营业'] == 1)]
current_table = df_tongbi_cunliang
for person in ['大区经理','省经理','区域经理']:
    for duibiqi in ['同比']:
        print(f'正在生成{person}{liangji}店铺{duibiqi}表格')
        calculate_result(current_table,person,duibiqi)
#环比存量店铺
liangji = '存量'
df_huanbi_cunliang = df_zongbiao[(df_zongbiao['本期是否营业'] == 1) & (df_zongbiao['环比期是否营业'] == 1)]
current_table = df_huanbi_cunliang
for person in ['大区经理','省经理','区域经理']:
    for duibiqi in ['环比']:
        calculate_result(current_table,person,duibiqi)
        print(f'已生成{person}{liangji}店铺{duibiqi}表……')

#同环比汇总表制作
def XifenPart(current_table,duibiqi):#表格，对比期
    pivot_values = ['本期是否营业',f'{duibiqi}期是否营业','本期实收金额',f'{duibiqi}期实收金额']
    pivot_df = current_table.pivot_table(index=['大区经理','省经理','区域经理'],values = pivot_values,aggfunc=np.sum)
    pivot_df['店铺增减'] = pivot_df['本期是否营业']-pivot_df[f'{duibiqi}期是否营业']
    pivot_df[f'实收{duibiqi}'] = calculate_percentage(pivot_df['本期实收金额'] , pivot_df[f'{duibiqi}期实收金额'])
    pivot_df.rename(columns={'本期是否营业': '本期店铺数',f'{duibiqi}期是否营业': f'{duibiqi}期店铺数'}, inplace=True)
    columns = ['本期店铺数',f'{duibiqi}期店铺数','店铺增减','本期实收金额',f'{duibiqi}期实收金额',f'实收{duibiqi}']
    pivot_df = pivot_df.reindex(columns=columns)
    return pivot_df

quanliang_tongbi = XifenPart(df_zongbiao,'同比') #全量同比
quanliang_huanbi = XifenPart(df_zongbiao,'环比') #全量环比
cunliang_tongbi = XifenPart(df_tongbi_cunliang,'同比') #存量同比
cunliang_huanbi = XifenPart(df_huanbi_cunliang,'环比') #存量环比

tongbi_weidu = pd.merge(quanliang_tongbi,cunliang_tongbi,on=['大区经理','省经理','区域经理'], how='outer')
huanbi_weidu = pd.merge(quanliang_huanbi,cunliang_huanbi,on=['大区经理','省经理','区域经理'], how='outer')
zong_weidu = pd.merge(tongbi_weidu,huanbi_weidu,on=['大区经理','省经理','区域经理'])
zong_weidu.drop(columns=['同比期店铺数_y','店铺增减_y_x','环比期店铺数_y','店铺增减_y_y'],inplace=True) 
zong_weidu.rename(columns={'本期是否营业': '本期店铺数','同比期是否营业': '同比期店铺数'}, inplace=True)
with pd.ExcelWriter(output_file_name, mode='a',engine="openpyxl") as writer:
    zong_weidu.to_excel(writer, sheet_name='总表',merge_cells=False)
    print('已经生成同环比表总表……')
#区域经理省经理分离
df_sheng = pd.read_excel(output_file_name,sheet_name="省经理全量店铺同比")
shengjingli_list = list(df_sheng['Unnamed: 0'])
shengjingli_list[:] = [item for item in shengjingli_list if item not in ['本期未营业', '已解约','合计']]


with pd.ExcelWriter(output_file_name, mode='a', engine="openpyxl") as writer:
    for liangji in ['全量', '存量']:
        for duibiqi in ['同比', '环比']:
            df_quyu = pd.read_excel(output_file_name, sheet_name=f"区域经理{liangji}店铺{duibiqi}", index_col=0)
            df_quyu = df_quyu.iloc[:-1]  # 删除合计行

            df_quyu_sheng = df_quyu[df_quyu.index.isin(shengjingli_list)]
            total_df_sheng = total_row(df_quyu_sheng, f'{duibiqi}')
            output_df_sheng = pd.concat([df_quyu_sheng, total_df_sheng])
            output_df_sheng.to_excel(writer, sheet_name=f'区域经理{liangji}店铺{duibiqi}(省代)', index=True)

            df_quyu_quyu = df_quyu[~df_quyu.index.isin(shengjingli_list)]
            total_df_quyu = total_row(df_quyu_quyu, f'{duibiqi}')
            output_df_quyu = pd.concat([df_quyu_quyu, total_df_quyu])
            output_df_quyu.to_excel(writer, sheet_name=f'区域经理{liangji}店铺{duibiqi}(区域)', index=True)
print(f'生成的文件位于:{output_file_name}')