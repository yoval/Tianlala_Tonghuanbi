# -*- coding: utf-8 -*-
"""
Created on Fri May 19 15:32:33 2023

@author: fuwenyue
"""

import os,glob,configparser,openpyxl,time

conf = configparser.ConfigParser()
conf.read('config.ini', encoding="utf-8-sig")
workfloder = conf.get('Path','workfloder')
mobanPath = conf.get('Path','fenbiaomoban')
fileList = glob.glob(os.path.join(workfloder,'*.xlsx'))
fliePath = [file for file in fileList if os.path.basename(file).startswith('同环比分表')][-1] #最新同环比分表
print(f'检测到当前配置工作目录为:{workfloder}')
print(f'检测到当前模板配置路径为:{mobanPath}')
print(f'检测到同环比分表位于：{fliePath}')
print('*'*10)
now = time.strftime('%Y-%m-%d')
output_file_name = workfloder+f'\\大区省区域经理分表输出_{now}.xlsx'
source_file = openpyxl.load_workbook(fliePath)
target_file = openpyxl.load_workbook(mobanPath)

#大区经理
target_sheet = target_file["大区经理排名"]
source_sheet = source_file["大区经理全量店铺同比"]
for r in range(1, source_sheet.max_row + 1):
    for c in range(1, source_sheet.max_column + 1):
        target_sheet.cell(row=r+2, column=c+2).value = source_sheet.cell(row=r, column=c).value

source_sheet = source_file["大区经理全量店铺环比"]
for r in range(1, source_sheet.max_row + 1):
    for c in range(1, source_sheet.max_column + 1):
        target_sheet.cell(row=r+2, column=c+21).value = source_sheet.cell(row=r, column=c).value

source_sheet = source_file["大区经理存量店铺同比"]
for r in range(1, source_sheet.max_row + 1):
    for c in range(1, source_sheet.max_column + 1):
        target_sheet.cell(row=r+34, column=c+2).value = source_sheet.cell(row=r, column=c).value

source_sheet = source_file["大区经理存量店铺环比"]
for r in range(1, source_sheet.max_row + 1):
    for c in range(1, source_sheet.max_column + 1):
        target_sheet.cell(row=r+34, column=c+21).value = source_sheet.cell(row=r, column=c).value

target_file.save("大区省区域经理分表输出.xlsx")

#省经理
target_sheet = target_file["省经理排名"]

source_sheet = source_file["省经理全量店铺同比"]

for r in range(1, source_sheet.max_row + 1):
    for c in range(1, source_sheet.max_column + 1):
        target_sheet.cell(row=r+2, column=c+2).value = source_sheet.cell(row=r, column=c).value

source_sheet = source_file["省经理全量店铺环比"]

for r in range(1, source_sheet.max_row + 1):
    for c in range(1, source_sheet.max_column + 1):
        target_sheet.cell(row=r+2, column=c+21).value = source_sheet.cell(row=r, column=c).value

source_sheet = source_file["省经理存量店铺同比"]

for r in range(1, source_sheet.max_row + 1):
    for c in range(1, source_sheet.max_column + 1):
        target_sheet.cell(row=r+58, column=c+2).value = source_sheet.cell(row=r, column=c).value

source_sheet = source_file["省经理存量店铺环比"]

for r in range(1, source_sheet.max_row + 1):
    for c in range(1, source_sheet.max_column + 1):
        target_sheet.cell(row=r+58, column=c+21).value = source_sheet.cell(row=r, column=c).value

#target_file.save("大区省区域经理分表输出.xlsx")
#区域经理(区域)
target_sheet = target_file["区域经理排名"]

source_sheet = source_file["区域经理全量店铺同比(区域)"]

for r in range(1, source_sheet.max_row + 1):
    for c in range(1, source_sheet.max_column + 1):
        target_sheet.cell(row=r+2, column=c+2).value = source_sheet.cell(row=r, column=c).value

source_sheet = source_file["区域经理全量店铺环比(区域)"]

for r in range(1, source_sheet.max_row + 1):
    for c in range(1, source_sheet.max_column + 1):
        target_sheet.cell(row=r+2, column=c+21).value = source_sheet.cell(row=r, column=c).value

source_sheet = source_file["区域经理存量店铺同比(区域)"]

for r in range(1, source_sheet.max_row + 1):
    for c in range(1, source_sheet.max_column + 1):
        target_sheet.cell(row=r+143, column=c+2).value = source_sheet.cell(row=r, column=c).value

source_sheet = source_file["区域经理存量店铺环比(区域)"]

for r in range(1, source_sheet.max_row + 1):
    for c in range(1, source_sheet.max_column + 1):
        target_sheet.cell(row=r+143, column=c+21).value = source_sheet.cell(row=r, column=c).value

#区域经理(省代)        
source_sheet = source_file["区域经理全量店铺同比(省代)"]

for r in range(1, source_sheet.max_row + 1):
    for c in range(1, source_sheet.max_column + 1):
        target_sheet.cell(row=r+283, column=c+2).value = source_sheet.cell(row=r, column=c).value

source_sheet = source_file["区域经理全量店铺环比(省代)"]

for r in range(1, source_sheet.max_row + 1):
    for c in range(1, source_sheet.max_column + 1):
        target_sheet.cell(row=r+283, column=c+21).value = source_sheet.cell(row=r, column=c).value

source_sheet = source_file["区域经理存量店铺同比(省代)"]

for r in range(1, source_sheet.max_row + 1):
    for c in range(1, source_sheet.max_column + 1):
        target_sheet.cell(row=r+334, column=c+2).value = source_sheet.cell(row=r, column=c).value

source_sheet = source_file["区域经理存量店铺环比(省代)"]

for r in range(1, source_sheet.max_row + 1):
    for c in range(1, source_sheet.max_column + 21):
        target_sheet.cell(row=r+334, column=c+21).value = source_sheet.cell(row=r, column=c).value        
        
        
        
target_file.save(output_file_name)

print(f'生成的文件位于:{output_file_name}')




