#!/usr/bin/env python
# coding: utf-8

# In[1]:


import os
import glob

# 获取当前目录路径
current_dir = os.getcwd()

# 获取当前目录下所有 .xlsx 文件的文件名列表
xlsx_files = glob.glob(os.path.join(current_dir, '*.xlsx'))

# 遍历文件名列表并删除文件
for file in xlsx_files:
    os.remove(file)
    print('已删除%s'%file)


# In[ ]:




