#!/usr/bin/env python
# coding: utf-8

# # 单词打卡整合脚本
# 

# In[1]:

import sys
from tkinter.filedialog import askopenfilename
from tkinter import Tk, messagebox
import pandas as pd
import numpy as np


# ## 读取相应打卡数据

# In[2]:
Tk().withdraw()

# 打开文件对话框让用户选择文件
file_path30 = askopenfilename(title="选择每天背诵30-70个单词的打卡数据", filetypes=[("Excel files", "*.xlsx")])
if file_path30:
    series30 = pd.read_excel(file_path30)
    print(series30)
    messagebox.showinfo("读取成功", f"已录入文件:{file_path30}作为每天背诵30-70个单词的打卡数据")
else:
    messagebox.showerror("读取失败", "未选择文件")
    sys.exit()
    
file_path70 = askopenfilename(title="选择每天背诵71个单词以上的打卡数据", filetypes=[("Excel files", "*.xlsx")])
if file_path70:
    series70 = pd.read_excel(file_path70)
    print(series70)
    messagebox.showinfo("读取成功", f"已录入文件:{file_path70}作为每天背诵71个单词以上的打卡数据")
else:
    messagebox.showerror("读取失败", "未选择文件")
    sys.exit()


#dataframeof30 = pd.DataFrame()


# In[3]:


# series70


# In[4]:


result_name = []


# ## 整合两个打卡项目的名单

# In[5]:


def intergartion_name(data30, data70):
    result_name = []
    
    for da30 in data30["圈子昵称"]:
        result_name.append(da30)
        
    for da70 in data70["圈子昵称"]:
        if da70 not in result_name:
            result_name.append(da70)
        
    return result_name


# In[6]:


a = intergartion_name(series30, series70)
len(a)


# ## 创建最终结果的DataFrame

# In[7]:


df = pd.DataFrame(a)
df


# In[8]:


# df.to_excel('output.xlsx', index=False) 


# ## 下面分别更改不同档次名单DataFrame的行index

# In[9]:


df30_reindexed = series30.set_index(pd.Index(series30["圈子昵称"]))
df30_reindexed


# In[10]:


df70_reindexed = series70.set_index(pd.Index(series70["圈子昵称"]))
df70_reindexed


# In[11]:


df_reindexed = df.set_index(pd.Index(df[0]))
df_reindexed


# ## 创建两个函数，分别计算每个人的打卡天数以及积分

# ### 计算打卡天数函数

# In[12]:


def integrate_days(df30, df70, df):
    days = []
    scores = []
    
    for df in df[0]:
        a = 0
        b = 0
        for data30 in df30["圈子昵称"]:
            if data30 == df:
                a = df30.loc[data30, "打卡天数"]
                b = df30.loc[data30, "打卡天数"]
                
        for data70 in df70["圈子昵称"]:
            if data70 == df:
                a = a + df70.loc[data70, "打卡天数"]
                b = b + 2 * df70.loc[data70, "打卡天数"]
        days.append(a)
        scores.append(b)
            
            
#        for data70 in df70["圈子昵称"]:
#            if data70 == df:
#                a = a + df70.loc[data70, "打卡天数"]
                
#        days.append(a)
    
    return days


# ### 计算积分的函数

# In[13]:


def integrate_scores(df30, df70, df):
    days = []
    scores = []
    
    for df in df[0]:
        a = 0
        b = 0
        for data30 in df30["圈子昵称"]:
            if data30 == df:
                a = df30.loc[data30, "打卡天数"]
                b = df30.loc[data30, "打卡天数"]
                
        for data70 in df70["圈子昵称"]:
            if data70 == df:
                a = a + df70.loc[data70, "打卡天数"]
                b = b + 2 * df70.loc[data70, "打卡天数"]
        days.append(a)
        scores.append(b)
            
            
#        for data70 in df70["圈子昵称"]:
#            if data70 == df:
#                a = a + df70.loc[data70, "打卡天数"]
                
#        days.append(a)
    
    return scores


# ## 将积分以列表形式导出

# In[14]:


days = integrate_days(df30_reindexed, df70_reindexed, df_reindexed)
scores = integrate_scores(df30_reindexed, df70_reindexed, df_reindexed)


# In[15]:


len(days)


# In[16]:


len(scores)


# ## 在最终结果中的DataFrame中创建两个新的行，然后将打卡天数和积分导入进去

# In[17]:


df["打卡天数"] = days
df["积分"] = scores


# In[18]:


# df


# ## 将最终结果导出到excel

# In[19]:


df.to_excel('单词打卡积分表01.xlsx', index=False) 
messagebox.showinfo("处理成功", "处理成功：已导出文件单词打卡积分表01.xlsx")

# In[ ]:




