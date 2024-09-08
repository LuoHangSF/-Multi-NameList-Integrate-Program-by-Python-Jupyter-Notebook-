#!/usr/bin/env python
# coding: utf-8

# # 全部打卡整合脚本
# 

# ## 导入需要的Python库

# In[1]:


# 必须导入，不导入没办法运行
import sys
from tkinter.filedialog import askopenfilename
from tkinter import Tk, messagebox
import pandas as pd
import numpy as np
import os


# ## 载入所有名单(在此处修改项目数量以及名字)

# In[2]:


# # 可以修改，有多少项目就添加多少项目，项目名称需要与excel文件 除去后缀的 名称一致
# # 项目名称要加英文单引号，不同项目之间要用逗号隔开
# # event_list = ['练字', '单词', '跑步', '光盘', '记账'] 

# # 打开文件并读取所有行
# with open('events.txt', 'r', encoding='utf-8') as file:
#     event_list = [line.strip() for line in file]

Tk().withdraw()

# 打开文件对话框让用户选择文件
file_path = askopenfilename(title="选择项目列表的txt文件", filetypes=[("Text files", "*.txt")])

# 检查用户是否选择了文件
if file_path:
    # 打开文件并读取所有行
    with open(file_path, 'r', encoding='utf-8') as file:
        event_list = [line.strip() for line in file]
    messagebox.showinfo("读取成功", f"读取成功！项目共计{len(event_list)}个，分别为：\n{', '.join(event_list)}")
    # 打印读取到的项目列表
    print(event_list)
else:
    messagebox.showerror("错误", "未选择文件")
    sys.exit()
    raise Exception("未选择文件")
    print("未选择文件")

# 打印读取到的项目列表
print(event_list)


# excel文件的格式（默认为xlsx，如果是csv的话就改成.csv）
suffix = ".xlsx"


# ## 读取相应打卡数据

# In[3]:


# 初始化一个字典，用于根据上面列表提供的项目读取对应的项目打卡分数数据
data_dict = {}
# 分别导入每个项目的excel数据

for item in event_list:
    item_name = item + suffix
    try:
        data_dict[item] = pd.read_excel(item_name)
    except:
        messagebox.showerror("错误", f"未找到文件：{item_name} \n程序将终止运行，请检查文件是否存在")
        sys.exit()




# ## 创建整合全部打卡项目的名单的函数

# In[4]:


def sum_name(dict):
    result_name = []
    num = 1
    for key in dict: # 遍历字典 data_dict 中的每个键
        events = dict[key] # 获取当前键对应的值（即数据帧），并将其赋值给 df
        if num == 1:
            for name in events["圈子昵称"]:
                result_name.append(name)
                
        else:
            for name in events["圈子昵称"]: # 遍历名单
                if name not in result_name:  # 判断：如果最终名单中没有这个人，就将这个人导入最终名单
                    result_name.append(name)
                
        num += 1
        
    return result_name


# ## 运行该函数，获得打卡名单总表

# In[5]:


a = sum_name(data_dict) # 调用函数
len(a)


# ## 创建最终结果的DataFrame

# In[6]:


final_df = pd.DataFrame(a) # 用最终名单列表的创建一个新的DataFrame
final_df


# ## 下面分别更改不同档次名单DataFrame的行index

# In[8]:


data_dict_reindexed = {}


# In[9]:


def reindex(dict):
    rein = "_reindexed"
    for key in dict:
        event_reindexed = key + rein
        tmp_df = dict[key]
        tmp_df = tmp_df.set_index(pd.Index(tmp_df["圈子昵称"]))
        data_dict_reindexed[event_reindexed] = tmp_df


# In[10]:


reindex(data_dict)
# data_dict_reindexed["跑步_reindexed"]


# In[11]:


final_df_reindexed = final_df.set_index(pd.Index(final_df[0]))

# final_df_reindexed


# ## 创建五个函数，分别计算每个人的打卡天数以及积分

# ### 给参与者统计每一个项目获得的分数的函数

# In[12]:


def 单独项目赋分(项目, df):
    积分 = [] # 创建一个新的积分空列表
    # scores = []
    
    for df1 in df[0]: # 按顺序遍历每一个名字
        a = 0
        # b = 0
        for 人员 in 项目["圈子昵称"]: # 遍历圈子中的每一个名字
            if 人员 == df1: # 最终名单得里的名字在圈子的名单中，就将圈子中的积分赋给a
                a = 项目.loc[人员, "积分"]
                
        积分.append(a) # 按遍历的顺序将a加到积分的列表中
        
    return 积分 # 返回列表


# ## 将积分以列表形式导出

# In[13]:


所有项目积分 = {}


# In[14]:


def 全部项目整合(dict, dict_reindexed):
    score = "积分"
    rein = "_reindexed"
    for key in dict:
        key_n = key + score
        key_rein = key + rein
        所有项目积分[key_n] = 单独项目赋分(dict_reindexed[key_rein], final_df_reindexed)


# In[15]:


全部项目整合(data_dict, data_dict_reindexed)


# In[16]:





# In[18]:


# len(所有项目积分['记账积分'])


# In[19]:





# In[20]:





# In[21]:





# ## 在最终结果中的DataFrame中创建五个新的列，然后将各项目的积分加到列中

# In[22]:


final_dataframe = final_df_reindexed


# In[23]:


def dataframe_add_events_columns(data_score):
    for key in data_score:
        final_dataframe[key] = data_score[key]


# In[24]:


dataframe_add_events_columns(所有项目积分)


# In[26]:


final_dataframe


# ## 计算总积分

# ### 计算总积分的函数

# In[27]:


events_score_key = []
for key in 所有项目积分:
    events_score_key.append(key)
events_score_key


# In[28]:


def final_score_sum(df, data_key):
    final_score = [] # 创建总积分列表
    
    for name in df[0]: # 遍历每一个名字
        a = 0 
        for events in data_key:
            a = a + df.loc[name, events]
        # a = df.loc[df1, "单词积分"] + df.loc[df1, "跑步积分"] + df.loc[df1, "记账积分"] + df.loc[df1, "练字积分"] + df.loc[df1, "光盘积分"]
        # 将每一个人的各项目积分按行求和
        final_score.append(a) # 按顺序加入到总积分列表后面
        
    return final_score # 返回总积分


# ### 计算总积分

# In[29]:


final_score = final_score_sum(final_dataframe, events_score_key) # 调用函数


# In[30]:


len(final_score)


# ### 将总积分加入到DataFrame中

# In[31]:


final_dataframe["总积分"] = final_score # 创建新列，载入总积分
final_dataframe


# ## 将最终结果导出到excel

# In[32]:


final_dataframe.to_excel('总积分表.xlsx', index=False) # 导出，文件名为“总积分表.xlsx”
messagebox.showinfo("处理完成", "项目积分整合完成！")

# In[ ]:




