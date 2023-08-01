#!/usr/bin/env python
# coding: utf-8

# In[64]:


get_ipython().system('pip install pandas')
get_ipython().system('pip install openpyxl')
get_ipython().system('pip install matplotlib')
get_ipython().system('pip install seaborn')
get_ipython().system('pip install scipy')
get_ipython().system('pip install statsmodels')
get_ipython().system('pip install xlsxwriter')



# In[2]:


#delete sc2_number_3.vsi files
import pandas as pd

# 读取原始Excel文件
df = pd.read_excel('/Users/sc17237/Desktop/new region.xlsx')

# 使用正则表达式筛选出 'Image' 列中以 '_3.vsi - 10x_DAPI' 结尾的行
df = df[~df['Image'].str.contains('sc2_\d+_3\.vsi - 10x_DAPI$', regex=True)]

# 创建新的Excel文件并将数据写入
df.to_excel('/Users/sc17237/Desktop/除掉3.xlsx', index=False)

print('操作完成，已经生成新的Excel文件。')


# In[12]:


#add treatment group to each animal

import pandas as pd
import re

# Load the Excel file
df = pd.read_excel('/Users/sc17237/Desktop/除掉3.xlsx')

# Define the groups
group_dict = {
    'IP OF': ['27', '30', '31', '32', '5', '6', '17', '18'],
    'IP HC': ['10', '12', '26', '28', '7', '8', '21', '22'],
    'Oral OF': ['14', '16', '29', '19', '20', '30', '23', '24', '25'],
    'Oral HC': ['9', '11', '13', '15', '1', '2', '3', '4']
}

# Function to retrieve the animal number from the filename
def get_animal_number(filename):
    match = re.search(r'sc2_(\d+)_', filename)
    if match:
        return match.group(1)
    return None

# Add a new column to the DataFrame for the animal number
df['Animal Number'] = df['Image'].apply(get_animal_number)

# Add a new column to the DataFrame for the treatment group
def get_treatment_group(animal_number):
    for group, numbers in group_dict.items():
        if animal_number in numbers:
            return group
    return None

df['Treatment Group'] = df['Animal Number'].apply(get_treatment_group)

# Save the DataFrame to a new Excel file
df.to_excel('/Users/sc17237/Desktop/treatment_group_added.xlsx', index=False)


# In[15]:


#sorting the file by brain area
with pd.ExcelWriter('/Users/sc17237/Desktop/treatment_group_added.xlsx') as writer:
    for area_name, data in df.groupby('Area Name '):
        data.to_excel(writer, sheet_name=area_name, index=False)


# In[79]:


#sorting the file by their treatment group
import pandas as pd

# 读取原始Excel文件
df = pd.read_excel('/Users/sc17237/Desktop/除掉3.xlsx')

# 按照 "Treatment Group" 进行分类
groups = df.groupby('Treatment Group')

# 创建一个新的Excel writer 对象
with pd.ExcelWriter('/Users/sc17237/Desktop/sorted_by_Treatmentgroup.xlsx') as writer:
    for name, group in groups:
        if name in ['Oral HC', 'Oral OF', 'IP HC', 'IP OF']:
            # 写入数据到对应的sheet
            group.to_excel(writer, sheet_name=name, index=False)

print('操作完成，已经生成新的Excel文件.')


# In[1]:


#info about each sheet

import pandas as pd

# 路径
file_path = "/Users/sc17237/Desktop/QuPath_ABBA_output_trap_sorted_by_area.xlsx"

# 使用pandas的ExcelFile函数打开Excel文件
excel_file = pd.ExcelFile(file_path)

# 对Excel文件中的每个sheet进行操作
for sheet_name in excel_file.sheet_names:
    # 读取当前sheet的数据
    df = excel_file.parse(sheet_name)
    
    # 查看有哪些Animal Number
    animal_numbers = df['Animal Number'].unique()
    print(f"In sheet {sheet_name}, we have these animal numbers: {animal_numbers}")
    
    # 计算每个Animal Number出现的次数
    animal_count = df['Animal Number'].value_counts()
    print(f"The count of each animal number in sheet {sheet_name} is:\n{animal_count}\n")


# In[14]:


print(df.columns)


# In[16]:


#read excel file and plot summary stats

import pandas as pd
import matplotlib.pyplot as plt

# Excel文件路径
file_path = "/Users/sc17237/Desktop/by_area.xlsx"

# 读取所有的sheets
all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')

# 按照sheet进行遍历
for sheet_name, df in all_sheets.items():
    print(f"Processing {sheet_name}...")

    # 基本信息
    print(df.describe())

    # 选择一列进行绘图，例如'TRAP Cell Count'
    df['TRAP Cell Count'].hist(bins=20)
    plt.title(f'Histogram of TRAP Cell Count in {sheet_name}')
    plt.xlabel('TRAP Cell Count')
    plt.ylabel('Frequency')
    plt.show()

    # 按Treatment Group 绘制平均'TRAP Cell Count'
    group_mean = df.groupby('Treatment Group')['TRAP Cell Count'].mean()
    group_mean.plot(kind='bar')
    plt.title(f'Average TRAP Cell Count by Treatment Group in {sheet_name}')
    plt.xlabel('Treatment Group')
    plt.ylabel('Average TRAP Cell Count')
    plt.show()


# In[71]:


import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

# Excel文件路径
file_path = "/Users/sc17237/Desktop/除掉了3.xlsx"

# 输出文件路径
output_path = "/Users/sc17237/Desktop/TRAP_DAPI_Percentage_Average.xlsx"

# 创建一个ExcelWriter对象，用于输出到Excel
with pd.ExcelWriter(output_path) as writer:
    # 读取所有的sheets
    all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')

    # 按照sheet进行遍历
    for sheet_name, df in all_sheets.items():
        print(f"Processing {sheet_name}...")
        
        # 计算每一只小鼠的TRAP/DAPI percentage
        df['TRAP/DAPI'] = df['TRAP Cell Count'] / df['DAPI Cell Count'] * 100

        # 计算每个Animal Number的TRAP/DAPI平均值
        df_mean = df.groupby(['Animal Number', 'Treatment Group'])['TRAP/DAPI'].mean().reset_index()

        # 将结果写入Excel
        df_mean.to_excel(writer, sheet_name=sheet_name, index=False)

        # 打印TRAP/DAPI百分比表格
        print(f"\nAverage TRAP/DAPI percentages in {sheet_name}:\n")
        print(df_mean)
        print("\n"+"="*40+"\n")

        # 设置Treatment Group的顺序
        df_mean['Treatment Group'] = pd.Categorical(df_mean['Treatment Group'], ["Oral HC", "Oral OF", "IP HC", "IP OF"])

        # 为每个Treatment Group设定颜色
        colors = {"Oral HC": "blue", "Oral OF": "orange", "IP HC": "green", "IP OF": "red"}

        # 绘制散点图
        plt.figure(figsize=(10, 6))
        sns.scatterplot(data=df_mean, x='Treatment Group', y='TRAP/DAPI', hue='Treatment Group', palette=colors, s=100)
        
        # 计算并绘制每个Treatment Group的平均值和标准偏差
        treatment_group_stats = df_mean.groupby('Treatment Group')['TRAP/DAPI'].agg(['mean', 'std', 'count'])
        for i, treatment_group in enumerate(treatment_group_stats.index):
            mean = treatment_group_stats.loc[treatment_group, 'mean']
            std = treatment_group_stats.loc[treatment_group, 'std']
            plt.plot([i-0.2, i+0.2], [mean, mean], color='black', lw=2)  # mean line
            plt.fill_between([i-0.2, i+0.2], mean-std, mean+std, color='gray', alpha=0.2)  # std deviation area

        # 打印每个Treatment Group的样本数
        print(f"Sample counts in {sheet_name}:\n")
        print(treatment_group_stats['count'])
        print("\n"+"="*40+"\n")

        plt.title(f'Average TRAP/DAPI by Treatment Group in {sheet_name} (n={df_mean["Treatment Group"].value_counts().to_dict()})')
        plt.ylabel('Average TRAPed/DAPI (%)')
        plt.ylim(0, 6)  # 设置y轴的范围
        plt.grid(True)
        plt.legend(title='Treatment Group')
        plt.show()


# In[8]:


#graphs draft1

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from scipy.stats import f_oneway
from statsmodels.stats.multicomp import pairwise_tukeyhsd, MultiComparison

# Excel文件路径
file_path = "/Users/sc17237/Desktop/QuPath_ABBA_output_trap_sorted_by_area.xlsx"

# 输出文件路径
output_path = "/Users/sc17237/Desktop/TRAP_DAPI_Percentage_Average.xlsx"

# 创建一个ExcelWriter对象，用于输出到Excel
with pd.ExcelWriter(output_path) as writer:
    # 读取所有的sheets
    all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')

    # 按照sheet进行遍历
    for sheet_name, df in all_sheets.items():
        print(f"Processing {sheet_name}...")
        
        # 计算每一只小鼠的TRAP/DAPI percentage
        df['TRAP/DAPI'] = df['TRAP Cell Count'] / df['DAPI Cell Count'] * 100

        # 计算每个Animal Number的TRAP/DAPI平均值
        df_mean = df.groupby(['Animal Number', 'Treatment Group'])['TRAP/DAPI'].mean().reset_index()

        # 将结果写入Excel
        df_mean.to_excel(writer, sheet_name=sheet_name, index=False)

        # 打印TRAP/DAPI百分比表格
        print(f"\nAverage TRAP/DAPI percentages in {sheet_name}:\n")
        print(df_mean)
        print("\n"+"="*40+"\n")

        # 设置Treatment Group的顺序
        df_mean['Treatment Group'] = pd.Categorical(df_mean['Treatment Group'], ["Oral HC", "Oral OF", "IP HC", "IP OF"])

        # 为每个Treatment Group设定颜色
        colors = {"Oral HC": "blue", "Oral OF": "orange", "IP HC": "green", "IP OF": "red"}

        # 绘制散点图
        plt.figure(figsize=(10, 6))
        sns.scatterplot(data=df_mean, x='Treatment Group', y='TRAP/DAPI', hue='Treatment Group', palette=colors, s=100)
        
        # 计算并绘制每个Treatment Group的平均值和标准偏差
        treatment_group_stats = df_mean.groupby('Treatment Group')['TRAP/DAPI'].agg(['mean', 'std', 'count'])
        for i, treatment_group in enumerate(treatment_group_stats.index):
            mean = treatment_group_stats.loc[treatment_group, 'mean']
            std = treatment_group_stats.loc[treatment_group, 'std']
            plt.plot([i-0.2, i+0.2], [mean, mean], color='black', lw=2)  # mean line
            plt.fill_between([i-0.2, i+0.2], mean-std, mean+std, color='gray', alpha=0.2)  # std deviation area

        # 打印每个Treatment Group的样本数
        print(f"Sample counts in {sheet_name}:\n")
        print(treatment_group_stats['count'])
        print("\n"+"="*40+"\n")

        plt.title(f'Average TRAP/DAPI by Treatment Group in {sheet_name} (n={df_mean["Treatment Group"].value_counts().to_dict()})')
        plt.ylabel('Average TRAPed/DAPI (%)')
        plt.ylim(0, 6)  # 设置y轴的范围
        plt.grid(True)
        plt.legend(title='Treatment Group')

        # 执行ANOVA
        fvalue, pvalue = f_oneway(df_mean.loc[df_mean['Treatment Group'] == "Oral HC", 'TRAP/DAPI'],
                                  df_mean.loc[df_mean['Treatment Group'] == "Oral OF", 'TRAP/DAPI'],
                                  df_mean.loc[df_mean['Treatment Group'] == "IP HC", 'TRAP/DAPI'],
                                  df_mean.loc[df_mean['Treatment Group'] == "IP OF", 'TRAP/DAPI'])
        print(f"ANOVA results for {sheet_name}: F = {fvalue}, p = {pvalue}")

        # 如果p值<0.05，进行Tukey HSD多重比较
        if pvalue < 0.05:
            mc = MultiComparison(df_mean['TRAP/DAPI'], df_mean['Treatment Group'])
            tukey_result = mc.tukeyhsd()
            print(tukey_result)
            
            # 添加显著性标记
            significance_dict = {
                '***': 0.001,  # 显著性水平p<0.001
                '**': 0.01,  # 显著性水平p<0.01
                '*': 0.05  # 显著性水平p<0.05
            }
            for i, treatment_group1 in enumerate(treatment_group_stats.index):
                for j, treatment_group2 in enumerate(treatment_group_stats.index[i+1:], start=i+1):
                    p_ij = tukey_result.pvalues[mc.groupsunique.tolist().index(treatment_group1)]
                    for symbol, significance_level in significance_dict.items():
                        if p_ij < significance_level:
                            plt.plot([i, j], [5.8, 5.8], color='black', lw=1)  # 添加连线
                            plt.text((i+j)/2, 5.8, symbol, ha='center')  # 添加显著性标记
                            break

        # 保存图片
        plt.savefig(f"/Users/sc17237/Desktop/{sheet_name}_TRAP_DAPI_Percentage_Average.png")
        plt.show()


# In[18]:


#TRAP/DAPI graphs with ANOVA

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from scipy.stats import f_oneway
from statsmodels.stats.multicomp import pairwise_tukeyhsd, MultiComparison

# Excel文件路径
file_path = "/Users/sc17237/Desktop/QuPath_ABBA_output_trap_sorted_by_area.xlsx"

# 输出文件路径
output_path = "/Users/sc17237/Desktop/TRAP_DAPI_Percentage_Average.xlsx"

# 创建一个ExcelWriter对象，用于输出到Excel
with pd.ExcelWriter(output_path) as writer:
    # 读取所有的sheets
    all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')

    # 按照sheet进行遍历
    for sheet_name, df in all_sheets.items():
        print(f"Processing {sheet_name}...")
        
        # 计算每一只小鼠的TRAP/DAPI percentage
        df['TRAP/DAPI'] = df['TRAP Cell Count'] / df['DAPI Cell Count'] * 100

        # 计算每个Animal Number的TRAP/DAPI平均值
        df_mean = df.groupby(['Animal Number', 'Treatment Group'])['TRAP/DAPI'].mean().reset_index()

        # 将结果写入Excel
        df_mean.to_excel(writer, sheet_name=sheet_name, index=False)

        # 打印TRAP/DAPI百分比表格
        print(f"\nAverage TRAP/DAPI percentages in {sheet_name}:\n")
        print(df_mean)
        print("\n"+"="*40+"\n")

        # 设置Treatment Group的顺序
        df_mean['Treatment Group'] = pd.Categorical(df_mean['Treatment Group'], ["Oral HC", "Oral OF", "IP HC", "IP OF"])

        # 为每个Treatment Group设定颜色
        colors = {"Oral HC": "blue", "Oral OF": "orange", "IP HC": "green", "IP OF": "red"}

        # 绘制散点图
        plt.figure(figsize=(10, 6))
        sns.scatterplot(data=df_mean, x='Treatment Group', y='TRAP/DAPI', hue='Treatment Group', palette=colors, s=100)
        
        # 计算并绘制每个Treatment Group的平均值和标准偏差
        treatment_group_stats = df_mean.groupby('Treatment Group')['TRAP/DAPI'].agg(['mean', 'std', 'count'])
        for i, treatment_group in enumerate(treatment_group_stats.index):
            mean = treatment_group_stats.loc[treatment_group, 'mean']
            std = treatment_group_stats.loc[treatment_group, 'std']
            plt.plot([i-0.2, i+0.2], [mean, mean], color='black', lw=2)  # mean line
            plt.fill_between([i-0.2, i+0.2], mean-std, mean+std, color='gray', alpha=0.2)  # std deviation area

        print(f"Sample count in {sheet_name}:\n")
        print(treatment_group_stats['count'])
        print("\n"+"="*40+"\n")

        plt.title(f'Average TRAP/DAPI by Treatment Group in {sheet_name} (n={df_mean["Treatment Group"].value_counts().to_dict()})')
        plt.ylabel('Average TRAPed/DAPI (%)')
        plt.ylim(0, 6)  # 设置y轴的范围
        plt.grid(True)
        plt.legend(title='Treatment Group')

        # 执行ANOVA
        fvalue, pvalue = f_oneway(df_mean.loc[df_mean['Treatment Group'] == "Oral HC", 'TRAP/DAPI'],
                                  df_mean.loc[df_mean['Treatment Group'] == "Oral OF", 'TRAP/DAPI'],
                                  df_mean.loc[df_mean['Treatment Group'] == "IP HC", 'TRAP/DAPI'],
                                  df_mean.loc[df_mean['Treatment Group'] == "IP OF", 'TRAP/DAPI'])
        print(f"ANOVA results for {sheet_name}: F = {fvalue}, p = {pvalue}")

        # 如果p值<0.05，进行Tukey HSD多重比较
        if pvalue < 0.05:
            mc = MultiComparison(df_mean['TRAP/DAPI'], df_mean['Treatment Group'])
            tukey_result = mc.tukeyhsd()
            print(tukey_result)
            
            # 添加显著性标记
            significance_dict = {
                '***': 0.001,  # 显著性水平p<0.001
                '**': 0.01,  # 显著性水平p<0.01
                '*': 0.05  # 显著性水平p<0.05
            }
            for i, treatment_group1 in enumerate(treatment_group_stats.index):
                for j, treatment_group2 in enumerate(treatment_group_stats.index[i+1:], start=i+1):
                    p_ij = tukey_result.pvalues[mc.groupsunique.tolist().index(treatment_group1)]
                    for symbol, significance_level in significance_dict.items():
                        if p_ij < significance_level:
                            # 根据比较组的数量设定显著性标记的高度
                            sign_height = 4.5 + 0.5 * (j-i-1)
                            # 添加连线
                            plt.plot([i, i+0.1, j-0.1, j], [sign_height, sign_height+0.2, sign_height+0.2, sign_height], color='black', lw=1)
                            plt.text((i+j)/2, sign_height+0.2, symbol, ha='center')  # 添加显著性标记
                            break

        # 保存图片p_ij = tukey_result.pvalues[mc.groupsunique.tolist().index(treatment_group1)][mc.groupsunique.tolist().index(treatment_group2)]

        plt.savefig(f"/Users/sc17237/Desktop/{sheet_name}_TRAP_DAPI_Percentage_Average.png")
        plt.show()


# In[43]:


#TRAP/DAPI graphs with ANOVA

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from scipy.stats import f_oneway
from statsmodels.stats.multicomp import pairwise_tukeyhsd, MultiComparison

# Excel文件路径
file_path = "/Users/sc17237/Desktop/QuPath_ABBA_output_trap_sorted_by_area.xlsx"

# 输出文件路径
output_path = "/Users/sc17237/Desktop/TRAP_DAPI_Percentage_Average.xlsx"

# 创建一个ExcelWriter对象，用于输出到Excel
with pd.ExcelWriter(output_path) as writer:
    # 读取所有的sheets
    all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')

    # 按照sheet进行遍历
    for sheet_name, df in all_sheets.items():
        print(f"Processing {sheet_name}...")
        
        # 计算每一只小鼠的TRAP/DAPI percentage
        df['TRAP/DAPI'] = df['TRAP Cell Count'] / df['DAPI Cell Count'] * 100

        # 计算每个Animal Number的TRAP/DAPI平均值
        df_mean = df.groupby(['Animal Number', 'Treatment Group'])['TRAP/DAPI'].mean().reset_index()

        # 将结果写入Excel
        df_mean.to_excel(writer, sheet_name=sheet_name, index=False)

        # 打印TRAP/DAPI百分比表格
        print(f"\nAverage TRAP/DAPI percentages in {sheet_name}:\n")
        print(df_mean)
        print("\n"+"="*40+"\n")

        # 设置Treatment Group的顺序
        df_mean['Treatment Group'] = pd.Categorical(df_mean['Treatment Group'], ["Oral HC", "Oral OF", "IP HC", "IP OF"])

        # 为每个Treatment Group设定颜色
        colors = {"Oral HC": "blue", "Oral OF": "orange", "IP HC": "green", "IP OF": "red"}

        # 绘制散点图
        plt.figure(figsize=(10, 6))
        sns.scatterplot(data=df_mean, x='Treatment Group', y='TRAP/DAPI', hue='Treatment Group', palette=colors, s=100)
        
        # 计算并绘制每个Treatment Group的平均值和标准偏差
        treatment_group_stats = df_mean.groupby('Treatment Group')['TRAP/DAPI'].agg(['mean', 'std', 'count'])
        for i, treatment_group in enumerate(treatment_group_stats.index):
            mean = treatment_group_stats.loc[treatment_group, 'mean']
            std = treatment_group_stats.loc[treatment_group, 'std']
            plt.plot([i-0.2, i+0.2], [mean, mean], color='black', lw=2)  # mean line
            plt.fill_between([i-0.2, i+0.2], mean-std, mean+std, color='gray', alpha=0.2)  # std deviation area

        print(f"Sample count in {sheet_name}:\n")
        print(treatment_group_stats['count'])
        print("\n"+"="*40+"\n")

        plt.title(f'Average TRAP/DAPI by Treatment Group in {sheet_name} (n={df_mean["Treatment Group"].value_counts().to_dict()})')
        plt.ylabel('Average TRAPed/DAPI (%)')
        plt.ylim(0, 6)  # 设置y轴的范围
        plt.grid(True)
        plt.legend(title='Treatment Group')

        # 执行ANOVA
        fvalue, pvalue = f_oneway(df_mean.loc[df_mean['Treatment Group'] == "Oral HC", 'TRAP/DAPI'],
                                  df_mean.loc[df_mean['Treatment Group'] == "Oral OF", 'TRAP/DAPI'],
                                  df_mean.loc[df_mean['Treatment Group'] == "IP HC", 'TRAP/DAPI'],
                                  df_mean.loc[df_mean['Treatment Group'] == "IP OF", 'TRAP/DAPI'])
        print(f"ANOVA results for {sheet_name}: F = {fvalue}, p = {pvalue}")

        # 如果p值<0.05，进行Tukey HSD多重比较
        if pvalue < 0.05:
            mc = MultiComparison(df_mean['TRAP/DAPI'], df_mean['Treatment Group'])
            tukey_result = mc.tukeyhsd()
            print(tukey_result)
            
            # 添加显著性标记
            significance_dict = {
                '***': 0.001,  # 显著性水平p<0.001
                '**': 0.01,  # 显著性水平p<0.01
                '*': 0.05  # 显著性水平p<0.05
            }
            for i, treatment_group1 in enumerate(treatment_group_stats.index):
                for j, treatment_group2 in enumerate(treatment_group_stats.index[i+1:], start=i+1):
                    p_ij = tukey_result.pvalues[mc.groupsunique.tolist().index(treatment_group1)][mc.groupsunique.tolist().index(treatment_group2)]
                    for symbol, significance_level in significance_dict.items():
                        if p_ij < significance_level:
                            # 根据比较组的数量设定显著性标记的高度
                            sign_height = 4.5 + 0.5 * (j-i-1)
                            # 添加连线
                            plt.plot([i, i+0.1, j-0.1, j], [sign_height, sign_height+0.2, sign_height+0.2, sign_height], color='black', lw=1)
                            plt.text((i+j)/2, sign_height+0.2, symbol, ha='center')  # 添加显著性标记
                            break

        # 保存图片
        plt.savefig(f"/Users/sc17237/Desktop/{sheet_name}_TRAP_DAPI_Percentage_Average.png")
        plt.show()


# In[73]:


#plot TRAPed cell per mm2

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

# Excel文件路径
file_path = "/Users/sc17237/Desktop/除掉了3.xlsx"

# 读取所有的sheets
all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')

# 按照sheet进行遍历
for sheet_name, df in all_sheets.items():
    print(f"Processing {sheet_name}...")
    
    # 计算每一只小鼠的TRAP Cell Count/Area，把Area单位从um²转化为mm²
    df['Area mm²'] = df['Area'] * 1e-6
    df['TRAPed cells/mm²'] = df['TRAP Cell Count'] / df['Area mm²']
    
    # 计算每个Animal Number的TRAPed cells/mm²平均值
    df_mean = df.groupby(['Animal Number', 'Treatment Group'])['TRAPed cells/mm²'].mean().reset_index()
    
    # 打印TRAPed cells/mm²平均值表格
    print(f"\nAverage TRAPed cells/mm² in {sheet_name}:\n")
    print(df_mean)
    print("\n"+"="*40+"\n")
    
    # 设置Treatment Group的顺序
    df_mean['Treatment Group'] = pd.Categorical(df_mean['Treatment Group'], ["Oral HC", "Oral OF", "IP HC", "IP OF"])
    
    # 为每个Treatment Group设定颜色
    colors = {"Oral HC": "blue", "Oral OF": "orange", "IP HC": "green", "IP OF": "red"}
    
    # 绘制散点图
    plt.figure(figsize=(10, 6))
    sns.scatterplot(data=df_mean, x='Treatment Group', y='TRAPed cells/mm²', hue='Treatment Group', palette=colors, s=100)
    
    # 计算并绘制每个Treatment Group的平均值和标准偏差
    treatment_group_stats = df_mean.groupby('Treatment Group')['TRAPed cells/mm²'].agg(['mean', 'std'])
    for i, treatment_group in enumerate(treatment_group_stats.index):
        mean = treatment_group_stats.loc[treatment_group, 'mean']
        std = treatment_group_stats.loc[treatment_group, 'std']
        plt.plot([i-0.2, i+0.2], [mean, mean], color='black', lw=2)  # mean line
        plt.fill_between([i-0.2, i+0.2], mean-std, mean+std, color='gray', alpha=0.2)  # std deviation area

    plt.title(f'Average TRAPed cells/mm² by Treatment Group in {sheet_name}')
    plt.ylabel('Average TRAPed cells/mm²')
    plt.ylim(0, 250)  # 设置y轴的范围
    plt.grid(True)
    plt.legend(title='Treatment Group')

    # 保存图片
    plt.savefig(f"/Users/sc17237/Desktop/{sheet_name}_TRAPed_cells_per_mm2.png")
    plt.show()


# In[46]:


import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from scipy.stats import f_oneway
from statsmodels.stats.multicomp import pairwise_tukeyhsd, MultiComparison

# Excel文件路径
file_path = "/Users/sc17237/Desktop/QuPath_ABBA_output_trap_sorted_by_area.xlsx"

# 输出文件路径
output_path = "/Users/sc17237/Desktop/TRAP_DAPI_Percentage_Average.xlsx"

# 创建一个ExcelWriter对象，用于输出到Excel
with pd.ExcelWriter(output_path) as writer:
    # 读取所有的sheets
    all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')

    # 按照sheet进行遍历
    for sheet_name, df in all_sheets.items():
        print(f"Processing {sheet_name}...")
        
        # 计算每一只小鼠的TRAP/DAPI percentage
        df['TRAP/DAPI'] = df['TRAP Cell Count'] / df['DAPI Cell Count'] * 100

        # 计算每个Animal Number的TRAP/DAPI平均值
        df_mean = df.groupby(['Animal Number', 'Treatment Group'])['TRAP/DAPI'].mean().reset_index()

        # 将结果写入Excel
        df_mean.to_excel(writer, sheet_name=sheet_name, index=False)

        # 打印TRAP/DAPI百分比表格
        print(f"\nAverage TRAP/DAPI percentages in {sheet_name}:\n")
        print(df_mean)
        print("\n"+"="*40+"\n")

        # 设置Treatment Group的顺序
        df_mean['Treatment Group'] = pd.Categorical(df_mean['Treatment Group'], ["Oral HC", "Oral OF", "IP HC", "IP OF"])

        # 为每个Treatment Group设定颜色
        colors = {"Oral HC": "blue", "Oral OF": "orange", "IP HC": "green", "IP OF": "red"}

        # 绘制散点图
        plt.figure(figsize=(10, 6))
        sns.scatterplot(data=df_mean, x='Treatment Group', y='TRAP/DAPI', hue='Treatment Group', palette=colors, s=100)
        
        # 计算并绘制每个Treatment Group的平均值和标准偏差
        treatment_group_stats = df_mean.groupby('Treatment Group')['TRAP/DAPI'].agg(['mean', 'std', 'count'])
        for i, treatment_group in enumerate(treatment_group_stats.index):
            mean = treatment_group_stats.loc[treatment_group, 'mean']
            std = treatment_group_stats.loc[treatment_group, 'std']
            plt.plot([i-0.2, i+0.2], [mean, mean], color='black', lw=2)  # mean line
            plt.fill_between([i-0.2, i+0.2], mean-std, mean+std, color='gray', alpha=0.2)  # std deviation area

        print(f"Sample count in {sheet_name}:\n")
        print(treatment_group_stats['count'])
        print("\n"+"="*40+"\n")

        plt.title(f'Average TRAP/DAPI by Treatment Group in {sheet_name} (n={df_mean["Treatment Group"].value_counts().to_dict()})')
        plt.ylabel('Average TRAPed/DAPI (%)')
        plt.ylim(0, 6)  # 设置y轴的范围
        plt.grid(True)
        plt.legend(title='Treatment Group')

        # 执行ANOVA
        fvalue, pvalue = f_oneway(df_mean.loc[df_mean['Treatment Group'] == "Oral HC", 'TRAP/DAPI'],
                                  df_mean.loc[df_mean['Treatment Group'] == "Oral OF", 'TRAP/DAPI'],
                                  df_mean.loc[df_mean['Treatment Group'] == "IP HC", 'TRAP/DAPI'],
                                  df_mean.loc[df_mean['Treatment Group'] == "IP OF", 'TRAP/DAPI'])
        print(f"ANOVA results for {sheet_name}: F = {fvalue}, p = {pvalue}")

        # 如果p值<0.05，进行Tukey HSD多重比较
        if pvalue < 0.05:
            mc = MultiComparison(df_mean['TRAP/DAPI'], df_mean['Treatment Group'])
            tukey_result = mc.tukeyhsd()
            print(tukey_result)
            
            # 添加显著性标记
            significance_dict = {
                '***': 0.001,  # 显著性水平p<0.001
                '**': 0.01,  # 显著性水平p<0.01
                '*': 0.05  # 显著性水平p<0.05
            }
            
            # 为每一对配对生成一个列表
            pairs = tukey_result._results_table.data[1:]  # each element is a list of [group1, group2, meandiff, pval, lower, upper, reject]

            for pair in pairs:
                group1, group2, _, p_ij, _, _, reject = pair
                if not reject:  # if not significant, skip this pair
                    continue
                i = treatment_group_stats.index.to_list().index(group1)
                j = treatment_group_stats.index.to_list().index(group2)
                for symbol, significance_level in significance_dict.items():
                    if p_ij < significance_level:
                        # 根据比较组的数量设定显著性标记的高度
                        sign_height = 4. + 0.5 * (j-i-1)
                        # 添加连线
                        plt.plot([i, i+0.1, j-0.1, j], [sign_height, sign_height+0.2, sign_height+0.2, sign_height], color='black', lw=1)
                        plt.text((i+j)/2, sign_height+0.2, symbol, ha='center')  # 添加显著性标记
                        break

        # 保存图片
        plt.savefig(f"/Users/sc17237/Desktop/{sheet_name}_TRAP_DAPI_Percentage_Average.png")
        plt.show()


# In[72]:


#create prism-friendly excel file
import pandas as pd

# 文件路径
file_path = '/Users/sc17237/Desktop/TRAP_DAPI_Percentage_Average.xlsx'

# 需要读取的sheets
sheets = ['ACA', 'HPF', 'LH', 'PVT', 'PVZ']

# 创建一个ExcelWriter对象，用于写入多个sheets
writer = pd.ExcelWriter('/Users/sc17237/Desktop/TRAP_DAPI_Percentage_Average_converted.xlsx', engine='openpyxl')

# 遍历每一个sheet
for sheet in sheets:
    # 读取Excel文件的sheet
    df = pd.read_excel(file_path, sheet_name=sheet)

    # 将 'Treatment Group' 栏位转为四个单独的栏位
    df['Oral HC'] = df.loc[df['Treatment Group'] == 'Oral HC', 'TRAP/DAPI']
    df['Oral OF'] = df.loc[df['Treatment Group'] == 'Oral OF', 'TRAP/DAPI']
    df['IP HC'] = df.loc[df['Treatment Group'] == 'IP HC', 'TRAP/DAPI']
    df['IP OF'] = df.loc[df['Treatment Group'] == 'IP OF', 'TRAP/DAPI']

    # 删除 'Treatment Group' 和 'TRAP/DAPI'栏位
    df = df.drop(columns=['Treatment Group', 'TRAP/DAPI'])

    # 根据 'Animal Number' 对数据进行分组，并求平均值
    df = df.groupby('Animal Number').mean().reset_index()

    # 将处理后的数据写入新的Excel文件的对应sheet
    df.to_excel(writer, sheet_name=sheet, index=False)

# 保存Excel文件
writer.close()


# In[74]:


import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

# Excel文件路径
file_path = "/Users/sc17237/Desktop/除掉了3.xlsx"

# 读取所有的sheets
all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')

# 创建一个ExcelWriter对象，准备写入一个新的Excel文件
writer = pd.ExcelWriter('/Users/sc17237/Desktop/TRAPed_per_mm2.xlsx', engine='openpyxl')

# 按照sheet进行遍历
for sheet_name, df in all_sheets.items():
    print(f"Processing {sheet_name}...")
    
    # 计算每一只小鼠的TRAP Cell Count/Area，把Area单位从um²转化为mm²
    df['Area mm²'] = df['Area'] * 1e-6
    df['TRAPed cells/mm²'] = df['TRAP Cell Count'] / df['Area mm²']
    
    # 计算每个Animal Number的TRAPed cells/mm²平均值
    df_mean = df.groupby(['Animal Number', 'Treatment Group'])['TRAPed cells/mm²'].mean().reset_index()
    
    # 打印TRAPed cells/mm²平均值表格
    print(f"\nAverage TRAPed cells/mm² in {sheet_name}:\n")
    print(df_mean)
    print("\n"+"="*40+"\n")
    
    # 将平均值数据写入新的Excel文件
    df_mean.to_excel(writer, sheet_name=sheet_name, index=False)
    
    # 设置Treatment Group的顺序
    df_mean['Treatment Group'] = pd.Categorical(df_mean['Treatment Group'], ["Oral HC", "Oral OF", "IP HC", "IP OF"])
    
    # 为每个Treatment Group设定颜色
    colors = {"Oral HC": "blue", "Oral OF": "orange", "IP HC": "green", "IP OF": "red"}
    
    # 绘制散点图
    plt.figure(figsize=(10, 6))
    sns.scatterplot(data=df_mean, x='Treatment Group', y='TRAPed cells/mm²', hue='Treatment Group', palette=colors, s=100)
    
    # 计算并绘制每个Treatment Group的平均值和标准偏差
    treatment_group_stats = df_mean.groupby('Treatment Group')['TRAPed cells/mm²'].agg(['mean', 'std'])
    for i, treatment_group in enumerate(treatment_group_stats.index):
        mean = treatment_group_stats.loc[treatment_group, 'mean']
        std = treatment_group_stats.loc[treatment_group, 'std']
        plt.plot([i-0.2, i+0.2], [mean, mean], color='black', lw=2)  # mean line
        plt.fill_between([i-0.2, i+0.2], mean-std, mean+std, color='gray', alpha=0.2)  # std deviation area

    plt.title(f'Average TRAPed cells/mm² by Treatment Group in {sheet_name}')
    plt.ylabel('Average TRAPed cells/mm²')
    plt.ylim(0, 250)  # 设置y轴的范围
    plt.grid(True)
    plt.legend(title='Treatment Group')

    # 保存图片
    plt.savefig(f"/Users/sc17237/Desktop/{sheet_name}_TRAPed_cells_per_mm2.png")
    plt.show()

# 关闭ExcelWriter并保存Excel文件
writer.close()


# In[76]:


#sorting the file for prism
import pandas as pd

# 读取Excel文件
file_path = "/Users/sc17237/Desktop/TRAPed_per_mm2.xlsx"
data = pd.read_excel(file_path, sheet_name=None)

# 创建一个新的Excel文件
output_file_path = "/Users/sc17237/Desktop/Transformed_TRAPed_per_mm2.xlsx"
writer = pd.ExcelWriter(output_file_path, engine="xlsxwriter")

# 定义Treatment Group的顺序
treatment_order = ["Oral HC", "Oral OF", "IP HC", "IP OF"]

# 遍历每个sheet
for sheet_name, sheet_data in data.items():
    # 将Treatment Group的值作为列名，创建新的DataFrame，并按照顺序重排列
    transformed_data = pd.pivot_table(sheet_data, index="Animal Number", columns="Treatment Group", values="TRAPed cells/mm²")
    transformed_data = transformed_data[treatment_order]

    # 将新的DataFrame写入到新的Excel文件中的不同sheet中
    transformed_data.to_excel(writer, sheet_name=sheet_name)

# 保存并关闭Excel文件
writer.close()



# In[86]:


#trap/dapi treatment group
import pandas as pd

# 创建一个Excel writer对象
with pd.ExcelWriter('/Users/sc17237/Desktop/processed_output.xlsx') as writer:
    # 处理每个sheet
    for sheet_name in ['IP HC', 'IP OF', 'Oral HC', 'Oral OF']:
        # 读取每个sheet
        df = pd.read_excel('/Users/sc17237/Desktop/sorted_by_Treatmentgroup.xlsx', sheet_name=sheet_name)
        
        # 计算每只小鼠在每一个Area Name的TRAP Cell Count和DAPI Cell Count的总和
        df_sum = df.groupby(['Animal Number', 'Area Name']).agg({'TRAP Cell Count': 'sum', 'DAPI Cell Count': 'sum'}).reset_index()

        # 计算TRAP/DAPI percentage
        df_sum['TRAP/DAPI Percentage'] = (df_sum['TRAP Cell Count'] / df_sum['DAPI Cell Count']) * 100
        
        # 将数据重新格式化，以使每个Area Name拥有一个单独的column
        df_pivot = df_sum.pivot(index='Animal Number', columns='Area Name', values='TRAP/DAPI Percentage')

        # 写入新的Excel文件
        df_pivot.to_excel(writer, sheet_name=sheet_name)

print('操作完成，已生成新的Excel文件。')



# In[85]:


#traped per mm2 treatment group
import pandas as pd

# 打开已存在的xlsx文件
xlsx = pd.ExcelWriter('/Users/sc17237/Desktop/TRAPed_per_mm2.xlsx', engine='openpyxl')

# 读取excel文件中的所有sheet，返回一个字典，键是sheet名，值是对应的DataFrame
all_sheets = pd.read_excel('/Users/sc17237/Desktop/sorted_by_Treatmentgroup.xlsx', sheet_name=None)

for sheet_name, df in all_sheets.items():
    # 计算每一个项的TRAPed cells per mm²，这里将um²转换为mm²，所以需要乘以1e-6
    df['Area mm²'] = df['Area'] * 1e-6
    df['TRAPed cells/mm²'] = df['TRAP Cell Count'] / df['Area mm²']
    
    # 按照Animal Number和Area Name分组，并计算每一组的TRAPed cells per mm²的平均值
    df_mean = df.groupby(['Animal Number', 'Area Name'])['TRAPed cells/mm²'].mean().reset_index()
    
    # 将计算结果reshape，转换为每个Animal Number一个行，每个Area Name一个列的形式
    df_pivot = df_mean.pivot(index='Animal Number', columns='Area Name', values='TRAPed cells/mm²')
    
    # 将计算结果写入新的excel文件的sheet
    df_pivot.to_excel(xlsx, sheet_name=sheet_name)

# 保存并关闭xlsx文件
xlsx.close()



# In[ ]:




