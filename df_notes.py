# -*_ codeing=utf-8 -*-
# @Time: 2024/11/8 11:07
# @Author: foxhuty
# @File: df_notes.py
# @Software: PyCharm
# @Based on python 3.13
import pandas as pd
import numpy as np

file_path = r'D:\data_test\高2026级学生10月考+成绩 赋分统计.xlsx'
df = pd.read_excel(file_path)
print(df.head())
subjects = [col for col in df.columns if
            col in ['语文', '数学', '英语', '物理', '历史',
                    '化学', '政治', '地理', '生物',
                    '化学赋分', '政治赋分','地理赋分','生物赋分', '总分', '总分赋分']]
print(subjects)

