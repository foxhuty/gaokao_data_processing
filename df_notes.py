# -*_ codeing=utf-8 -*-
# @Time: 2024/11/8 11:07
# @Author: foxhuty
# @File: df_notes.py
# @Software: PyCharm
# @Based on python 3.13
import pandas as pd
import numpy as np

file_path = r'D:\data_test\高2026级学生10月考成绩汇总+++成绩分析统计结果.xlsx'
data = pd.read_excel(file_path)
print(data.head())
subjects = [col for col in data.columns if
            col in ['语文', '数学', '英语', '物理', '历史',
                    '化学', '政治', '地理', '生物',
                    '化学赋分', '政治赋分', '地理赋分', '生物赋分', '总分', '总分赋分']]


def subjects_average(data, subjects_av):
    class_av = data.groupby('班级')[subjects_av].mean().round(2)
    av_general = data[subjects_av].apply(np.nanmean, axis=0).round(2)
    # av_general = data[subjects_av].mean().round(2)
    av_percentage = class_av / av_general.round(2)
    # pandas 2.0以上用map替换applymap
    av_percentage = av_percentage.map(lambda x: format(x, '.2%'))  # 以百分号显示
    av_percentage = av_percentage.map(lambda x: x.replace('nan%', ''))  # 不显示nan%
    # print(av_percentage.columns)
    av_percentage_cols = [col + '占比' for col in av_percentage.columns]
    col_dict = dict(zip(av_percentage.columns, av_percentage_cols))
    av_percentage.rename(columns=col_dict, inplace=True)
    # print(av_percentage.columns)
    class_av.loc['年级平均'] = av_general
    class_av['参考人数'] = data['班级'].value_counts()
    class_av.loc['年级平均', '参考人数'] = class_av['参考人数'].sum()
    final_av_percentage = pd.concat([class_av, av_percentage], axis=1)
    final_av_percentage.to_excel('final_av.xlsx')
    # final_av_percentage = self.change_columns_order(final_av_percentage)
    return final_av_percentage


def get_average_school(self):
    # data = self.get_grade_data()[0]

    subjects_av = [col for col in data.columns if
                   col in ['语文', '数学', '英语', '物理', '历史',
                           '化学', '政治', '地理', '生物',
                           '化学赋分', '政治赋分', '地理赋分', '生物赋分', '总分', '总分赋分']]
    # print(subjects_av)
    class_av = data.groupby('班级')[subjects_av].mean().round(2)
    # print(class_av)
    av_general = data[subjects_av].apply(np.nanmean, axis=1).round(2)
    print(av_general)
    final_av_percentage = subjects_average(data, subjects_av)
    print(final_av_percentage)
    print(final_av_percentage.columns)
if __name__ == '__main__':



    get_average_school(data)
