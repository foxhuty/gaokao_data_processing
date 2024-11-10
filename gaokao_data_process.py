# -*_ codeing=utf-8 -*-
# @Time: 2024/11/8 10:56
# @Author: foxhuty
# @File: gaokao_data_process.py
# @Software: PyCharm
# @Based on python 3.13

import pandas as pd
from sympy import symbols, solve
import numpy as np
import os
import time


class GaokaoData2025:
    """
    从2025年起四川采用新高考（3+1+2)模式。这是专门针对新高考模式的政治, 地理, 生物, 化学四门赋分学科而设计的程序，
    以帮助高中年级老师做成绩分析用。还可以计算各班平均分及占比，有效分统计等。
    该程序设计用于帮助高中教师分析新高考模式下的学生成绩，特别是政治、地理、生物和化学四门学科的赋分情况。
    程序可以读取Excel文件中的成绩数据，计算各学科的赋分等级和赋分值，统计班级平均分，各科有效分并生成包含这些分析结果的Excel文件。
    """
    # 各个等级赋分区间（T1——T2）
    A_T_range = [86, 100]
    B_T_range = [71, 85]
    C_T_range = [56, 70]
    D_T_range = [41, 55]
    E_T_range = [30, 40]
    # A_T_range = [76, 90]
    # B_T_range = [61, 75]
    # C_T_range = [46, 60]
    # D_T_range = [31, 45]
    # E_T_range = [20, 30]

    subjects_good_scores_all = {'语文': None, '数学': None, '英语': None, '物理': None, '历史': None, '政治': None,
                                '地理': None, '化学': None, '生物': None, '总分': None}
    subjects_good_scores_physics = {'语文': None, '数学': None, '英语': None, '物理': None, '化学': None,
                                    '生物': None, '政治': None, '地理': None, '总分': None}
    subjects_good_scores_history = {'语文': None, '数学': None, '英语': None, '历史': None, '生物': None,
                                    '政治': None, '地理': None, '总分': None}
    total_line = None

    def __init__(self, file):
        if not os.path.isfile(file) or not file.endswith('.xlsx'):
            raise FileNotFoundError('文件路径无效')
        self.file = file
        try:
            excel_file = pd.ExcelFile(self.file)
        except FileNotFoundError as e:
            raise IOError(f'无法读取文件：{e}')
        self.sheet_names = [sheet_name for sheet_name in excel_file.sheet_names if
                            sheet_name in ['物理类', '历史类', '总表']]
        self.data_list = [pd.read_excel(self.file, sheet_name=sheet_name, dtype={'考号': str, '考生号': str}) for
                          sheet_name in
                          self.sheet_names]

    def __str__(self):
        return os.path.basename(self.file)

    def get_grade_data(self):
        """
        分别获取政治，地理，化学和生物学科的赋分等级和赋分值后生成excel数据表
        :return: None
        """
        data = self.data_list
        # data=self.get_mixed_data()
        data_added = []
        for item in range(len(data)):
            subjects = [col for col in data[item].columns if col in ['政治', '地理', '生物', '化学']]
            print(f'{self.sheet_names[item]}各学科等级人数及卷面分值区间'.center(100, '*'))
            subject_added_data = []
            for subject in subjects:
                data_item = self.get_grade(data[item], subject)
                subject_added_data.append(data_item)
            subjects_added_data = subject_added_data[-1]
            # subjects_added_data = pd.concat(subject_added_data, ignore_index=False,join='inner', axis=1)
            subjects_added_data_cols = [col for col in subjects_added_data.columns if
                                        col in ['语文', '数学', '英语', '物理', '历史', '政治赋分', '地理赋分',
                                                '生物赋分', '化学赋分']]
            # 新版pandass要用astype(float).round(2)转换成2位小数。
            subjects_added_data['总分赋分'] = subjects_added_data[subjects_added_data_cols].sum(axis=1).astype(
                float).round(2)
            subjects_added_data.sort_values(by='总分赋分', ascending=False, inplace=True)
            subjects_added_data['序号'] = [i + 1 for i in range(len(subjects_added_data))]
            data_added.append(subjects_added_data)
        return data_added

    def get_data_processed(self):
        '''
        此函数还没有完成。目前还不能调用
        :return:
        '''
        data = self.get_grade_data()
        data_processed_list = []
        if len(data) == 1:
            data[0].drop(['化学', '生物', '政治', '地理', '总分', '化学等级', '地理等级', '生物等级', '政治等级'],
                         axis=1, inplace=True)
            data_all = data[0].copy()
            # data_all.columns.rename = [self.rename_columns(col, '赋值') for col in data_all.columns]
            data_all.columns.map(lambda x: self.rename_columns(x, '赋值'))

            print(data_all.columns)
            data_processed_list.append(data_all)
        else:
            for i in range(len(data)):
                if '化学' in data[i].columns:
                    data[i].drop(
                        ['化学', '生物', '政治', '地理', '总分', '化学等级', '地理等级', '生物等级', '政治等级'],
                        axis=1, inplace=True)
                    data_physics = data[i]
                    # data_physics.colunms = [self.rename_columns(col, '赋值') for col in data_physics.columns]
                    print(data_physics.columns)
                    data_processed_list.append(data_physics)

                else:
                    data[i].drop(['生物', '政治', '地理', '总分', '化学等级', '地理等级', '生物等级', '政治等级'],
                                 axis=1, inplace=True)
                    data_history = data[i]
                    data_history.colunms = [self.rename_columns(col, '赋值') for col in data_history.columns]
                    print(data_history.columns)
                    data_processed_list.append(data_history)
        return data_processed_list

    def get_grade(self, data, subject):
        """
        分别获取政治，地理，化学和生物学科的赋分等级和赋分值
        :param data: df数据表
        :param subject: 学科
        :return: 返回获取了赋分等级和赋分值的df数据
        """
        max_score, min_score = self.get_subject_max_min_score(data, subject)
        data[subject + '等级'] = data[subject].apply(lambda x: self.get_level(x, min_score))
        data[subject + '赋分'] = data[subject].apply(
            lambda x: self.get_final_scores(x, min_score, max_score))
        return data

    def get_mixed_data(self):
        '''
        此函数暂时没有调用。用处待定
        :return:
        '''
        df_list = []
        if '总表' in self.sheet_names:
            df_list.append(self.data_list[0])
            return self.data_list[0]
        else:
            data_physics = [data for data in self.data_list if '物理' in data.columns][0]
            data_history = [data for data in self.data_list if '历史' in data.columns][0]
            df_mixed = pd.concat([data_physics, data_history])
            # print(df_mixed.head())
            df_mixed.sort_values(by='总分', ascending=False, inplace=True)
            df_mixed['序号'] = [i + 1 for i in range(len(df_mixed))]
            # df_mixed.to_excel(r'D:\data_test\mixed_df.xlsx', sheet_name='总表', index=False)
            df_list.append(df_mixed)
            return df_list

    @staticmethod
    def separate_data(data_mixed):
        '''
        用于分科后考试，赋分后再生成物理类和历史类两个df，用于生成excel文件
        :param data_mixed:
        :return:
        '''

        history_min = data_mixed['历史'].min(skipna=True)
        physics_min = data_mixed['物理'].min(skipna=True)
        data_history = data_mixed[data_mixed['历史'] >= history_min]
        # 删除多列
        data_history = data_history.copy()
        data_history.drop(['物理', '化学', '化学等级', '化学赋分'], axis=1, inplace=True)

        data_physics = data_mixed[data_mixed['物理'] >= physics_min]
        # 删除单列
        del data_physics['历史']
        return data_physics, data_history

    def get_average(self):
        '''
        计算区级及以上考试的各学科平均分。不能用赋分后的df来计算
        :return: 返回一个有df(物理类，历史类)元素的列表
        '''
        # data = self.data_list
        final_av = []
        for data in self.data_list:
            subjects = [col for col in data.columns if
                        col in ['语文', '数学', '英语', '物理', '历史', '政治', '地理', '生物', '化学', '总分']]
            final_av_percentage = self.subjects_average(data, subjects)
            final_av.append(final_av_percentage)

        return final_av

    def subjects_average(self, data, subjects):
        '''
        计算各学科的平均分。
        :param data: df数据
        :param subjects: 参加计算平均分的学科
        :return: 返回一个df。包含有各班各科平均分，年级平均分和各科平均分在年级平均分中的占比，用于制作折线图
        '''
        class_av = data.groupby('班级')[subjects].mean().round(2)
        # 求这个平均数，只能用apply(np.mean)才行。用其它的如mean(),apply(np.nanmean)都会报错
        av_general = data[subjects].apply(np.mean, axis=0).round(2)
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
        final_av_percentage = self.change_columns_order(final_av_percentage)
        return final_av_percentage

    def get_average_school(self, data):
        '''
        计算赋分后作有学科的平均分及平均分占比
        :param data: df
        :return: 返回一个有df元素的列表。
        '''
        final_av = []
        subjects_av = [col for col in data.columns if
                       col in ['语文', '数学', '英语', '物理', '历史',
                               '化学', '政治', '地理', '生物',
                               '化学赋分', '政治赋分', '地理赋分', '生物赋分', '总分', '总分赋分']]

        final_av_percentage = self.subjects_average(data, subjects_av)
        final_av.append(final_av_percentage)
        return final_av

    def excel_files(self):
        '''
        区级及以上考试不用赋分,excel表上有物理类和历史类两个工作表（sheet)
        :return:
        '''
        # scores_added_data = self.get_grade_data()
        average_added = self.get_average()
        good_scores = self.good_scores()
        with pd.ExcelWriter(self.file.split('.xlsx')[0] + '---成绩分析统计结果.xlsx') as writer:
            for item in range(len(average_added)):
                # scores_added_data[item].to_excel(writer, sheet_name=f'{self.sheet_names[item]}-赋分表', index=False,
                #                                  float_format='%.2f')
                average_added[item].to_excel(writer, sheet_name=f'{self.sheet_names[item]}-平均分统计')
            if len(self.sheet_names) == 1:
                good_scores[0].to_excel(writer, sheet_name=f'{self.sheet_names[item]}--有效分')
            else:
                good_scores[0].to_excel(writer, sheet_name='物理类有效分统计')
                good_scores[1].to_excel(writer, sheet_name='历史类有效分统计')
        print(f'对{self.__str__()}文件数据分析处理完成')

    def excel_school_files(self):
        '''
        学校组织的考试，要对化学政治地理和生物四科赋分。excel文件上要有一个总表工作表（sheet)
        :return:
        '''

        scores_added_data_list = self.get_grade_data()
        with pd.ExcelWriter(self.file.split('.xlsx')[0] + '+++成绩分析统计结果.xlsx') as writer:
            for item in range(len(scores_added_data_list)):
                scores_added_data_list[item].to_excel(writer, sheet_name=f'{self.sheet_names[item]}-赋分表',
                                                      index=False,
                                                      float_format='%.2f')

                average_added = self.get_average_school(scores_added_data_list[item])
                average_added[item].to_excel(writer, sheet_name=f'{self.sheet_names[item]}-平均分统计')

                good_score_data = self.good_scores_school(scores_added_data_list[item])
                good_score_data.to_excel(writer, sheet_name=f'{self.sheet_names[item]}--有效分')
                # physics_data, history_data = self.separate_data(scores_added_data_list[item])
                # physics_data.to_excel(writer, sheet_name=f'{self.sheet_names[item]}-物理类')
                # history_data.to_excel(writer, sheet_name=f'{self.sheet_names[item]}-历史类')

    def get_final_scores(self, score, min_score, max_score):
        """
        计算获取学生卷面分数的赋分值
        :param score:
        :param min_score:
        :param max_score:
        :return:
        """
        a_t = GaokaoData2025.A_T_range
        b_t = GaokaoData2025.B_T_range
        c_t = GaokaoData2025.C_T_range
        d_t = GaokaoData2025.D_T_range
        e_t = GaokaoData2025.E_T_range
        if score >= min_score[0]:
            score_added = self.get_added_score(score, min_score[0], max_score[0], a_t[0], a_t[1])
            return score_added
        elif score >= min_score[1]:
            score_added = self.get_added_score(score, min_score[1], max_score[1], b_t[0], b_t[1])
            return score_added
        elif score >= min_score[2]:
            score_added = self.get_added_score(score, min_score[2], max_score[2], c_t[0], c_t[1])
            return score_added
        elif score >= min_score[3]:
            score_added = self.get_added_score(score, min_score[3], max_score[3], d_t[0], d_t[1])
            return score_added
        elif score < min_score[3]:
            score_added = self.get_added_score(score, min_score[4], max_score[4], e_t[0], e_t[1])
            return score_added

    def good_scores(self):
        '''
        用区级及以上考试的有效分统计，用所给的有效分数计算各班各科有效分上线情况
        :return: 返回一个含有效分统计的df元素列表
        '''

        if '总表' in self.sheet_names:
            single_double_list = []
            single_df_all, double_df_all = self.get_single_double_data(self.data_list[0],
                                                                       **GaokaoData2025.subjects_good_scores_all)
            single_double = pd.concat([single_df_all, double_df_all], axis=0, keys=['单有效', '双有效'])
            single_double_list.append(single_double)
            return single_double_list
        else:
            single_double_list = []
            data_physics = [data for data in self.data_list if '物理' in data.columns][0]
            single_df_physics, double_df_physics = self.get_single_double_data(data_physics,
                                                                               **GaokaoData2025.subjects_good_scores_physics)
            single_double_physics = pd.concat([single_df_physics, double_df_physics], axis=0,
                                              keys=['单有效', '双有效'])
            single_double_list.append(single_double_physics)
            data_history = [data for data in self.data_list if '历史' in data.columns][0]
            single_df_history, double_df_history = self.get_single_double_data(data_history,
                                                                               **GaokaoData2025.subjects_good_scores_history)
            # single_df_history.sort_index(ascending=True, inplace=True)
            single_double_history = pd.concat([single_df_history, double_df_history], axis=0,
                                              keys=['单有效', '双有效'])
            single_double_list.append(single_double_history)
            return single_double_list

    def good_scores_school(self, data):
        '''
        用于学校考试的有效分统计，有效分数是由指定的划线分数计算得到
        :param data: df
        :return: 返回一个有效分统计的df
        '''
        good_scores_data = self.get_single_double_school_data(data)
        return good_scores_data

    def get_single_double_data(self, data, **kwargs):
        '''
        按照已知有效分，计算各学科有效分上线数据。主要用于区级及以上考试的成绩数据分析
        :param data:
        :param kwargs: 以字典形式输入各学科的有效分数
        :return: 单有效和双有效统计数据
        '''
        subjects = [col for col in data.columns if
                    col in ['总分', '语文', '数学', '英语', '物理', '历史', '化学', '生物', '政治', '地理']]
        single_data_list = []
        double_data_list = []
        # total = kwargs['总分']
        for subject in subjects:
            single_data = self.get_single_subject_data(data, subject, kwargs[subject])
            double_data = self.get_double_subject_data(data, subject, '总分', kwargs[subject], kwargs['总分'])
            single_data_list.append(single_data)
            double_data_list.append(double_data)
        single_data = pd.concat(single_data_list, axis=1)
        double_data = pd.concat(double_data_list, axis=1)
        # 单有效统计：增加一列参考人数和一行年级共计
        single_data['参考人数'] = data['班级'].value_counts()
        single_data.loc['年级共计'] = [single_data[col].sum() for col in single_data.columns]
        single_data = self.change_columns_order(single_data)
        # 增加一列上线率并用百分号显示
        single_data['上线率'] = single_data['总分'] / single_data['参考人数']
        single_data['上线率'] = single_data['上线率'].apply(lambda x: format(x, '.2%'))  # 以百分号显示

        # 双有效统计：增加一列参考人数和一行年级共计
        double_data['参考人数'] = data['班级'].value_counts()
        double_data.loc['年级共计'] = [double_data[col].sum() for col in double_data.columns]
        double_data = self.change_columns_order(double_data)

        return single_data, double_data

    def get_single_double_school_data(self, data):
        """
         按照新高考的模式，计算中线和高线的有效分，完成单有效和双有效的统计
        :param data:
        :return: 返回含有有效分数，单有效和双有效的df

        """

        subjects = [col for col in data.columns if
                    col in ['语文', '数学', '英语', '物理', '历史', '化学', '政治', '地理', '生物',
                            '化学赋分', '政治赋分', '地理赋分', '生物赋分', '总分', '总分赋分']]
        single_data_list = []
        double_data_list = []
        subjects_name = []
        subjects_scores = []
        for subject in subjects:
            subject_good_score = self.get_subject_good_score(data, '总分赋分', subject)
            subjects_name.append(subject)
            subjects_scores.append(subject_good_score)

            single_data = self.get_single_subject_data(data, subject, subject_good_score)
            double_data = self.get_double_subject_data(data, subject, '总分赋分', subject_good_score,
                                                       GaokaoData2025.total_line)

            single_data_list.append(single_data)
            double_data_list.append(double_data)
        # 获取有效分数并转换成df
        good_scores_dict = dict(zip(subjects_name, subjects_scores))
        good_scores_df = pd.DataFrame(good_scores_dict, index=[0])
        # 合成单有效和双有效统计的df
        #
        single_data = pd.concat(single_data_list, axis=1)
        double_data = pd.concat(double_data_list, axis=1)

        # 计算错位人数
        unmatched_df = self.get_unmatched_data(double_data, '总分赋分')

        # 单有效统计：增加一列参考人数和一行年级共计
        single_data['参考人数'] = data['班级'].value_counts()
        single_data.loc['年级共计'] = [single_data[col].sum() for col in single_data.columns]

        # 增加一列上线率并用百分号显示
        single_data['上线率'] = single_data['总分赋分'] / single_data['参考人数']
        single_data['上线率'] = single_data['上线率'].apply(lambda x: format(x, '.2%'))  # 以百分号显示

        # 双有效统计：增加一列参考人数和一行年级共计
        double_data['参考人数'] = data['班级'].value_counts()
        double_data.loc['年级共计'] = [double_data[col].sum() for col in double_data.columns]

        combined_data = pd.concat([good_scores_df, single_data, double_data, unmatched_df],
                                  keys=['有效分数', '单有效', '双有效', '错位人数'],
                                  axis=0)
        # 改变列顺序，把参考人数一列放在班级后面
        combined_data = self.change_columns_order(combined_data)

        return combined_data

    @staticmethod
    def get_unmatched_data(data, total_col):
        """
        计算错位人数
        :param total_col: 总分或总分赋分
        :param data: 双有效数据
        :return: 返回一个错位人数的df

        """
        unmatched_list = []
        for subject in data.columns:
            unmatched_subject = data[total_col] - data[subject]
            unmatched_list.append(unmatched_subject)

        unmatched_df_dict = dict(zip(data.columns, unmatched_list))
        unmatched_df = pd.DataFrame(unmatched_df_dict, index=data.index)
        unmatched_df.loc['年级共计'] = [unmatched_df[col].sum() for col in unmatched_df.columns]
        del unmatched_df[total_col]
        return unmatched_df

    @staticmethod
    def get_subject_max_min_score(data, subject):
        """
        用于计算获取各等级的卷面分区间（Y1-Y2)
        :param data:
        :param subject:
        :return: 各等级的卷面分区间（Y1-Y2)
        """
        max_score = []
        min_score = []
        subject_num = data[subject].count()
        A_num = int(subject_num * 0.15)
        B_num = int(subject_num * 0.35)
        C_num = int(subject_num * 0.35)
        D_num = int(subject_num * 0.13)
        data.sort_values(by=subject, ascending=False, inplace=True, ignore_index=True)
        # 计算A等级的卷面上下限分值及学生人数
        data_subject_A = data.loc[:A_num - 1, subject]
        A_max = data_subject_A.max()
        A_min = data_subject_A.min()
        max_score.append(A_max)
        min_score.append(A_min)
        final_data_subject_A = data[data[subject] >= A_min]
        # 计算B等级的卷面上下限分值及学生人数
        data_subject_B = data.loc[len(final_data_subject_A):(len(final_data_subject_A) + B_num - 1), subject]
        B_max = data_subject_B.max()
        B_min = data_subject_B.min()
        max_score.append(B_max)
        min_score.append(B_min)
        final_data_subject_B = data[(data[subject] >= B_min) & (data[subject] <= B_max)]
        # 计算C等级的卷面上下限分值及学生人数
        data_subject_C = data.loc[(len(final_data_subject_A) + len(final_data_subject_B)):(
                len(final_data_subject_A) + len(final_data_subject_B) + C_num - 1), subject]
        C_max = data_subject_C.max()
        C_min = data_subject_C.min()
        max_score.append(C_max)
        min_score.append(C_min)
        final_data_subject_C = data[(data[subject] >= C_min) & (data[subject] <= C_max)]
        # 计算D等级的卷面上下限分值及学生人数
        data_subject_D = data.loc[(len(final_data_subject_A) + len(final_data_subject_B) + len(final_data_subject_C)):(
                len(final_data_subject_A) + len(final_data_subject_B) + len(final_data_subject_C) + D_num - 1), subject]
        d_max = data_subject_D.max()
        d_min = data_subject_D.min()
        max_score.append(d_max)
        min_score.append(d_min)
        final_data_subject_D = data[(data[subject] >= d_min) & (data[subject] <= d_max)]
        # 计算E等级的卷面上下限分值及学生人数
        final_data_subject_E = data.loc[
                               (len(final_data_subject_A) + len(final_data_subject_B) + len(final_data_subject_C) + len(
                                   final_data_subject_D)):subject_num - 1, subject]
        E_max = final_data_subject_E.max()
        E_min = final_data_subject_E.min()
        max_score.append(E_max)
        min_score.append(E_min)
        print(
            f'{subject}学科A等级人数为：{len(final_data_subject_A)},'
            f'B等级人数为：{len(final_data_subject_B)},'
            f'C等级人数为：{len(final_data_subject_C)},'
            f'D等级人数为：{len(final_data_subject_D)},'
            f'E等级人数为：{len(final_data_subject_E)}')
        # 新版pandas2.0以上，接收到的数字在列表中会显示成np.float64(81.0)，所以要用下标读取出来才能正常显示小数
        print(
            f'{subject}学科等级卷面Y1的值分别为：{min_score[0]}, {min_score[1]}, {min_score[2]}, {min_score[3]}, {min_score[4]}, '
            f'Y2的值分别为：{max_score[0]}, {max_score[1]}, {max_score[2]}, {max_score[3]}, {max_score[4]}')
        return max_score, min_score

    @staticmethod
    def get_single_subject_data(data, subject, subject_score):
        '''
        计算获取一个学科单有效的df
        :param data: df
        :param subject: 学科
        :param subject_score: 一个学科的有效分数
        :return: 返回一个单有效的series
        '''
        single_subject = data[data[subject] >= subject_score].groupby(['班级'])[subject].count()
        return single_subject

    @staticmethod
    def get_double_subject_data(data, subject, total_col, subject_score, total):
        '''
        计算获取学科双有校的df
        :param data: df
        :param subject: 学科
        :param total_col: 总分字段名
        :param subject_score: 学科有效分数
        :param total: 总分
        :return: 返回一个学科双有效的series
        '''
        double_subject_data = data[data[total_col] >= total]
        double_subject = double_subject_data[double_subject_data[subject] >= subject_score].groupby('班级')[
            subject].count()
        return double_subject

    @staticmethod
    def get_added_score(y, y1, y2, t1, t2):
        """
        计算赋分值的公式：高考赋分方法，其中y为原始卷面得分，t为赋值得分，
        t1和t2为所在等级赋值区间的下限和上限，y1和y2为卷面所在等级分数区间的下限和上限。
        :param y:
        :param y1:
        :param y2:
        :param t1:
        :param t2:
        :return: 一个学生的赋值得分（四舍五入取整）
        """
        t = symbols('t')
        if y == y1:
            scores_added = t1
            return scores_added
        else:
            scores_added = solve((t2 - t) / (t - t1) - (y2 - y) / (y - y1), t)
            return round(scores_added[0])

    @staticmethod
    def get_level(score, min_score):
        '''
        计算赋分学科的等级
        :param score: 学科的卷面得分
        :param min_score: 所在等级的最低分
        :return: 相应的等级
        '''
        if score >= min_score[0]:
            return 'A'
        elif score >= min_score[1]:
            return 'B'
        elif score >= min_score[2]:
            return 'C'
        elif score >= min_score[3]:
            return 'D'
        elif score < min_score[3]:
            return 'E'

    @staticmethod
    def change_columns_order(data):
        col_first = '参考人数'
        cols_left = [col for col in data.columns if col != col_first]
        new_columns_data = data[[col_first] + cols_left]
        return new_columns_data

    @staticmethod
    def get_subject_good_score(data, total_col, subject):
        '''
        计算学科有效分。按照总分上线率，计算学科有效分
        :param data: DataFrame
        :param subject: 学科
        :param total: 总分（分为高线和中线）
        :return: 学科的有效分（高线和低线）
        '''
        total_num = data.shape[0]
        good_total_num = data.loc[data[total_col] >= GaokaoData2025.total_line].shape[0]
        good_percent_ratio = good_total_num / total_num
        good_subject_ratio = int(data[subject].count() * good_percent_ratio)
        data.sort_values(by=subject, ascending=False, inplace=True, ignore_index=True)
        good_subject_score = data.loc[:good_subject_ratio - 1, subject].min()

        return good_subject_score

    @staticmethod
    def rename_columns(col_name, word):
        if word in col_name:
            return col_name.replace(word, '')
        else:
            return col_name

    @staticmethod
    def time_use(fn):
        def inner(*args, **kwargs):
            print(f'开始运行主函数{fn.__name__}'.center(106, '_'))
            start_time = time.time()
            result = fn(*args, **kwargs)
            end_time = time.time()
            time_used = end_time - start_time
            print(f'主函数{fn.__name__}结束运行'.center(106, '_'))
            print(f'运行主函数{fn.__name__}共计耗时{round(time_used, 2)}秒')
            return result

        return inner

    @time_use
    def show_menu(self):
        print(self.__str__())

        while True:
            flag = eval(input(f'按键功能选择:\n '
                              f'    1:年级考试成绩分析；\n '
                              f'    2:区级以上考试成绩分析；\n'
                              f'     3:按其它数字键退出程序.\n请选择：'))
            if flag == 1:
                GaokaoData2025.total_line = int(input('请输入划线分数(高线或中线): '))
                self.excel_school_files()
                print('成绩分析已完成，谢谢使用！')
                break
            elif flag == 2:
                self.excel_files()
                print('成绩分析已完成，谢谢使用！')
                break

            else:
                break


if __name__ == '__main__':
    file_path = r'D:\data_test\高2026级学生10月考成绩汇总.xlsx'
    # file_path = r'D:\data_test\高2022级零诊成绩测试数据.xlsx'

    # 不分科的各科有效分
    GaokaoData2025.subjects_good_scores_all = {'语文': 85, '数学': 74, '英语': 68, '物理': 31, '历史': 46, '政治': 41,
                                               '地理': 39, '化学': 42, '生物': 51, '总分': 370}
    # 物理类各科有效分
    GaokaoData2025.subjects_good_scores_physics = {'语文': 84, '数学': 64, '英语': 63, '物理': 25, '政治': 61,
                                                   '地理': 61, '化学': 56, '生物': 57, '总分': 370}
    # 历史类各科有效分
    GaokaoData2025.subjects_good_scores_history = {'语文': 87, '数学': 45, '英语': 62, '历史': 46, '政治': 62,
                                                   '地理': 63, '化学': None, '生物': 52, '总分': 370}
    # 划线分数（中线或高线）
    GaokaoData2025.total_line = 400

    newgaokao = GaokaoData2025(file_path)

    # newgaokao.excel_files()
    newgaokao.excel_school_files()
    # newgaokao.get_mixed_data()
    # newgaokao.get_data_processed()
    # newgaokao.get_average_school()
    # newgaokao.get_average()
