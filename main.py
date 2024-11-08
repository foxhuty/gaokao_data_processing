# -*_ codeing=utf-8 -*-
# @Time: 2024/11/8 10:57
# @Author: foxhuty
# @File: main.py
# @Software: PyCharm
# @Based on python 3.13
from gaokao_data_process import GaokaoData2025
import sys
class MyException(object):
    def __init__(self, obj_data):
        self.obj_data = obj_data
        import traceback
        import logging
        logging.basicConfig(
            level=logging.DEBUG,
            filename='D:\\my_logging\\error.log',
            format='%(asctime)s %(levelname)s\n %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        logging.error(traceback.format_exc())
        logging.info(msg=self.obj_data)

def main(file):
    try:
        new_gaokao = GaokaoData2025(file)
        # new_gaokao.excel_files()
        new_gaokao.excel_school_files()
        MyException(new_gaokao)

    except Exception as e:
        print(f"An error occurred: {e}", file=sys.stderr)


if __name__ == '__main__':
    file_path = r'D:\data_test\高2026级学生10月考成绩汇总.xlsx'
    # file_path = r'D:\data_test\高2022级零诊成绩测试数据.xlsx'
    # file_path = r'D:\data_test\高2022级零诊成绩测试数据11.xlsx'
    # file_path = r'D:\data_test\一诊考试成绩分析统计（中线）.xlsx'

    # 不分科的各科有效分
    GaokaoData2025.subjects_good_scores_all = {'语文': 80, '数学': 80, '英语': 80, '物理': 40, '历史': 50, '政治': 40,
                                               '地理': 40, '化学': 50, '生物': 40, '总分': 400}
    # 物理类各科有效分
    GaokaoData2025.subjects_good_scores_physics = {'语文': 84, '数学': 64, '英语': 63, '物理': 25, '政治': 61,
                                                   '地理': 61, '化学': 56, '生物': 57, '总分': 370}
    # 历史类各科有效分
    GaokaoData2025.subjects_good_scores_history = {'语文': 87, '数学': 45, '英语': 62, '历史': 46, '政治': 62,
                                                   '地理': 63, '化学': None, '生物': 52, '总分': 370}
    # 高线
    GaokaoData2025.high_line = 460
    # 中线
    GaokaoData2025.mid_line = 370

    main(file_path)


