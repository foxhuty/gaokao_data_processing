# -*- coding: utf-8 -*-
# @Time: 2024/11/10 14:35
# @Author: foxhuty
# @File: gaokao_logging.py

class MyException:
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
