# -*- coding: utf-8 -*-
'''
@Time    : 2019/11/12 14:00
@Author  : CC
@File    : rn.py
'''
import os

def rn(file):
    f = open(file,encoding='utf-8')
    data = f.readlines()
    for item in data:
        try:
            src_path_file,tar_file = tuple(str(item).replace('\n','').split(','))
            path,src_file = os.path.split(src_path_file)
            tar_path_file = os.path.join(path,tar_file)
            os.rename(src_path_file,tar_path_file)
            print("{} -> {} success!".format(src_file,tar_file))
        except Exception as e:
            print(e)
        finally:
            f.close()
