# -*- coding: utf-8 -*-
'''
@Time    : 2019/11/12 13:19
@Author  : CC
@File    : lf.py
'''
import os
import csv
import codecs

#列出指定目录的文件名，并生成CSV文件

def lf(path):
    file_path_list = []
    files = os.listdir(path)
    for file in files:
        file_path = os.path.join(path,file)
        #如果是文件夹，递归调用
        if os.path.isdir(file_path):
            lf(file_path)
        elif os.path.isfile(file_path):
            file_path_list.append(file_path)

    #在可执行目录生成csv文件
    csvfile = os.path.join(os.getcwd(),'filenames.csv')
    write_csv(csvfile,file_path_list)



def write_csv(file_name,datas):
    #追加写
    file_csv = codecs.open(file_name,'w+','gbk')
    writer = csv.writer(file_csv)
    for data in datas:
        writer.writerow((str(data),))
    print('Success！')