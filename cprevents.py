# -*- coding: utf-8 -*-
'''
@Time    : 2019/12/3 15:16
@Author  : CC
@File    : cprevents.py
'''

import xlwings as xw
import os

# 主执行函数
def run():
    current_path,filename = os.path.split(__file__)

    fdpath_2018 = os.path.join(current_path,'2018')
    fdpath_2019 = os.path.join(current_path,'2019')

    if  not os.path.exists(fdpath_2018):
        print("2018 folder not found!")
        return

    if not os.path.exists(fdpath_2019):
        print("2019 folder not found!")
        return

    extract_2018_records(fdpath_2018)
    extract_2019_records(fdpath_2019)




#提取2018文件夹下的报表，并返回EventRecord对象列表
def extract_2018_records(fdpath):
    record_list = []
    sht_voice_name = '普通语音2G'
    sht_volte_name = 'VoLTE语音报表'
    sht_dropdetail_name = '电信掉话详情'
    sht_blockdetail_name = '电信未接通详情'

    filenames = os.listdir(fdpath)
    for filename in filenames:
        if filename[-5:] != '.xlsx':
            continue
        filepathname = os.path.join(fdpath, filename)
        print("Extracting file : {}".format(filename))
        try:
            app = xw.App(visible=True,add_book=False)
            # app.display_alerts = False
            # app.screen_updating = False

            wb = app.books.open(filepathname)
            shts = wb.sheets
            testpoint_name = ''
            testscene = ''
            servicetype = ''
            totalcount = 0
            coveredcount = 0
            abnormaltype = ''
            abnormaldetaillist = []

            for sht in shts:
                if sht.name == sht_voice_name:
                    servicetype = 'voice'
                    print(servicetype)

                    #遍历2g语音指标表,查找掉话和未接通>0的记录


                elif sht.name == sht_volte_name:
                    servicetype = 'volte'
                    sht_kpi = sht
                    print(servicetype)

            if servicetype=='':return





        except Exception as e:
            print(e)


    return record_list



#提取2019文件夹下的报表，并返回EventRecord对象列表
def extract_2019_records(fdpath):
    record_list = []

    return record_list




















class EventRecord():
    def __init__(self,testpoint_name,testscene,servicetype,totalcount,coveredcount,abnormaltype,abnormaldetaillist):
        self.testpoint_name = testpoint_name    #测试点名称
        self.testscene = testscene   #测试场景类型（深度，浅度）
        self.servicetype = servicetype    #业务类型（voice,volte）
        self.totalcount = totalcount      #采样点分母
        self.coveredcount = coveredcount       #采样点分子
        self.abnormaltype = abnormaltype       #异常类型（掉话、未接通）
        self.abnormaldetaillist = abnormaldetaillist    #异常详情列表


    class AbnormalDetail():
        def __init__(self,mofilename,mocallattempttime):
            self.mofilename = mofilename
            self.mocallattempttime = mocallattempttime




if __name__ == '__main__':
    run()