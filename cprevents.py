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
    current_path, filename = os.path.split(__file__)

    fdpath_2018 = os.path.join(current_path, '2018')
    fdpath_2019 = os.path.join(current_path, '2019')

    if not os.path.exists(fdpath_2018):
        print("2018 folder not found!")
        return

    if not os.path.exists(fdpath_2019):
        print("2019 folder not found!")
        return

    extract_2018_records(fdpath_2018)
    extract_2019_records(fdpath_2019)


# 提取2018文件夹下的报表，并返回EventRecord对象列表
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
            app = xw.App(visible=True, add_book=False)
            # app.display_alerts = False
            # app.screen_updating = False

            wb = app.books.open(filepathname)
            shts = wb.sheets
            city = ''
            testpoint_name = ''
            testscene = ''
            servicetype = ''
            totalcount = 0
            coveredcount = 0
            blockedcount = 0
            droppedcount = 0
            abnormaltype = ''
            abnormaldetaillist = []

            for sht in shts:
                if sht.name == sht_voice_name:
                    servicetype = 'voice'
                    # 遍历2g语音指标表,查找掉话和未接通>0的记录
                    sht_max_row = sht.cells(1048576, 'D').end('up').row
                    drop_sht_max_row = shts[sht_dropdetail_name].cells(1048576, 'C').end('up').row
                    block_sht_max_row = shts[sht_blockdetail_name].cells(1048576, 'C').end('up').row
                    sht_content = sht.used_range.value
                    drop_event_content = shts[sht_dropdetail_name].range(shts[sht_dropdetail_name].cells(1,1),shts[sht_dropdetail_name].cells(drop_sht_max_row,26)).value
                    block_event_content = shts[sht_blockdetail_name].range(shts[sht_blockdetail_name].cells(1,1),shts[sht_blockdetail_name].cells(block_sht_max_row,25)).value
                    for i in range(5, sht_max_row):
                        blockedcount = sht_content[i][32]
                        droppedcount = sht_content[i][41]

                        # 查找电信未接通详情表找到主叫文件名、主叫起呼时间
                        if blockedcount > 0:
                            city = sht_content[i][2]
                            testpoint_name = sht_content[i][3]
                            testscene = sht_content[i][4]
                            totalcount = sht_content[i][19]
                            coveredcount = sht_content[i][18]
                            blocked_detail_list = []
                            for j in range(5, block_sht_max_row):
                                mofilename = block_event_content[j][2]
                                moattempttime = block_event_content[j][3]
                                if testpoint_name in mofilename and testscene in mofilename:
                                    abnormaltype = '未接通'
                                    abnormal_event = EventRecord.AbnormalDetail(mofilename, moattempttime)
                                    blocked_detail_list.append(abnormal_event)
                                    if len(blocked_detail_list) >= blockedcount:break
                            event_record = EventRecord(testpoint_name, testscene, servicetype, totalcount, coveredcount,
                                                       abnormaltype, blocked_detail_list)
                            record_list.append(event_record)

                        # 查找电信掉话详情表找到主叫文件名、主叫起呼时间
                        if droppedcount > 0:
                            city = sht_content[i][2]
                            testpoint_name = sht_content[i][3]
                            testscene = sht_content[i][4]
                            totalcount = sht_content[i][19]
                            coveredcount = sht_content[i][18]
                            dropped_detail_list = []
                            for k in range(5, drop_sht_max_row):
                                mofilename = drop_event_content[k][2]
                                moattempttime = drop_event_content[k][3]
                                if testpoint_name in mofilename and testscene in mofilename:
                                    abnormaltype = '掉话'
                                    abnormal_event = EventRecord.AbnormalDetail(mofilename, moattempttime)
                                    dropped_detail_list.append(abnormal_event)
                                    if len(dropped_detail_list) >= droppedcount:break
                            event_record = EventRecord(testpoint_name, testscene, servicetype, totalcount, coveredcount,
                                                       abnormaltype, dropped_detail_list)
                            record_list.append(event_record)

                elif sht.name == sht_volte_name:
                    servicetype = 'volte'
                    # 遍历volte语音指标表,查找掉话和未接通>0的记录
                    sht_max_row = sht.cells(1048576, 'D').end('up').row
                    drop_sht_max_row = shts[sht_dropdetail_name].cells(1048576, 'C').end('up').row
                    block_sht_max_row = shts[sht_blockdetail_name].cells(1048576, 'C').end('up').row
                    sht_content = sht.used_range.value
                    drop_event_content = shts[sht_dropdetail_name].used_range.value
                    block_event_content = shts[sht_blockdetail_name].range(shts[sht_blockdetail_name].cells(1,1),shts[sht_blockdetail_name].cells(block_sht_max_row,25)).value
                    for i in range(5, sht_max_row):
                        blockedcount = sht_content[i][32]
                        droppedcount = sht_content[i][41]

                        # 查找电信未接通详情表找到主叫文件名、主叫起呼时间
                        if blockedcount > 0:
                            city = sht_content[i][2]
                            testpoint_name = sht_content[i][3]
                            testscene = sht_content[i][4]
                            totalcount = sht_content[i][19]
                            coveredcount = sht_content[i][18]
                            blocked_detail_list = []
                            for j in range(5, block_sht_max_row):
                                mofilename = block_event_content[j][2]
                                moattempttime = block_event_content[j][3]
                                if testpoint_name in mofilename and testscene in mofilename:
                                    abnormaltype = '未接通'
                                    abnormal_event = EventRecord.AbnormalDetail(mofilename, moattempttime)
                                    blocked_detail_list.append(abnormal_event)
                                    if len(blocked_detail_list) >= blockedcount:break
                            event_record = EventRecord(testpoint_name, testscene, servicetype, totalcount, coveredcount,
                                                       abnormaltype, blocked_detail_list)
                            record_list.append(event_record)

                        # 查找电信掉话详情表找到主叫文件名、主叫起呼时间
                        if droppedcount > 0:
                            city = sht_content[i][2]
                            testpoint_name = sht_content[i][3]
                            testscene = sht_content[i][4]
                            totalcount = sht_content[i][19]
                            coveredcount = sht_content[i][18]
                            dropped_detail_list = []
                            for k in range(5, drop_sht_max_row):
                                mofilename = drop_event_content[k][2]
                                moattempttime = drop_event_content[k][3]
                                if testpoint_name in mofilename and testscene in mofilename:
                                    abnormaltype = '掉话'
                                    abnormal_event = EventRecord.AbnormalDetail(mofilename, moattempttime)
                                    dropped_detail_list.append(abnormal_event)
                                    if len(dropped_detail_list) >= droppedcount:break
                            event_record = EventRecord(testpoint_name, testscene, servicetype, totalcount, coveredcount,
                                                       abnormaltype, dropped_detail_list)
                            record_list.append(event_record)
            if servicetype == '': return





        except Exception as e:
            raise
            print(e)

    return record_list


# 提取2019文件夹下的报表，并返回EventRecord对象列表
def extract_2019_records(fdpath):
    record_list = []

    return record_list


class EventRecord():
    def __init__(self, testpoint_name, testscene, servicetype, totalcount, coveredcount, abnormaltype,
                 abnormaldetaillist):
        self.testpoint_name = testpoint_name  # 测试点名称
        self.testscene = testscene  # 测试场景类型（深度，浅度）
        self.servicetype = servicetype  # 业务类型（voice,volte）
        self.totalcount = totalcount  # 采样点分母
        self.coveredcount = coveredcount  # 采样点分子
        self.abnormaltype = abnormaltype  # 异常类型（掉话、未接通）
        self.abnormaldetaillist = abnormaldetaillist  # 异常详情列表

    def __repr__(self):
        outstr = "testpoint_name:{}\ntestscene:{}\nservicetype:{}\ntotalcount:{}\ncoveredcount:{}\nabnormaltype:{}\nabnormaldetaillist:{}\n".format(
            self.testpoint_name,self.testscene,self.servicetype,self.totalcount,self.coveredcount,self.abnormaltype,self.abnormaldetaillist)
        return outstr

    class AbnormalDetail():
        def __init__(self, mofilename, mocallattempttime):
            self.mofilename = mofilename
            self.mocallattempttime = mocallattempttime

        def __repr__(self):
            outstr = "MoCallAttemptTime:{}".format(self.mocallattempttime)
            return outstr


if __name__ == '__main__':
    run()
