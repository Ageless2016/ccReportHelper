# -*- coding: utf-8 -*-
'''
@Time    : 2019/12/3 15:16
@Author  : CC
@File    : cprevents.py
'''

import xlwings as xw
import os
import time
import datetime


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

    records_2018 =  extract_2018_records(fdpath_2018)
    records_2019 = extract_2019_records(fdpath_2019)

    set_records_2018 = set(records_2018)
    set_records_2019 = set(records_2019)

    obj_new2018 = list(set_records_2018 - set_records_2019)
    obj_new2019 = list(set_records_2019 - set_records_2018)
    lst_new2018 = []
    lst_new2019 = []
    for obj_event in obj_new2018:
        lst_new2018.append(obj_event.content)

    for obj_event in obj_new2019:
        lst_new2019.append(obj_event.content)

    export_result(lst_new2018,lst_new2019)



# 提取2018文件夹下的报表，并返回EventRecord对象列表
def extract_2018_records(fdpath):
    record_list = []
    sht_voice_name = '普通语音2G'
    sht_volte_name = 'VoLTE指标汇总'
    sht_dropdetail_name = '电信掉话详情'
    sht_blockdetail_name = '电信未接通详情'
    filenames = os.listdir(fdpath)
    for filename in filenames:
        if filename[-5:] != '.xlsx':
            continue
        filepathname = os.path.join(fdpath, filename)
        print("Extracting file : {}".format(filename))
        try:
            app = xw.App(visible=False, add_book=False)
            app.display_alerts = False
            app.screen_updating = False

            wb = app.books.open(filepathname)
            shts = wb.sheets
            servicetype = ''
            abnormaltype = ''

            for sht in shts:
                if sht.name == sht_voice_name:
                    servicetype = 'voice'
                    # 遍历2g语音指标表,查找掉话和未接通>0的记录
                    sht_max_row = sht.cells(1048576, 'D').end('up').row
                    drop_sht_max_row = shts[sht_dropdetail_name].cells(1048576, 'C').end('up').row
                    block_sht_max_row = shts[sht_blockdetail_name].cells(1048576, 'C').end('up').row
                    sht_content = sht.range(sht.cells(1,1),sht.cells(sht_max_row,42)).value
                    drop_event_content = shts[sht_dropdetail_name].range(shts[sht_dropdetail_name].cells(1,1),shts[sht_dropdetail_name].cells(drop_sht_max_row,4)).value
                    block_event_content = shts[sht_blockdetail_name].range(shts[sht_blockdetail_name].cells(1,1),shts[sht_blockdetail_name].cells(block_sht_max_row,4)).value
                    for i in range(5, sht_max_row):
                        blockedcount = sht_content[i][32]
                        droppedcount = sht_content[i][41]

                        # 查找电信未接通详情表找到主叫文件名、主叫起呼时间
                        if not blockedcount is None and blockedcount > 0:
                            city = sht_content[i][2]
                            testpoint_name = sht_content[i][3]
                            testscene = sht_content[i][4]
                            totalcount = sht_content[i][19]
                            coveredcount = sht_content[i][18]
                            blocked_detail_list = []
                            for j in range(5, block_sht_max_row):
                                mofilename = block_event_content[j][2]
                                moattempttime = block_event_content[j][3]
                                if testpoint_name in str(mofilename) and testscene in str(mofilename):
                                    abnormaltype = '未接通'
                                    event_record = EventRecord(city,testpoint_name, testscene, servicetype, totalcount, coveredcount,
                                                       abnormaltype,moattempttime,mofilename)
                                    record_list.append(event_record)

                        # 查找电信掉话详情表找到主叫文件名、主叫起呼时间
                        if not droppedcount is None and droppedcount > 0:
                            city = sht_content[i][2]
                            testpoint_name = sht_content[i][3]
                            testscene = sht_content[i][4]
                            totalcount = sht_content[i][19]
                            coveredcount = sht_content[i][18]
                            dropped_detail_list = []
                            for k in range(5, drop_sht_max_row):
                                mofilename = drop_event_content[k][2]
                                moattempttime = drop_event_content[k][3]
                                if testpoint_name in str(mofilename) and testscene in str(mofilename):
                                    abnormaltype = '掉话'
                                    event_record = EventRecord(city,testpoint_name, testscene, servicetype, totalcount, coveredcount,
                                                       abnormaltype,moattempttime,mofilename)
                                    record_list.append(event_record)

                elif sht.name == sht_volte_name:
                    servicetype = 'volte'
                    # 遍历volte语音指标表,查找掉话和未接通>0的记录
                    sht_max_row = sht.cells(1048576, 'D').end('up').row
                    drop_sht_max_row = shts[sht_dropdetail_name].cells(1048576, 'F').end('up').row
                    block_sht_max_row = shts[sht_blockdetail_name].cells(1048576, 'F').end('up').row
                    sht_content = sht.range(sht.cells(1,1),sht.cells(sht_max_row,103)).value
                    drop_event_content = shts[sht_dropdetail_name].range(shts[sht_dropdetail_name].cells(1,1),shts[sht_dropdetail_name].cells(drop_sht_max_row,7)).value
                    block_event_content = shts[sht_blockdetail_name].range(shts[sht_blockdetail_name].cells(1,1),shts[sht_blockdetail_name].cells(block_sht_max_row,7)).value
                    for i in range(4, sht_max_row):
                        blockedcount = sht_content[i][86]
                        droppedcount = sht_content[i][102]

                        # 查找电信未接通详情表找到主叫文件名、主叫起呼时间
                        if  not blockedcount is None and blockedcount > 0:
                            city = sht_content[i][2]
                            testpoint_name = sht_content[i][3]
                            testscene = sht_content[i][4]
                            totalcount = sht_content[i][12]
                            coveredcount = sht_content[i][11]
                            blocked_detail_list = []
                            for j in range(5, block_sht_max_row):
                                mofilename = block_event_content[j][5]
                                moattempttime = block_event_content[j][6]
                                if testpoint_name in str(mofilename) and testscene in str(mofilename):
                                    abnormaltype = '未接通'
                                    event_record = EventRecord(city,testpoint_name, testscene, servicetype, totalcount, coveredcount,
                                                       abnormaltype,moattempttime,mofilename)
                                    record_list.append(event_record)

                        # 查找电信掉话详情表找到主叫文件名、主叫起呼时间
                        if  not droppedcount is None and  droppedcount > 0:
                            city = sht_content[i][2]
                            testpoint_name = sht_content[i][3]
                            testscene = sht_content[i][4]
                            totalcount = sht_content[i][12]
                            coveredcount = sht_content[i][11]
                            for k in range(5, drop_sht_max_row):
                                mofilename = drop_event_content[k][5]
                                moattempttime = drop_event_content[k][6]
                                if testpoint_name in str(mofilename) and testscene in str(mofilename):
                                    abnormaltype = '掉话'
                                    event_record = EventRecord(city,testpoint_name, testscene, servicetype, totalcount, coveredcount,
                                                       abnormaltype,moattempttime,mofilename)
                                    record_list.append(event_record)

            if servicetype == '': return
        except Exception as e:
            # raise
            print(e)
        finally:
            wb.close()
            app.quit()
    return record_list


# 提取2019文件夹下的报表，并返回EventRecord对象列表
def extract_2019_records(fdpath):
    record_list = []
    sht_voice_name = 'Voice指标'
    sht_volte_name = 'VoLTE指标'
    sht_volte_dropdetail_name = '电信掉话详情'
    sht_volte_blockdetail_name = '电信未接通详情'
    sht_voice_dropdetail_name = '电信Voice掉话详情'
    sht_voice_blockdetail_name = '电信Voice未接通详情'
    filenames = os.listdir(fdpath)
    for filename in filenames:
        if filename[-5:] != '.xlsx':
            continue
        filepathname = os.path.join(fdpath, filename)
        print("Extracting file : {}".format(filename))

        try:
            app = xw.App(visible=False,add_book=False)
            app.display_alerts = False
            app.screen_updating = False
            wb = app.books.open(filepathname)
            shts = wb.sheets
            abnormaltype = ''

            for sht in shts:
                if sht.name == sht_voice_name:
                    servicetype = 'voice'
                    sht_max_row = sht.cells(1048576, 'D').end('up').row
                    drop_sht_max_row = shts[sht_voice_dropdetail_name].cells(1048576, 'C').end('up').row
                    block_sht_max_row = shts[sht_voice_blockdetail_name].cells(1048576, 'C').end('up').row
                    sht_content = sht.range(sht.cells(1,1),sht.cells(sht_max_row,16)).value
                    drop_event_content = shts[sht_voice_dropdetail_name].range(shts[sht_voice_dropdetail_name].cells(1,1),shts[sht_voice_dropdetail_name].cells(drop_sht_max_row,4)).value
                    block_event_content = shts[sht_voice_blockdetail_name].range(shts[sht_voice_blockdetail_name].cells(1,1),shts[sht_voice_blockdetail_name].cells(block_sht_max_row,4)).value
                    for i in range(4,sht_max_row):
                        blockedcount = sht_content[i][13]
                        droppedcount = sht_content[i][15]

                        if not blockedcount is None and blockedcount > 0:
                            city = sht_content[i][2]
                            testpoint_name = sht_content[i][3]
                            testscene = sht_content[i][4]
                            totalcount = sht_content[i][9]
                            coveredcount = sht_content[i][8]
                            blocked_detail_list = []
                            for j in range(5, block_sht_max_row):
                                mofilename = block_event_content[j][2]
                                moattempttime = block_event_content[j][3]
                                if testpoint_name in str(mofilename) and testscene in str(mofilename):
                                    abnormaltype = '未接通'
                                    event_record = EventRecord(city,testpoint_name, testscene, servicetype, totalcount, coveredcount,
                                                       abnormaltype,moattempttime,mofilename)
                                    record_list.append(event_record)

                        if not droppedcount is None and droppedcount > 0:
                            city = sht_content[i][2]
                            testpoint_name = sht_content[i][3]
                            testscene = sht_content[i][4]
                            totalcount = sht_content[i][9]
                            coveredcount = sht_content[i][8]
                            dropped_detail_list = []
                            for k in range(5, drop_sht_max_row):
                                mofilename = drop_event_content[k][2]
                                moattempttime = drop_event_content[k][3]
                                if testpoint_name in str(mofilename) and testscene in str(mofilename):
                                    abnormaltype = '掉话'
                                    event_record = EventRecord(city,testpoint_name, testscene, servicetype, totalcount, coveredcount,
                                                       abnormaltype,moattempttime,mofilename)
                                    record_list.append(event_record)

                elif sht.name == sht_volte_name:
                    servicetype = 'volte'
                    # 遍历volte语音指标表,查找掉话和未接通>0的记录
                    sht_max_row = sht.cells(1048576, 'D').end('up').row
                    drop_sht_max_row = shts[sht_volte_dropdetail_name].cells(1048576, 'C').end('up').row
                    block_sht_max_row = shts[sht_volte_blockdetail_name].cells(1048576, 'C').end('up').row
                    sht_content = sht.range(sht.cells(1,1),sht.cells(sht_max_row,103)).value
                    drop_event_content = shts[sht_volte_dropdetail_name].range(shts[sht_volte_dropdetail_name].cells(1,1),shts[sht_volte_dropdetail_name].cells(drop_sht_max_row,5)).value
                    block_event_content = shts[sht_volte_blockdetail_name].range(shts[sht_volte_blockdetail_name].cells(1,1),shts[sht_volte_blockdetail_name].cells(block_sht_max_row,5)).value
                    for i in range(3, sht_max_row):
                        blockedcount = sht_content[i][86]
                        droppedcount = sht_content[i][102]
                        # 查找电信未接通详情表找到主叫文件名、主叫起呼时间
                        if  not blockedcount is None and blockedcount > 0:
                            city = sht_content[i][2]
                            testpoint_name = sht_content[i][3]
                            testscene = sht_content[i][4]
                            totalcount = sht_content[i][12]
                            coveredcount = sht_content[i][11]
                            blocked_detail_list = []
                            for j in range(5, block_sht_max_row):
                                mofilename = block_event_content[j][2]
                                moattempttime = block_event_content[j][4]
                                if testpoint_name in str(mofilename) and testscene in str(mofilename):
                                    abnormaltype = '未接通'
                                    event_record = EventRecord(city,testpoint_name, testscene, servicetype, totalcount, coveredcount,
                                                       abnormaltype,moattempttime,mofilename)
                                    record_list.append(event_record)

                        # 查找电信掉话详情表找到主叫文件名、主叫起呼时间
                        if  not droppedcount is None and  droppedcount > 0:
                            city = sht_content[i][2]
                            testpoint_name = sht_content[i][3]
                            testscene = sht_content[i][4]
                            totalcount = sht_content[i][12]
                            coveredcount = sht_content[i][11]
                            dropped_detail_list = []
                            for k in range(5, drop_sht_max_row):
                                mofilename = drop_event_content[k][2]
                                moattempttime = drop_event_content[k][4]
                                if testpoint_name in str(mofilename) and testscene in str(mofilename):
                                    abnormaltype = '掉话'
                                    event_record = EventRecord(city,testpoint_name, testscene, servicetype, totalcount, coveredcount,
                                                       abnormaltype,moattempttime,mofilename)
                                    record_list.append(event_record)

        except Exception as e:
            # raise
            print(e)
        finally:
            wb.close()
            app.quit()

    return record_list

# 进行详情记录对比
def export_result(newevents_2018,newevents_2019):
    fp,fn = os.path.split(os.path.abspath(__file__))
    timesn = int(time.time())
    result_fn = 'results_' + str(timesn) + '.xlsx'
    result_full_path = os.path.join(fp,result_fn)
    app = xw.App(visible=False,add_book=False)
    app.screen_updating = False
    app.display_alerts = False
    new_wb = app.books.add()
    sht_new_2018 =  new_wb.sheets.add('2018新增')
    sht_new_2019 = new_wb.sheets.add('2019新增')
    new_wb.sheets('sheet1').delete()

    xlheader = ['城市','测试点','场景','业务类型','采样点分母','采样点分子','异常类型','起呼时间','文件名']
    sht_new_2018.cells(1,1).value = xlheader
    sht_new_2018.cells(2, 1).value = newevents_2018
    sht_new_2019.cells(1,1).value = xlheader
    sht_new_2019.cells(2,1).value = newevents_2019

    app.screen_updating = True
    app.display_alerts = True
    new_wb.save(result_full_path)
    app.quit()
    print('Success! 文件路径: {}'.format(result_full_path))


class EventRecord():
    def __init__(self, city,testpoint_name, testscene, servicetype, totalcount, coveredcount, abnormaltype,
                 mocallattempttime,mofilename):
        self.city = city #城市
        self.testpoint_name = testpoint_name  # 测试点
        self.testscene = testscene  # 场景（深度，浅度）
        self.servicetype = servicetype  # 业务类型（voice,volte）
        self.totalcount = totalcount  # 采样点分母
        self.coveredcount = coveredcount  # 采样点分子
        self.abnormaltype = abnormaltype  # 异常类型（掉话、未接通）
        self.mocallattempttime = mocallattempttime #主叫起呼时间
        self.mofilename = mofilename #主叫文件名
        self.content = [self.city,self.testpoint_name,self.testscene,self.servicetype,self.totalcount,self.coveredcount,self.abnormaltype,self.mocallattempttime,self.mofilename]

    def __eq__(self, other):
        return self.city+self.testpoint_name+self.testscene+self.servicetype+self.abnormaltype + str(round(self.mocallattempttime,6)) == \
               other.city+other.testpoint_name+other.testscene+other.servicetype+other.abnormaltype + str(round(self.mocallattempttime,6))

    def __hash__(self):
        return hash(self.city+self.testpoint_name+self.testscene+self.servicetype+self.abnormaltype + str(round(self.mocallattempttime,6)) )

    def __repr__(self):
        fmt_mocallattempttime = floattimetostr(self.mocallattempttime)
        outstr = "city:{}\ntestpoint_name:{}\ntestscene:{}\nservicetype:{}\ntotalcount:{}\ncoveredcount:{}\nabnormaltype:{}\nfmt_mocallattempttime:{}\nmofilename:{}\n".format(
            self.city,self.testpoint_name,self.testscene,self.servicetype,self.totalcount,self.coveredcount,self.abnormaltype,fmt_mocallattempttime,self.mofilename)
        return outstr



def floattimetostr(floattime):
    stamp =round((floattime-25569)*1000000*86400)
    dateArray = datetime.datetime.utcfromtimestamp(int(str(stamp)[0:10]))
    microSec = round(int(str(stamp)[-6:])/1000)
    custom_time_format = str(dateArray.strftime("%Y-%m-%d %H:%M:%S")) + "." + str(microSec)
    return custom_time_format


if __name__ == '__main__':
    run()
