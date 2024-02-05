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
    current_path = os.getcwd()
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
                            for j in range(5, block_sht_max_row):
                                mofilename = block_event_content[j][2]
                                if testpoint_name in str(mofilename) and testscene in str(mofilename):
                                    abnormaltype = '未接通'
                                    moattempttime = block_event_content[j][3]
                                    extend_detail = [None]*32
                                    extend_detail[0] = mofilename  #主叫文件名
                                    event_record = EventRecord(city,
                                                               testpoint_name,
                                                               testscene,
                                                               servicetype,
                                                               abnormaltype,
                                                               moattempttime,extend_detail)
                                    record_list.append(event_record)

                        # 查找电信掉话详情表找到主叫文件名、主叫起呼时间
                        if not droppedcount is None and droppedcount > 0:
                            city = sht_content[i][2]
                            testpoint_name = sht_content[i][3]
                            testscene = sht_content[i][4]
                            dropped_detail_list = []
                            for k in range(5, drop_sht_max_row):
                                mofilename = drop_event_content[k][2]
                                if testpoint_name in str(mofilename) and testscene in str(mofilename):
                                    abnormaltype = '掉话'
                                    moattempttime = drop_event_content[k][3]
                                    extend_detail = [None]*32
                                    extend_detail[0] = mofilename
                                    event_record = EventRecord(city,
                                                               testpoint_name,
                                                               testscene,
                                                               servicetype,
                                                               abnormaltype,
                                                               moattempttime,extend_detail)
                                    record_list.append(event_record)

                elif sht.name == sht_volte_name:
                    servicetype = 'volte'
                    # 遍历volte语音指标表,查找掉话和未接通>0的记录
                    sht_max_row = sht.cells(1048576, 'D').end('up').row
                    drop_sht_max_row = shts[sht_dropdetail_name].cells(1048576, 'F').end('up').row
                    block_sht_max_row = shts[sht_blockdetail_name].cells(1048576, 'F').end('up').row
                    sht_content = sht.range(sht.cells(1,1),sht.cells(sht_max_row,103)).value
                    drop_event_content = shts[sht_dropdetail_name].range(shts[sht_dropdetail_name].cells(1,1),shts[sht_dropdetail_name].cells(drop_sht_max_row,57)).value
                    block_event_content = shts[sht_blockdetail_name].range(shts[sht_blockdetail_name].cells(1,1),shts[sht_blockdetail_name].cells(block_sht_max_row,57)).value
                    for i in range(4, sht_max_row):
                        blockedcount = sht_content[i][86]
                        droppedcount = sht_content[i][102]

                        # 查找电信未接通详情表找到主叫文件名、主叫起呼时间
                        if  not blockedcount is None and blockedcount > 0:
                            city = sht_content[i][2]
                            testpoint_name = sht_content[i][3]
                            testscene = sht_content[i][4]
                            for j in range(5, block_sht_max_row):
                                mofilename = block_event_content[j][5]
                                if testpoint_name in str(mofilename) and testscene in str(mofilename):
                                    abnormaltype = '未接通'
                                    moattempttime = block_event_content[j][6]
                                    extend_detail = [None]*32
                                    extend_detail[0] = mofilename #主叫文件名
                                    extend_detail[1] = block_event_content[j][7]  # 主叫振铃时间
                                    extend_detail[2] = block_event_content[j][8]  # 主叫无线接通时间
                                    extend_detail[3] = block_event_content[j][10]  # 主叫呼叫结束网络状态
                                    extend_detail[4] = block_event_content[j][11]  # 主叫呼叫类型
                                    extend_detail[5] = block_event_content[j][15]  # Bye_Request
                                    extend_detail[6] = block_event_content[j][16]  # InVite
                                    extend_detail[7] = block_event_content[j][17]  # Cancel
                                    extend_detail[8] = block_event_content[j][18]  # SIP_Bye
                                    extend_detail[9] = block_event_content[j][19]  # SIP_Message
                                    extend_detail[10] = block_event_content[j][22]  # Session End
                                    extend_detail[11] = block_event_content[j][27]  # RSRP
                                    extend_detail[12] = block_event_content[j][25]  # SINR
                                    extend_detail[13] = block_event_content[j][26]  # TxPower
                                    extend_detail[14] = block_event_content[j][28]  # PDSCH BLER
                                    extend_detail[15] = block_event_content[j][30]  # 被叫文件名
                                    extend_detail[16] = block_event_content[j][31]  # 被叫起呼时间
                                    extend_detail[17] = block_event_content[j][33]  # 被叫振铃时间
                                    extend_detail[18] = block_event_content[j][34]  # 被叫无线接通时间
                                    extend_detail[19] = block_event_content[j][36]  # 被叫呼叫结束网络状态
                                    extend_detail[20] = block_event_content[j][37]  # 被叫呼叫类型
                                    extend_detail[21] = block_event_content[j][41]  # Bye_Request
                                    extend_detail[22] = block_event_content[j][42]  # InVite
                                    extend_detail[23] = block_event_content[j][43]  # Cancel
                                    extend_detail[24] = block_event_content[j][44]  # SIP_Bye
                                    extend_detail[25] = block_event_content[j][45]  # SIP_Message
                                    extend_detail[26] = block_event_content[j][48]  # Session End
                                    extend_detail[27] = block_event_content[j][53]  # RSRP
                                    extend_detail[28] = block_event_content[j][51]  # SINR
                                    extend_detail[29] = block_event_content[j][52]  # TxPower
                                    extend_detail[30] = block_event_content[j][54]  # PDSCH BLER
                                    extend_detail[31] = block_event_content[j][55]  # 原因
                                    event_record = EventRecord(city,
                                                               testpoint_name,
                                                               testscene,
                                                               servicetype,
                                                               abnormaltype,
                                                               moattempttime,extend_detail)
                                    record_list.append(event_record)

                        # 查找电信掉话详情表找到主叫文件名、主叫起呼时间
                        if  not droppedcount is None and  droppedcount > 0:
                            city = sht_content[i][2]
                            testpoint_name = sht_content[i][3]
                            testscene = sht_content[i][4]
                            for k in range(5, drop_sht_max_row):
                                mofilename = drop_event_content[k][5]
                                if testpoint_name in str(mofilename) and testscene in str(mofilename):
                                    abnormaltype = '掉话'
                                    extend_detail = [None]*32
                                    moattempttime = drop_event_content[k][6]
                                    extend_detail[0] = mofilename
                                    extend_detail[1] = drop_event_content[k][7]  # 主叫振铃时间
                                    extend_detail[2] = drop_event_content[k][8]  # 主叫无线接通时间
                                    extend_detail[3] = drop_event_content[k][10]  # 主叫呼叫结束网络状态
                                    extend_detail[4] = drop_event_content[k][12]  # 主叫呼叫类型
                                    extend_detail[5] = drop_event_content[k][16]  # Bye_Request
                                    extend_detail[6] = drop_event_content[k][17]  # InVite
                                    extend_detail[7] = drop_event_content[k][18]  # Cancel
                                    extend_detail[8] = drop_event_content[k][19]  # SIP_Bye
                                    extend_detail[9] = drop_event_content[k][20]  # SIP_Message
                                    extend_detail[10] = drop_event_content[k][23]  # Session End
                                    extend_detail[11] = drop_event_content[k][28]  # RSRP
                                    extend_detail[12] = drop_event_content[k][26]  # SINR
                                    extend_detail[13] = drop_event_content[k][27]  # TxPower
                                    extend_detail[14] = drop_event_content[k][29]  # PDSCH BLER
                                    extend_detail[15] = drop_event_content[k][31]  # 被叫文件名
                                    extend_detail[16] = drop_event_content[k][32]  # 被叫起呼时间
                                    extend_detail[17] = drop_event_content[k][33]  # 被叫振铃时间
                                    extend_detail[18] = drop_event_content[k][34]  # 被叫无线接通时间
                                    extend_detail[19] = drop_event_content[k][36]  # 被叫呼叫结束网络状态
                                    extend_detail[20] = drop_event_content[k][38]  # 被叫呼叫类型
                                    extend_detail[21] = drop_event_content[k][42]  # Bye_Request
                                    extend_detail[22] = drop_event_content[k][43]  # InVite
                                    extend_detail[23] = drop_event_content[k][44]  # Cancel
                                    extend_detail[24] = drop_event_content[k][45]  # SIP_Bye
                                    extend_detail[25] = drop_event_content[k][46]  # SIP_Message
                                    extend_detail[26] = drop_event_content[k][49]  # Session End
                                    extend_detail[27] = drop_event_content[k][54]  # RSRP
                                    extend_detail[28] = drop_event_content[k][52]  # SINR
                                    extend_detail[29] = drop_event_content[k][53]  # TxPower
                                    extend_detail[30] = drop_event_content[k][55]  # PDSCH BLER
                                    extend_detail[31] = drop_event_content[k][56]  # 原因
                                    event_record = EventRecord(city,
                                                               testpoint_name,
                                                               testscene,
                                                               servicetype,
                                                               abnormaltype,
                                                               moattempttime,extend_detail)
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
                            blocked_detail_list = []
                            for j in range(5, block_sht_max_row):
                                mofilename = block_event_content[j][2]
                                if testpoint_name in str(mofilename) and testscene in str(mofilename):
                                    abnormaltype = '未接通'
                                    moattempttime = block_event_content[j][3]
                                    extend_detail = [None]*32
                                    extend_detail[0] = mofilename
                                    event_record = EventRecord(city,
                                                               testpoint_name,
                                                               testscene,
                                                               servicetype,
                                                               abnormaltype,
                                                               moattempttime,extend_detail)
                                    record_list.append(event_record)

                        if not droppedcount is None and droppedcount > 0:
                            city = sht_content[i][2]
                            testpoint_name = sht_content[i][3]
                            testscene = sht_content[i][4]
                            for k in range(5, drop_sht_max_row):
                                mofilename = drop_event_content[k][2]
                                if testpoint_name in str(mofilename) and testscene in str(mofilename):
                                    abnormaltype = '掉话'
                                    moattempttime = drop_event_content[k][3]
                                    extend_detail = [None]*32
                                    extend_detail[0] = mofilename
                                    event_record = EventRecord(city,
                                                               testpoint_name,
                                                               testscene,
                                                               servicetype,
                                                               abnormaltype,
                                                               moattempttime,extend_detail)
                                    record_list.append(event_record)

                elif sht.name == sht_volte_name:
                    servicetype = 'volte'
                    # 遍历volte语音指标表,查找掉话和未接通>0的记录
                    sht_max_row = sht.cells(1048576, 'D').end('up').row
                    drop_sht_max_row = shts[sht_volte_dropdetail_name].cells(1048576, 'C').end('up').row
                    block_sht_max_row = shts[sht_volte_blockdetail_name].cells(1048576, 'C').end('up').row
                    sht_content = sht.range(sht.cells(1,1),sht.cells(sht_max_row,103)).value
                    drop_event_content = shts[sht_volte_dropdetail_name].range(shts[sht_volte_dropdetail_name].cells(1,1),shts[sht_volte_dropdetail_name].cells(drop_sht_max_row,62)).value
                    block_event_content = shts[sht_volte_blockdetail_name].range(shts[sht_volte_blockdetail_name].cells(1,1),shts[sht_volte_blockdetail_name].cells(block_sht_max_row,62)).value
                    for i in range(3, sht_max_row):
                        blockedcount = sht_content[i][86]
                        droppedcount = sht_content[i][102]
                        # 查找电信未接通详情表找到主叫文件名、主叫起呼时间
                        if  not blockedcount is None and blockedcount > 0:
                            city = sht_content[i][2]
                            testpoint_name = sht_content[i][3]
                            testscene = sht_content[i][4]
                            for j in range(5, block_sht_max_row):
                                mofilename = block_event_content[j][2]
                                moattempttime = block_event_content[j][4]
                                if testpoint_name in str(mofilename) and testscene in str(mofilename):
                                    abnormaltype = '未接通'
                                    extend_detail = [None]*32
                                    extend_detail[0] = mofilename
                                    extend_detail[1] = block_event_content[j][5]  # 主叫振铃时间
                                    extend_detail[2] = block_event_content[j][6]  # 主叫无线接通时间
                                    extend_detail[3] = block_event_content[j][12]  # 主叫呼叫结束网络状态
                                    extend_detail[4] = block_event_content[j][9]  # 主叫呼叫类型
                                    extend_detail[5] = block_event_content[j][18]  # Bye_Request
                                    extend_detail[6] = block_event_content[j][19]  # InVite
                                    extend_detail[7] = block_event_content[j][20]  # Cancel
                                    extend_detail[8] = block_event_content[j][21]  # SIP_Bye
                                    extend_detail[9] = block_event_content[j][22]  # SIP_Message
                                    extend_detail[10] = block_event_content[j][25]  # Session End
                                    extend_detail[11] = block_event_content[j][26]  # RSRP
                                    extend_detail[12] = block_event_content[j][27]  # SINR
                                    extend_detail[13] = block_event_content[j][28]  # TxPower
                                    extend_detail[14] = block_event_content[j][29]  # PDSCH BLER
                                    extend_detail[15] = block_event_content[j][31]  # 被叫文件名
                                    extend_detail[16] = block_event_content[j][32]  # 被叫起呼时间
                                    extend_detail[17] = block_event_content[j][33]  # 被叫振铃时间
                                    extend_detail[18] = block_event_content[j][34]  # 被叫无线接通时间
                                    extend_detail[19] = block_event_content[j][40]  # 被叫呼叫结束网络状态
                                    extend_detail[20] = block_event_content[j][37]  # 被叫呼叫类型
                                    extend_detail[21] = block_event_content[j][46]  # Bye_Request
                                    extend_detail[22] = block_event_content[j][47]  # InVite
                                    extend_detail[23] = block_event_content[j][48]  # Cancel
                                    extend_detail[24] = block_event_content[j][49]  # SIP_Bye
                                    extend_detail[25] = block_event_content[j][50]  # SIP_Message
                                    extend_detail[26] = block_event_content[j][53]  # Session End
                                    extend_detail[27] = block_event_content[j][54]  # RSRP
                                    extend_detail[28] = block_event_content[j][55]  # SINR
                                    extend_detail[29] = block_event_content[j][56]  # TxPower
                                    extend_detail[30] = block_event_content[j][57]  # PDSCH BLER
                                    extend_detail[31] = block_event_content[j][61]  # 原因
                                    event_record = EventRecord(city,
                                                               testpoint_name,
                                                               testscene,
                                                               servicetype,
                                                               abnormaltype,
                                                               moattempttime,extend_detail)
                                    record_list.append(event_record)

                        # 查找电信掉话详情表找到主叫文件名、主叫起呼时间
                        if  not droppedcount is None and  droppedcount > 0:
                            city = sht_content[i][2]
                            testpoint_name = sht_content[i][3]
                            testscene = sht_content[i][4]
                            for k in range(5, drop_sht_max_row):
                                mofilename = drop_event_content[k][2]
                                moattempttime = drop_event_content[k][4]
                                if testpoint_name in str(mofilename) and testscene in str(mofilename):
                                    abnormaltype = '掉话'
                                    extend_detail = [None]*32
                                    extend_detail[0] = mofilename
                                    extend_detail[1] = drop_event_content[k][5]  # 主叫振铃时间
                                    extend_detail[2] = drop_event_content[k][6]  # 主叫无线接通时间
                                    extend_detail[3] = drop_event_content[k][12]  # 主叫呼叫结束网络状态
                                    extend_detail[4] = drop_event_content[k][9]  # 主叫呼叫类型
                                    extend_detail[5] = drop_event_content[k][18]  # Bye_Request
                                    extend_detail[6] = drop_event_content[k][19]  # InVite
                                    extend_detail[7] = drop_event_content[k][20]  # Cancel
                                    extend_detail[8] = drop_event_content[k][21]  # SIP_Bye
                                    extend_detail[9] = drop_event_content[k][22]  # SIP_Message
                                    extend_detail[10] = drop_event_content[k][25]  # Session End
                                    extend_detail[11] = drop_event_content[k][26]  # RSRP
                                    extend_detail[12] = drop_event_content[k][27]  # SINR
                                    extend_detail[13] = drop_event_content[k][28]  # TxPower
                                    extend_detail[14] = drop_event_content[k][29]  # PDSCH BLER
                                    extend_detail[15] = drop_event_content[k][31]  # 被叫文件名
                                    extend_detail[16] = drop_event_content[k][32]  # 被叫起呼时间
                                    extend_detail[17] = drop_event_content[k][33]  # 被叫振铃时间
                                    extend_detail[18] = drop_event_content[k][34]  # 被叫无线接通时间
                                    extend_detail[19] = drop_event_content[k][40]  # 被叫呼叫结束网络状态
                                    extend_detail[20] = drop_event_content[k][37]  # 被叫呼叫类型
                                    extend_detail[21] = drop_event_content[k][46]  # Bye_Request
                                    extend_detail[22] = drop_event_content[k][47]  # InVite
                                    extend_detail[23] = drop_event_content[k][48]  # Cancel
                                    extend_detail[24] = drop_event_content[k][49]  # SIP_Bye
                                    extend_detail[25] = drop_event_content[k][50]  # SIP_Message
                                    extend_detail[26] = drop_event_content[k][53]  # Session End
                                    extend_detail[27] = drop_event_content[k][54]  # RSRP
                                    extend_detail[28] = drop_event_content[k][55]  # SINR
                                    extend_detail[29] = drop_event_content[k][56]  # TxPower
                                    extend_detail[30] = drop_event_content[k][57]  # PDSCH BLER
                                    extend_detail[31] = drop_event_content[k][61]  # 原因
                                    event_record = EventRecord(city,
                                                               testpoint_name,
                                                               testscene,
                                                               servicetype,
                                                               abnormaltype,
                                                               moattempttime,extend_detail)
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
    fp = os.getcwd()
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

    xlheader = ['城市',
                '测试点','场景','业务类型','异常类型',
                '主叫起呼时间','主叫文件名','主叫振铃时间','主叫无线接通时间','主叫呼叫结束网络状态','主叫呼叫类型',
                'Bye_Request','InVite','Cancel','SIP_Bye','SIP_Message','Session End',
                'RSRP','SINR', 'TxPower','PDSCH BLER',
                '被叫文件名','被叫起呼时间','被叫振铃时间','被叫无线接通时间','被叫呼叫结束网络状态','被叫呼叫类型',
                'Bye_Request', 'InVite', 'Cancel', 'SIP_Bye', 'SIP_Message', 'Session End',
                'RSRP','SINR', 'TxPower','PDSCH BLER',
                '原因']

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
    def __init__(self,
                 city,
                 testpoint_name,
                 testscene,
                 servicetype,
                 abnormaltype,
                 mocallattempttime,
                 extenddetail=[]):
        self.city = city #城市
        self.testpoint_name = testpoint_name  # 测试点
        self.testscene = testscene  # 场景（深度，浅度）
        self.servicetype = servicetype  # 业务类型（voice,volte）
        self.abnormaltype = abnormaltype  # 异常类型（掉话、未接通）
        self.mocallattempttime = mocallattempttime #主叫起呼时间
        self.extenddetail = extenddetail

    @property
    def content(self):
        val = [self.city,self.testpoint_name,self.testscene,self.servicetype,self.abnormaltype,self.mocallattempttime]
        val.extend(self.extenddetail)
        return val

    def __eq__(self, other):
        return self.city+self.testpoint_name+self.testscene+self.servicetype+self.abnormaltype + str(round(self.mocallattempttime,6)) == \
               other.city+other.testpoint_name+other.testscene+other.servicetype+other.abnormaltype + str(round(self.mocallattempttime,6))

    def __hash__(self):
        return hash(self.city+self.testpoint_name+self.testscene+self.servicetype+self.abnormaltype + str(round(self.mocallattempttime,6)) )

    def __repr__(self):
        fmt_mocallattempttime = floattimetostr(self.mocallattempttime)
        outstr = "city:{}\ntestpoint_name:{}\ntestscene:{}\nservicetype:{}\nabnormaltype:{}\nfmt_mocallattempttime:{}\n".format(
            self.city,self.testpoint_name,self.testscene,self.servicetype,self.abnormaltype,fmt_mocallattempttime)
        return outstr


def floattimetostr(floattime):
    stamp =round((floattime-25569)*1000000*86400)
    dateArray = datetime.datetime.utcfromtimestamp(int(str(stamp)[0:10]))
    milliSec = round(int(str(stamp)[-6:])/1000)
    custom_time_format = str(dateArray.strftime("%Y-%m-%d %H:%M:%S")) + "." + str(milliSec)
    return custom_time_format


if __name__ == '__main__':
    run()
