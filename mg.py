import os
import xlwings as xw
from sheetconfig import *

dict_fill = {}
header = []
def start(folder,fn):
    #初始化报表模板
    try:
        wb = xw.Book(fn)
        xw.App.visible = False
        shts = wb.sheets
        data_sht = wb.sheets[0]
        config_sht = wb.sheets['CONFIG']

        #===============获取指标表初始行、列，结束行、列=====================================
        datasht_startrow = int(config_sht.range("G1").value)
        if not isInt(datasht_startrow):
            return
        datasht_startcolumn = config_sht.range("G2").value
        if not isInt(datasht_startcolumn):
            return
        datasht_endrow = data_sht.cells(1048576, int(datasht_startcolumn)).end('up').row
        data_endcolumn = data_sht.cells(1, 16384).end('left').column
        #==============================================================================

        header = data_sht.range(data_sht.cells(1, 1), data_sht.cells(1, data_endcolumn)).value
        config_range_list = config_sht.cells(1, 1).current_region.value
        config_list,key_list = init_config(config_range_list, header)

        # 初始化用于填充的内容字典，字典包含一个多个关键字组成的key元组，和一个内容字典============
        if datasht_endrow > datasht_startrow:
            datasht_content = data_sht.range(data_sht.cells(datasht_startrow,1),data_sht.cells(datasht_endrow,data_endcolumn)).value

            for row in range(len(datasht_content)):
                row_data = {}
                key_v = []
                for r_v in key_list:
                    key_v.append(datasht_content[row][r_v])
                key_t = tuple(key_v)
                for r_d in range(data_endcolumn):
                    row_data[r_d] = datasht_content[row][r_d]
                dict_fill[key_t]=row_data
        elif datasht_endrow == datasht_startrow:
            datasht_content = data_sht.range(data_sht.cells(datasht_startrow, 1),data_sht.cells(datasht_endrow, data_endcolumn)).value
            for row in range(len(datasht_content)):
                row_data={}
                key_v = []
                for r_v in key_list:
                    key_v.append(datasht_content[r_v])
                key_t = tuple(key_v)
                for r_d in range(data_endcolumn):
                    row_data[r_d]=datasht_content[r_d]
                dict_fill[key_t]=row_data


        filenames = os.listdir(folder)
        for filename in filenames:
            if filename[-5:] != '.xlsx':
                continue
            filepathname = os.path.join(folder, filename)
            print("Processing files:{}".format(filename))
            temp_wb = xw.Book(filepathname)
            temp_shts = temp_wb.sheets
            for cfg in config_list:
                for temp_sht in temp_shts:
                    if temp_sht.name == cfg.sheet_name:
                        temp_sht_max_row = temp_sht.cells(1048576,cfg.start_column).end('up').row
                        sht_content = temp_sht.used_range.value
                        print("insert row data from {}...".format(cfg.sheet_name))
                        insert_data(cfg,sht_content,temp_sht_max_row,data_endcolumn)

            filling_list = []
            for v1 in dict_fill.values():
                filling_row = []
                for v2 in v1.values():
                    filling_row.append(v2)
                filling_list.append(filling_row)
            temp_wb.close()

            data_sht.cells(3,1).value = filling_list

        print("Done!")
        xw.App.visible = True


    except Exception as e:
        print("Error:{}".format(e))
        xw.App.visible = True
        return


def insert_data(cfg,sht_content,sht_max_row,data_endcolumn):
    for i in range(cfg.start_row,sht_max_row):
        key_list=[]
        for key_column in cfg.key_columns:
            key_list.append(sht_content[i][key_column])

        key_tuple = tuple(key_list)
        if key_tuple in dict_fill.keys():

            dict_row_data = dict_fill[key_tuple]
            for k in range(len(cfg.self_columns)):
                dict_row_data[cfg.self_columns[k]] = sht_content[i][cfg.data_columns[k]]

        else:
            dict_blank_data = NewBlankRow(data_endcolumn).value
            dict_fill[key_tuple] = dict_blank_data
            for m in range(len(cfg.key_self_columns)):
                dict_blank_data[cfg.key_self_columns[m]] = sht_content[i][cfg.key_columns[m]]
            for k in range(len(cfg.self_columns)):
                dict_blank_data[cfg.self_columns[k]]=sht_content[i][cfg.data_columns[k]]




#初始化配置列表
def init_config(config_range_list,header):
    print("Initial configration...")
    obj_config_list = []
    for i in range(1,len(config_range_list)):
        tmp_row = config_range_list[i]
        obj_config = SheetConfig(tmp_row[0],tmp_row[1],tmp_row[2],tmp_row[3])
        obj_config_list.append(obj_config)
    #初始化cfg对象data_columns、key_columns值
    for cfg in obj_config_list:
        key_list = []
        for i in range(0,len(header)):
            if header[i] is None or header[i]== '' or header[i].find('_') == -1:
                continue
            tmp_sht_code = header[i].split('_')[0]
            tmp_sht_column = header[i].split('_')[1]
            if tmp_sht_code == 'PK':
                cfg.key_columns.append(colname2index(tmp_sht_column))
                cfg.key_self_columns.append(i)
                key_list.append(i)
            elif tmp_sht_code == cfg.sheet_code:
                cfg.data_columns.append(colname2index(tmp_sht_column))
                cfg.self_columns.append(i)

    return obj_config_list,key_list




def colname2index(colname):
    index = -1
    num = 65

    for char in colname:
        index = (index + 1) * 26 + ord(char) - num

    return index

def isInt(num):
    try:
        num = int(num)
        return isinstance(num,int)
    except:
        return False





