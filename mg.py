#coding=utf-8
import os
import xlwings as xw
import cls_header
import cls_config

def run(folder_path,template_path):

    xw.App.visible = False
    #加载template.xlsx文件
    tmp_wb = xw.Book(template_path)
    tmp_sheets = tmp_wb.sheets

    #获取配置sheet表
    try:
        tmp_config_sht = tmp_sheets['CONFIG']
        tmp_data_sht = tmp_sheets[0]
    except:
        print("No 'CONFIG' sheet found!")
        tmp_wb.close()
        return

    print("Initial configuration...")
    config_content = tmp_config_sht.cells(1,1).current_region.value
    user_config = []

    #获取用户设置列表
    for i in range(1,len(config_content)):
        cfg = cls_config.shtconfig(int(config_content[i][0]),config_content[i][1],
                                   config_content[i][2],int(config_content[i][3]),int(config_content[i][4]))
        user_config.append(cfg)

    end_column = tmp_data_sht.cells(1, 16384).end('left').column
    #获取EXCEL数据表中第一行解析表头，存到数组head_list
    header_list = tmp_data_sht.range(tmp_data_sht.cells(1,1),tmp_data_sht.cells(1,end_column)).value
    headers = []
    key_headers = []
    index = 0
    pk_count = 0
    for t_header in header_list:
        index = index + 1
        if t_header is None or t_header.find('_')== -1:
            continue
        ispk = str(t_header.split('_')[0]).upper() == 'PK'
        wb_columnname = t_header.split('_')[1].upper()
        if ispk:
            pk_count = pk_count + 1
            wb_sheetname = ''
            tmp_header = cls_header.header(index,wb_sheetname,wb_columnname,ispk)
            # headers.append(tmp_header)
            key_headers.append(tmp_header)
            # print(headers)
            continue
        wb_sheetcode = t_header.split('_')[0].upper()
        for cfg in user_config:
            if cfg.sheet_code == wb_sheetcode:
                wb_sheetname = cfg.sheet_name
                tmp_header = cls_header.header(index,wb_sheetname,wb_columnname,ispk)
                headers.append(tmp_header)

    #如果在模板中没找到PK字段，直接返回
    if pk_count == 0:
        print("在template表中未找到PK字段！")
        return

    #开始合并
    processMerge(tmp_data_sht,folder_path,user_config,headers,key_headers)

#遍历文件夹目录，读取EXCEL文件数据，合并到sheet1
def processMerge(tmp_data_sht,folder_path,user_config,headers,key_headers):
    print("Processing...")
    filenames = os.listdir(folder_path)
    for filename in filenames:
        if filename[-5:] != '.xlsx':
            continue
        filepathname = os.path.join(folder_path,filename)
        # print(filepathname)

        try:
            wb = xw.Book(filepathname)
        except Exception as e :
            print("Error:{}".format(e))
            return

        #提取用户配置数据
        for cfg in user_config:
            try:
                sht = wb.sheets[cfg.sheet_name]
                start_row = cfg.start_row
                start_column = cfg.start_column
                end_row = sht.cells(1048576,start_column).end('up').row
                end_column = sht.cells(start_row-1,16384).end('left').column
                #end_row < start_row 说明是空表
                if end_row < start_row:
                    continue
                sht_content = sht.range(sht.cells(start_row,start_column),sht.cells(end_row,end_column)).value
                tmp_data_content = tmp_data_sht.used_range
                for i in range(0,end_row-start_row+1):
                    keystring = ''
                    for key_header in key_headers:
                        colindex = colname2index(key_header.wb_columnname)-start_column+1
                        keystring = keystring + str(sht_content[i][colindex])

                    print(keystring)
            except Exception as e:
                print("Error:{}".format(e))
                continue


#EXCEL 列名转列号，列号从0开始代表A列
def colname2index(colname):
    index = -1
    num = 65

    for char in colname:
        index = (index + 1) * 26 + ord(char) - num

    return index








    xw.App.visible = True








