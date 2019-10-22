#coding=utf-8

import xlwings as xw

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

    config_content = tmp_config_sht.cells(1,1).current_region.value

    user_config = []
    dict_config = {}

    #获取用户设置列表
    for r in config_content:
        dict_config['index'] = r[0]
        dict_config['sheet_name']=r[1]
        dict_config['sheet_code']=r[2]
        dict_config['start_row']=r[3]
        dict_config['start_column']=r[4]
        user_config.append(dict_config)

    end_column = tmp_data_sht.cells(1, 16384).end('left').column



    #获取EXCEL数据表中第一行解析表头，存到数组head_list
    header_list = tmp_data_sht.range(tmp_data_sht.cells(1,1),tmp_data_sht.cells(1,end_column)).value

    header_pk = []

    #查找是否存在pk字段，如果不存在，退出程序，提示找不到pk列，如果存在，把pk字段存成一个列表

    # for  header in header_list:
    #
    #     if not '_' in header:
    #
    #         continue
    #
    #     if header.split('_')[0].upper == 'PK':
    #
    #         header_pk.append(header)
    #
    #








    xw.App.visible = True





