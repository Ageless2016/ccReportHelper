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










    xw.App.visible = True





