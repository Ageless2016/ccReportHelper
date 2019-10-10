#coding=utf-8
import xlwings as xw
import json
import rule
import threading

# 定义全局数组，接收检查异常信息
arrMsg = []
def start(fn):

    #加载工作簿
    wb = xw.Book(fn)
    shts = wb.sheets

    #加载config.json
    f = open('config.json', encoding='utf-8')
    dict= json.load(f)
    dict_scene = dict['scene']
    dict_daily = dict['daily']
    scene_rules = dict_scene['rules']
    scene_shtname = dict_scene['shtname']
    daily_rules = dict_daily['rules']
    daily_shtname = dict_daily['shtname']

    #获取点检查表配置
    scene_cfg = {}
    scene_cfg['start_row'] =dict_scene['start_row']
    scene_cfg['province'] =dict_scene['province']
    scene_cfg['city'] =dict_scene['city']
    scene_cfg['spot'] =dict_scene['spot']
    scene_cfg['scene'] =dict_scene['scene']
    scene_cfg['team'] =dict_scene['team']

    #获取测试组每天检查表配置
    daily_cfg = {}
    daily_cfg['start_row'] =dict_daily['start_row']
    daily_cfg['province'] =dict_daily['province']
    daily_cfg['city'] =dict_daily['city']
    daily_cfg['spot'] =dict_daily['spot']
    daily_cfg['scene'] =dict_daily['scene']
    daily_cfg['team'] =dict_daily['team']

    #判断sheet名称是否存在
    try:
        scene_sht = shts[scene_shtname]
        daily_sht = shts[daily_shtname]
    except:
        print("%s 或 %s Sheet表不存在！"%(scene_shtname,daily_shtname))
        return

    #设置excel程序不可见，禁止屏幕刷新提高excel操作效率
    wb.app.visible = False
    wb.app.screen_updating = False

    # rules_parser(daily_sht,daily_cfg,daily_rules)
    # rules_parser(scene_sht,scene_cfg,scene_rules)

    t_scene = threading.Thread(target=rules_parser,args=(fn,scene_shtname,scene_cfg,scene_rules))
    t_daily = threading.Thread(target=rules_parser,args=(fn,daily_shtname,daily_cfg,daily_rules))

    t_scene.start()
    t_daily.start()

    t_scene.join()
    t_daily.join()

    #写sheet日志
    print("检查完成，正在生成报告...")
    loging(shts,arrMsg)

    #恢复excel程序为可见，启用屏幕刷新
    wb.app.visible = True
    wb.app.screen_updating=True

    print("done!")


def rules_parser(fn,shtname,dict_cfg,rules):

    wb = xw.Book(fn)
    sht = wb.sheets[shtname]
    start_row = dict_cfg['start_row']
    end_row = sht.cells(start_row,dict_cfg['city']).end('down').row

    for cfg_rule in rules:

        rul = rule.rule(cfg_rule)

        for i in range(start_row,end_row+1):

            if not rul.param1 == "":
                param1 = sht.cells(i,rul.param1)
                expression = rul.expression.replace("param1", str(param1.value))
            if not rul.param2 == "":
                param2 = sht.cells(i,rul.param2)
                expression = expression.replace("param2",str(param2.value))
            if not rul.param3 == "":
                param3 = sht.cells(i,rul.param3)
                expression = expression.replace("param3", str(param3.value))
            if not rul.param4 =="":
                param4 = sht.cells(i,rul.param4)
                expression = expression.replace("param4", str(param4.value))

            try:
                exp = eval(expression)

                if exp :

                    if dict_cfg['province'] != "":
                        province = sht.cells(i,dict_cfg['province']).value
                    else:
                        province=""

                    if dict_cfg['city'] != "":
                        city = sht.cells(i,dict_cfg['city']).value
                    else:
                        city=""

                    if dict_cfg['spot'] != "":
                        spot = sht.cells(i,dict_cfg['spot']).value
                    else:
                        spot=""


                    if dict_cfg['scene'] != "":
                        scene = sht.cells(i,dict_cfg['scene']).value
                    else:
                        scene=""

                    if dict_cfg['team'] != "":
                        team = sht.cells(i,dict_cfg['team']).value
                    else:
                        team=""


                    dimension = rul.dimension  #检查维度
                    item = rul.item            #检查项
                    case = rul.case            #检查结果
                    recommend = rul.recommend  #处理建议
                    msg = [province,city,spot,scene,team,dimension,item,case,recommend]
                    arrMsg.append(msg)
                    print(msg)
                    light_cells = eval(rul.lightcells)
                    lightcell(light_cells,eval(rul.lightcolor))
            except:
                light_cells = eval(rul.lightcells)
                lightcell(light_cells, eval(rul.lightcolor))



def lightcell(arr,color):

    for rng in arr:
        rng.color = color



def loging(shts,arr_msg):

    sht = None

    def shtExist(shts):

        for sht in shts:
            if sht.name == 'Verifications':
                return True

        return False

    def add_log_sht(shts):

        shts.add("Verifications", after=shts['点检查'])
        log_title = ['省份','城市','测试点','场景','测试组','检查维度','核查项','核查结果','处理建议']
        i=0
        for t in log_title:
            i=i+1
            shts['Verifications'].cells(1,i).value = t
            shts['Verifications'].cells(1,i).color = (217,217,217)

    def add_log_data(sht,arr_msg):

        sht.range("A2:I1048576").clear
        sht.range("A2").value = arr_msg


    if not shtExist(shts):
        add_log_sht(shts)

    sht = shts("Verifications")

    add_log_data(sht,arr_msg)
