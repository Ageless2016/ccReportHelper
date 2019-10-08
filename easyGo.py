#coding=utf-8
import xlwings as xw
import json
import rule
import threading

def start(fn):

    wb = open_workbook(fn)

    if not wb:
        return

    shts = wb.sheets
    wb.app.visible = False
    print("Processing...")
    rule_parser(shts)
    print("done!")
    wb.app.visible = True


def rule_parser(shts):
    #记录异常结果的二维数组
    arrMsg = []
    dict = load_json()
    dict_scene = dict['scene']
    dict_daily = dict['daily']
    scene_rules = dict_scene['rules']
    daily_rules = dict_daily['rules']
    sht_point_check = shts['点检查']
    sht_daily_check = shts['组每天排列检查']

    def exec_rule(sht,rule):

        start_row = 0
        end_row = 0
        province_column = 0
        city_column = 0
        point_column = 0
        scene_column = 0
        testgroup_column = 0


        if sht.name == "点检查":
            start_row = 4
            end_row = sht.range("E:E").last_cell.end('up').row
            province_column = 3
            city_column = 4
            point_column = 5
            scene_column = 6

        elif sht.name == "组每天排列检查":
            start_row = 4
            end_row = sht.range("H:H").last_cell.end('up').row
            province_column = 3
            city_column = 4
            testgroup_column = 8

        for i in range(start_row,end_row+1):

            if not rule.param1 == "":
                param1 = sht.cells(i,rule.param1)
                expression = rule.expression.replace("param1", str(param1.value))
            if not rule.param2 == "":
                param2 = sht.cells(i,rule.param2)
                expression = expression.replace("param2",str(param2.value))
            if not rule.param3 == "":
                param3 = sht.cells(i,rule.param3)
                expression = expression.replace("param3", str(param3.value))
            if not rule.param4 =="":
                param4 = sht.cells(i,rule.param4)
                expression = expression.replace("param4", str(param4.value))

            try:
                exp = eval(expression)
                if exp :

                    province=""   #省
                    city=""         #市
                    point=""        #测试点
                    scene=""         #场景
                    testgroup=""      #测试组

                    if province_column > 0:
                        province = sht.cells(i,province_column).value

                    if city_column > 0:
                        city = sht.cells(i,city_column).value

                    if point_column > 0:
                        point = sht.cells(i,point_column).value

                    if scene_column > 0:
                        scene = sht.cells(i,scene_column).value

                    if testgroup_column > 0:
                        testgroup = sht.cells(i,testgroup_column).value

                    dimension = rule.dimension  #检查维度
                    item = rule.item            #检查项
                    case = rule.case            #检查结果
                    recommend = rule.recommend  #处理建议
                    msg = [province,city,point,scene,testgroup,dimension,item,case,recommend]
                    arrMsg.append(msg)
                    print(msg)
                    light_cells = eval(rule.lightcells)
                    lightcell(light_cells,eval(rule.lightcolor))
            except:
                light_cells = eval(rule.lightcells)
                lightcell(light_cells, eval(rule.lightcolor))

    def lightcell(arr,color):

        for rng in arr:
            rng.color = color

    #点检查
    for srl in scene_rules:
        scene_rule = rule.rule(srl)
        exec_rule(sht_point_check,scene_rule)

    #每日测试组检查
    for drl in daily_rules:
        daily_rule = rule.rule(drl)
        exec_rule(sht_daily_check,daily_rule)

    #写sheet日志
    print("检查完成，正在生成报告...")
    loging(shts,arrMsg)


def load_json():

    f = open('config.json', encoding='utf-8')
    dict = json.load(f)
    return dict

def open_workbook(fn):

    file_ext = fn[-5:]
    if file_ext == '.xlsx':
        wb = xw.Book(fn)
        return wb
    else:
        print("Invalid excel file!")
        return


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

        for i in range(len(arr_msg)):
            for j in range(len(arr_msg[0])):
                sht.cells(i+2,j+1).value = arr_msg[i][j]


    if not shtExist(shts):
        add_log_sht(shts)

    sht = shts("Verifications")

    add_log_data(sht,arr_msg)
