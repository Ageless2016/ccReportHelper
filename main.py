import xlwings as xw
import json
import rule

def main():

    fn = r'C:\Users\CC\Desktop\telecom\CQT-多网指标汇总_V5-样本.xlsx'
    wb = open_workbook(fn)
    shts = wb.sheets
    sht_point_check = shts['点检查']
    sht_daily_check = shts['组每天排列检查']
    dict = load_json()
    dict_scene = dict['scene']
    dict_daily = dict['daily']
    scene_rules = dict_scene['rules']
    daily_rules = dict_daily['rules']

    #点检查
    for srl in scene_rules:
        scene_rule = rule.rule(srl)
        rule_parser(sht_point_check,scene_rule)

    #每日测试组检查
    for drl in daily_rules:
        daily_rule = rule.rule(drl)
        rule_parser(sht_daily_check,daily_rule)


def rule_parser(sht,rule):


    def scene_check(rule):

        start_row = 4
        end_row = sht.range("E:E").last_cell.end('up').row

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
                if eval(expression):
                    light_cells = eval(rule.lightcells)
                    lightcell(light_cells,eval(rule.lightcolor))
                    print(rule.case,",",rule.recommend)
            except:
                pass



    def daily_check(rule):
        pass


    def lightcell(arr,color):

        for rng in arr:
            rng.color = color


    if sht.name == "点检查":

        scene_check(rule)

    elif sht.name == "组每天排列检查":

        daily_check(rule)




def load_json():

    f = open('config.json', encoding='utf-8')
    dict = json.load(f)
    return dict

def open_workbook(fn):

    wb = xw.Book(fn)
    return wb


def loging(shts,logData):

    def shtExist(shtname, shts):

        for sht in shts:
            if sht.name == shtname:
                return True

        return False

    if not shtExist("Verifications",shts):
        shts.add("Verifications", after=shts(shts.count))

    sht = shts("Verifications")

    def create_log_title(sht):

        tb_title = ['省份','城市','测试点','场景','分类','测试组','核查项','核查结果','建议']
        i=0
        for title in tb_title:
            i=i+1
            sht.cells(1,i).value = title

    def log_input(sht,log_data):
        r = sht.range("A1").current_region.last_cell.row
        i=0
        for log in log_data:
            i=i+1
            sht.cells(r+1,i).value = log


    create_log_title(sht)
    log_input(sht,logData)
    sht.autofit('c')



if __name__ == '__main__':
    main()