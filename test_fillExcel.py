import xlwings as xw

wb = xw.Book(r'C:\Users\CC\Desktop\电信巡检2019_CQT指标提取工具V2\template.xlsx')

sht = wb.sheets[0]

column_count = 230

content_start_row = 3
content_end_row = 2

header_list = sht.range(sht.cells(1,1),sht.cells(1,column_count)).value

print(header_list)

pk_list = []
dict_data = {}

for i in range(len(header_list)):

    dict_data[i] = header_list[i]

    header = header_list[i]

    if header is None or header=='':
        continue

    if header.split('_')[0] == 'PK':
        pk_list.append(header)


tup = tuple(pk_list)

d_template ={tup:dict_data}

contents = {}

if content_end_row < content_start_row:
    pk_list = ['广东省','东莞市','地铁2号线(东莞火车站-虎门火车站)','车厢','地铁']
    pk_tup = tuple(pk_list)
    empty_row = {}
    for i in range(column_count):
        empty_row[i] = ''

    pk_tup1 = ('广东省', '广州市', '地铁2号线(东莞火车站 - 虎门火车站)', '车厢', '地铁')
    pk_tup2 = ('广东省','广州市','地铁6号线(浔峰岗-香雪)','车厢','地铁')
    pk_tup3 = ('广东省', '广州市', '地铁9号线(高增-飞鹅岭)', '车厢', '地铁')
    pk_tup4 = ('广东省', '广州市', '海珠有轨电车(广州塔-万胜围)', '车厢', '地铁')





def insert_data(pk_tup,sht_row):
    if pk_tup in contents.keys():
        #获取元祖key对应的字典
        d_row = contents[pk_tup]

        for head in header_list:
            for i in range(len(sht_row)):
                print("根据元祖key，填充d_row")


        print(d_row)
    else:
        empty_row = {}
        for i in range(column_count):
            empty_row[i] = ''

        contents[pk_tup] = empty_row

        print(contents)


while True:

    a = input()

    insert_data(eval(a))






