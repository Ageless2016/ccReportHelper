class SheetConfig:
    def __init__(self,sheet_name,sheet_code,start_row,start_column):
        self.sheet_name = str(sheet_name)
        self.sheet_code = str(sheet_code)
        self.start_row = int(start_row)
        self.start_column = int(start_column)
        self.key_self_columns = []
        self.self_columns = []
        self.data_columns = []
        self.key_columns=[]



class NewBlankRow:
    def __init__(self,num):
        dict_blank = {}
        for i in range(num):
            dict_blank[i] = None
        self.value = dict_blank
