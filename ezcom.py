#coding=utf-8
from cmd import *
import os
import ck
import mg
class myCmd(Cmd):

    def __init__(self):
        Cmd.__init__(self)
        Cmd.intro="Easy come, easy go!"
        self.prompt = "> "

    def do_version(self,arg):

        print("v1.0.4  update:20191011")

    def do_whatsnew(self, arg):
        print(
            """
# 2019.10.11 新增查询工具版本号及更新记录命令
# 2019.10.11 修复报表行记录为1时，末行号获取不对的BUG
# 2019.10.10 新增支持多线程同时处理点检查和每天组检查表，提高工具检查效率
# 2019.10.10 修复当报表路径中存在空格，操作系统会在拖入的文件路径两端自动加上双引号，导致报表路径找不到的BUG
            """
        )

    def do_ck(self,arg):
        if not arg:
            self.help_ck()
        arg1 = str(arg).replace('"','')
        if os.path.exists(arg1):
            file_ext = arg1[-5:]
            if file_ext != '.xlsx':
                print("Invalid EXCEL file!")
            ck.start(arg1)
        else:
            print("Path does not exist!")


    def do_mg(self,arg):

        if not arg:
            self.help_mg()
            return
        try:
            arg1 = arg.split()[0]
            arg2 = arg.split()[1]
        except:
            print("The parameters you entered are not enough！")
            return

        arg1_0 = str(arg1).replace('"','')
        arg2_0 = str(arg2).replace('"','')

        if not os.path.exists(arg1_0):
            print("Folder path does not exist!")
            return

        if not os.path.exists(arg2_0):
            print("Template file path does not exist!")
            return

        file_ext = arg2_0[-5:]
        if file_ext != '.xlsx':
            print("Invalid EXCEL template file!")
            return
        else:
            print('starting...')
            mg.run(arg1_0,arg2_0)



    def help_version(self):
        print("Show version info.")


    def do_chickensoup(self,arg):
        print(
        """
            # 什么是有趣的人
              对一切未知报以好奇，对一切不同持以尊重。
              去接纳并喜欢自己，不再遮掩任何欢愉，尴尬，羞涩与失落
              去做一些接地气的事情，让自己用心去喜悦，而不是表情
              然后用你澎湃的生命力去唤醒另外一个人。
              
            # 生活最好的状态
              该看书时看书，改玩时尽情玩，看见优秀的人欣赏，看见落魄的人也不轻视
              有自己的小生活和小情趣
              不用去想改变世界，努力去活出自己。
              没人爱时专注自己，
              有人爱时，有能力拥抱彼此。
              
            # 与生活鏖战，有挫败和沮丧很正常。
              但是只要我们坚持不懈，就能走远一点，再远一点。
              
            # 我们要有最朴素的生活，与最遥远的梦想。
              即使明天天寒地冻，路远马亡。
              
            # 我们总是喜欢拿顺其自然来敷衍人生道路上的荆棘坎坷，
              却很少承认，真正的顺其自然，其实是竭尽所能之后的不强求
              而非两手一摊的不作为
              
            # 你要克服懒惰，你要克服游手好闲，你要克服漫长的白日梦
              你要克服一蹴而就的妄想，你要克服自以为是浅薄的幽默感
              你要独立生长在这世上，不寻找，不依靠
              因为冷漠寡情的人孤独一生。
              你要坚强，振作，自立
              不能软弱，逃避，害怕。
              不能沉溺在消极负面情绪里
              要正面阳光地对待生活和爱你的人。
        
            """
        )


    def precmd(self, line):
        #print("开始解析命令")
        return Cmd.precmd(self, line)

    def postcmd(self, stop, line):
        #print("命令解析完成！")
        return Cmd.postcmd(self, stop, line)

    def help_ck(self):
        print("Invalid command parameter! e.g.: ck filepathname.xlsx")

    def help_mg(self):
        print("Please input file folder path as arg1 and template.xlsx path as arg2!")

    def preloop(self):
        pass

    def postloop(self):
        pass

    def do_exit(self,line):
        return True

    def help_exit(self):
        print("Exit module")

    def do_quit(self,line):
        return True

    def help_quit(self):
        print("bye!")

    def emptyline(self):
        print("Command can not be empty!")

    def default(self,line):#输入无效命令处理办法
        print("No such the command!")

myCmd().cmdloop()
