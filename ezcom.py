#coding=utf-8
from cmd import *
import os
import main
class myCmd(Cmd):

    def __init__(self):
        Cmd.__init__(self)
        Cmd.intro="Easy come, easy go!"
        self.prompt = "> "

    def do_ck(self,arg):
        if not arg:
            self.help_ck()
        arg1 = str(arg).replace('"','')
        if os.path.exists(arg1):
            file_ext = arg1[-5:]
            if file_ext != '.xlsx':
                print("Invalid EXCEL file!")
            main.start(arg1)
        else:
            print("Path does not exist!")

    def do_version(self,arg):
        print("v1.0.2  update:20191011")


    def help_version(self):
        print("Show version info.")


    def precmd(self, line):
        #print("开始解析命令")
        return Cmd.precmd(self, line)

    def postcmd(self, stop, line):
        #print("命令解析完成！")
        return Cmd.postcmd(self, stop, line)

    def help_ck(self):
        print("Invalid command parameter! e.g.: ck filepathname.xlsx")

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
