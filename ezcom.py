#coding=utf-8
from cmd import *
import os
import main
import threading

class myCmd(Cmd):

    def __init__(self):
        Cmd.__init__(self)
        Cmd.intro="Easy come, easy go!"
        self.prompt = "> "

    def do_ck(self,arg):
        if not arg:
            self.help_ck()
        elif os.path.exists(arg):
            file_ext = arg[-5:]
            if file_ext != '.xlsx':
                print("Invalid EXCEL file!")
            main.start(arg)
        else:
            print("Path does not exist!")

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