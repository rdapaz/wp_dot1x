import win32com.client
import pprint
import json
import re
import datetime


def pretty_printer(o):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(o)

class Project:

    def __init__(self, filePath):
        self.filePath = filePath
        self.pjApp = win32com.client.gencache.EnsureDispatch('MSProject.Application')
        self.pjApp.Visible = True
        self.pjApp.FileOpen(self.filePath)
        self.mpp = self.pjApp.ActiveProject
        print(self.mpp.FullName)

    def listTasks(self):
        arr = []
        for tsk in self.mpp.Tasks:
            if tsk.Flag2:
                finish = datetime.datetime.strptime(f'{tsk.Finish}'[:10], '%Y-%m-%d')
                print(tsk.Name, finish.strftime('%d/%m/%Y'), sep="|")
                arr.append([tsk.Name, finish.strftime('%d/%m/%Y')])


def main():
    path = r'C:\Users\ric\Desktop\dot1x_schedule.mpp'
    pp = Project(filePath=path)
    arr = pp.listTasks()
    pretty_printer(arr)

if __name__ == "__main__":
    main()