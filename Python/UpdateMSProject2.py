import win32com.client
import pprint
import json
import re


def pretty_printer(o):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(o)


def fy17_sites():
    return  {
            'FORR  ( 3 X VG and 1x AP)': 'T1 - Forrestfield',
            'HO (6 x VG Devices and 2 x RB)': 'T1 - HO',
            'KALG (1 device)': 'T4 - Kalg',
            'MARGR (1 device)': 'T4 - Mriver',
            'NTHM (2 devices)': 'T2 - Northam',
            'PICT (2 devices)': 'T1 - Picton',
            'WRNA(1 x Router and 1 x RB)': 'T4 - Waroona',
            'MTCL (2 x AP)': 'T1 - Mt.Claremont',
            'STRLG (8 x AP)': 'T1 - Stirling',
            'GERN (3 x AP)': 'T2 - Geraldton',
            'Jerra (1 x RB)': 'T4 - Jerramungup',
            'Colli (1 x RB)': 'T4 - Collie'
            }

class Project:

    def __init__(self, filePath):
        self.filePath = filePath
        self.pjApp = win32com.client.gencache.EnsureDispatch('MSProject.Application')
        self.pjApp.Visible = True
        self.pjApp.FileOpen(self.filePath)
        self.mpp = self.pjApp.ActiveProject
        print(self.mpp.FullName)

    def updateTasks(self):
        for tsk in self.mpp.Tasks:
            suffix = ''
            if tsk.Name in fy17_sites():
                m = re.search(r'([A-Z]+)\s+(\(.*\))', tsk.Name, re.IGNORECASE)
                if m:
                    prefix = m.group(1)
                    suffix = m.group(2)
                    tsk_name = f"{fy17_sites()[prefix]} {suffix}"
                    print(tsk_name)
                    tsk.Name = tsk_name


def pretty_printer(o):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(o)


def main():
    path = r'D:\Projects\Western Power\NEW\temp.mpp'
    pp = Project(filePath=path)
    pp.updateTasks()


if __name__ == "__main__":
    main()