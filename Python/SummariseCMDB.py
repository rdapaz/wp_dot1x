# coding: utf-8
import re
import os
import win32com.client
import pprint
import json


def pretty_printer(o):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(o)


class MyExcel:

    def __init__(self, filePath):
        self.filePath = filePath
        self.xlApp = win32com.client.gencache.EnsureDispatch('Excel.Application')
        self.xlApp.Visible = True
        self.workBook = self.xlApp.Workbooks.Open(self.filePath)
        self.sh = self.workBook.Worksheets(2)
        print(self.sh.Name)
        self.goThroughSheet()

    def goThroughSheet(self):
        my_eof = self.sh.Range("A65536").End(-4162).Row
        print(my_eof)

    def writeToJson(self):
        import json
        try:
            folderPath = os.path.dirname(self.filePath)
            print(folderPath)
            with open(os.path.join(folderPath, r'scripts.json'), 'w') as f:
                json.dump(self.queries, f, indent=4)
        except:
            print("Error writing file")


if __name__ == '__main__':
    ROOTDIR = r'C:\Users\ric\projects\wp_dot1x\Resources'
    filePath = os.path.join(ROOTDIR, 'CMDB.xlsx')
    xl = MyExcel(filePath)