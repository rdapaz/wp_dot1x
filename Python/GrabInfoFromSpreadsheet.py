# coding: utf-8
import re
import os
import win32com.client
import pprint
import json
import datetime


def pretty_printer(o):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(o)


def monToDate(s_mon, p_completed):
    return_val = None
    if p_completed:
        return_val = datetime.date(2017, 5, 31)
    elif s_mon.lower() == 'june':
        return_val = datetime.date(2017, 6, 30)
    elif s_mon.lower() == "july":
        return_val = datetime.date(2017, 7, 31)
    elif s_mon.lower() == "aug":
        return_val = datetime.date(2017, 8, 31)
    elif s_mon.lower() == "sept":
        return_val = datetime.date(2017, 9, 30)
    else:
        return_val = None
    return return_val

class Excel:

    def __init__(self, filePath):
        self.filePath = filePath
        self.xlApp = win32com.client.gencache.EnsureDispatch('Excel.Application')
        self.xlApp.Visible = True
        self.wb = self.xlApp.Workbooks.Open(self.filePath)
        self.sh = self.wb.Worksheets('Site Timetable')
        self.sites = {}
        self.goThroughSheet()

    def goThroughSheet(self):
        eof = self.sh.Range("B65536").End(-4162).Row
        for row in range(3, eof+1):
            site = self.sh.Range(f"C{row}").Value if self.sh.Range(f"C{row}").Value else None
            if site:
                switches = self.sh.Range(f"D{row}").Value
                qty = self.sh.Range(f"F{row}").Value
                if type(qty) == float:
                    qty = int(qty)
                else:
                    qty = 0
                when = self.sh.Range(f"E{row}").Value
                p_completed = True if self.sh.Range(f"L{row}").Value and self.sh.Range(f"L{row}").Value.lower() == "completed" else False
                when = monToDate(when, p_completed).strftime("%Y-%m-%d")
                when = datetime.date(2017, 5, 31).strftime("%Y-%m-%d") if not when else when
                if site not in self.sites:
                    self.sites[site] = {}
                self.sites[site] = {'switches': switches,
                                    'qty': qty,
                                    'when': when
                                    }
        pretty_printer(self.sites)

    def writeToJson(self):
        import json
        try:
            # folderPath = os.path.dirname(self.filePath)
            folderPath = os.getcwd()
            # print(folderPath)
            with open(os.path.join(folderPath, r'dot1x_sites_dates.json'), 'w') as f:
                json.dump(self.sites, f, indent=4)
        except:
            print("Error writing file")


if __name__ == '__main__':
    ROOTDIR = r'D:\Projects\Western Power\NEW'
    filePath = os.path.join(ROOTDIR, 'List of all sites for IOS upgrades v1.xlsx')
    xl = Excel(filePath)
    xl.goThroughSheet()
    xl.writeToJson()