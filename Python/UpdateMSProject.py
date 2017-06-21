import win32com.client
import pprint
import json
import datetime


def fixDate(s_date):
    """
    s_date is in format YYYY-mm-dd or %Y-%m-%d
    """
    return datetime.datetime.strptime(s_date, '%Y-%m-%d').strftime('%d/%m/%Y')

class Project:

    def __init__(self, filePath):
        self.filePath = filePath
        self.pjApp = win32com.client.gencache.EnsureDispatch('MSProject.Application')
        self.pjApp.Visible = True
        self.pjApp.FileOpen(self.filePath)
        self.mpp = self.pjApp.ActiveProject
        print(self.mpp.FullName)

    def updateTasks(self, data):
        for tsk in self.mpp.Tasks:
            print(tsk.Name)
            if tsk.Name in data:
                tsk.Text2 = f"{tsk.Name} ({data[tsk.Name]['qty']} switches)"
                tsk.Notes = f"{data[tsk.Name]['switches']}"
                tsk.Text3 = tsk.Name

            elif tsk.Name.lower() == 'perform firmware upgrade' and tsk.OutlineParent.Name in data:
                tsk.Finish = fixDate(data[tsk.OutlineParent.Name]['when'])
                tsk.Text3 = tsk.OutlineParent.Name


    def doFinalUpdate(self, data, undo=False):
        for tsk in self.mpp.Tasks:
            if tsk.Text2 and not undo:
                tsk.Name = tsk.Text2
            elif tsk.Text3 and undo:
                if tsk.Text3 in data and tsk.Text3[:6] == tsk.Name[:6]:
                    tsk.Name = tsk.Text3



def pretty_printer(o):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(o)


def main():
    with open(r'C:\Users\ric\projects\wp_dot1x\Python\dot1x_sites_dates.json', 'r') as fin:
        data = json.load(fin)

    pretty_printer(data)
    pp = Project(r'D:\__NEW__\802.1X Full Deployment.mpp')
    # pp.updateTasks(data=data)
    pp.doFinalUpdate(data, undo=False)

if __name__ == "__main__":
    main()