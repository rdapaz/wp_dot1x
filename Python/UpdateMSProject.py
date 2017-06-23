import win32com.client
import pprint
import json
import datetime
import yaml


def pretty_printer(o):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(o)


def runs():
    run_text = """
Tier 1 Sites:
     - T1 - EPCC (20 switches)
     - T1 - Forrestfield (16 switches)
     - T1 - Jandakot Hope (18 switches)
     - T1 - Jandakot Princep (21 switches)
     - T1 - Kewdale (20 switches)
     - T1 - Mt.Claremont (10 switches)
     - T1 - Picton (14 switches)
     - T1 - Stirling (13 switches)
Run 1 (640 Km):
     - T2 - Mandurah (5 switches)
     - T4 - Busselton (2 switches)
     - T4 - Busselton Vasse
     - T4 - Mriver (2 switches)
     - T4 - Bridgetown (1 switch)
     - T2 - Albany (6 switches)
Run 2 (594 Km): 
     - T2 - Northam (5 switches)
     - T4 - Koorda (1 switch)
     - T4 - Merredin (1 switch)
     - T4 - Southern Cross (1 switch)
     - T4 - Kalg (2 switches)
Run 3 (732 Km):
     - T4 - Moora (1 switch)
     - T4 - Jurien (2 switches)
     - T4 - Three-Springs (1 switch)
     - T2 - Geraldton (4 switches)
Run 4 (762 Km):
     - T4 - Waroona (1 switch)
     - T4 - Collie (2 switches)
     - T4 - Narrogin (1 switch)
     - T4 - Kondinin (1 switch)
     - T4 - Katanning (1 switch)
     - T4 - Jerramungup (2 switches)
"""
    return yaml.load(run_text)

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

    def updateRuns(self, runs):
        for tsk in self.mpp.Tasks:
            if tsk.Name in runs:
                tsk.Text4 = runs[tsk.Name]


def pretty_printer(o):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(o)


def main():
    with open(r'C:\Users\rdapaz\projects\wp_dot1x\Python\dot1x_sites_dates.json', 'r') as fin:
        data = json.load(fin)

    pretty_printer(data)
    pp = Project(r'C:\Users\rdapaz\Desktop\Western Power 802.1X Enterprise Wired DeploymentV2.mpp')
    # pp.updateTasks(data=data)
    # pp.doFinalUpdate(data, undo=False)
    switch_runs = runs()
    inv_runs = dict((x, k) for k in switch_runs for x in switch_runs[k])
    pp.updateRuns(runs=inv_runs)

if __name__ == "__main__":
    main()