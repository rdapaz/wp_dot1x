# -*- coding: utf-8 -*-
import win32com.client
import re
import os

proj = win32com.client.gencache.EnsureDispatch('MSProject.Application')
proj.Visible = True

ROOT = r'D:\Projects\Western Power\NEW'
filepath = os.path.join(ROOT, 'Western Power 802.1X wired Deployment.mpp')

filepath = r'C:\Users\ric\Desktop\802.1X Full Deployment.mpp'
proj.FileOpen(filepath)
my_proj = proj.ActiveProject

sites = """
T4 - Waroona|Aug
T1 - Mt.Claremont|July
T2 - Albany Aug
T4 - Bridgetown Aug
T4 - Kalg|Aug
T4 - Busselton|Aug
T4 - Jerramungup|Aug
T4 - Collie Aug
T2 - Geraldton|Aug
T4 - Jurien Sept
T4 - Mriver Sept
T4 - Katanning|Sept
T4 - Kondinin|Sept
T4 - Koorda Sept
T2 - Mandurah|July
T4 - Merredin|Sept
T4 - Moora|Sept
T2 - Northam|July
T1 - Stirling|July
T4 - Narrogin|Aug
T1 - Forrestfield|July
T1 - Jandakot Hope|July
T1 - Jandakot Princep|July
T1 - Kewdale|July
T1 - Picton July
T4 - Southern Cross Aug
T4 - Three-Springs|Sept
""".splitlines()

# sites = {k: v for k, v in [x.split('|') for x in sites if len(x) > 0]}
sites = [x.split('|') for x in sites if len(x) > 0]

print sites
'''
rex = re.compile('T\d{1}.*', re.IGNORECASE)

for tsk in my_proj.Tasks:

    if tsk is None:
        continue
    else:
        if rex.search(tsk.OutlineParent.Name) or rex.search(tsk.Name):
            tsk.Text2 = sites[tsk.Name]
        else:
            pass
'''