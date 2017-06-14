# -*- coding: utf-8 -*-
import win32com.client
import re
import os
import pprint

try:
    from cStringIO import StringIO
except:
    from StringIO import StringIO

# Writing to a buffer
output = StringIO()


def pretty_print(o):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(o)


ROOT = r'D:\Projects\Western Power\NEW'

proj = win32com.client.gencache.EnsureDispatch('MSProject.Application')
proj.Visible = True

filepath = os.path.join(ROOT, 'Western Power 802.1X wired Deployment.mpp')

proj.FileOpen(filepath)
my_proj = proj.ActiveProject

rex = re.compile('T\d{1}.*', re.IGNORECASE)

for tsk in my_proj.Tasks:

    if tsk is None:
        continue
    else:
        if rex.search(tsk.Name):
            print tsk.Name
        SPACES = '    '
        task_desc = None
        if tsk.OutlineChildren.Count == 0:
            task_desc = '{}{}: {}d|{}'.format(
                                                SPACES * int(tsk.OutlineLevel -1),
                                                tsk.Name,
                                                tsk.Duration/480.0,
                                                tsk.ResourceNames
                                              )
        else:
            task_desc = '{}{}:'.format(
                                        SPACES * int(tsk.OutlineLevel -1),
                                        tsk.Name
                                      )
        print >>output, task_desc

json_text = output.getvalue()
with open('tasks.yaml', 'w') as f:
    f.write(json_text)

