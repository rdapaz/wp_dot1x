import win32com.client
import datetime

slide_objects = {
    0: 'Flowchart: Extract 39',
    1: 'Flowchart: Extract 40',
    2: 'Flowchart: Extract 41',
    3: 'Flowchart: Extract 42',
    4: 'Flowchart: Extract 43',
    5: 'Flowchart: Extract 44',
    6: 'Flowchart: Extract 45',
}


def calculateOffset(dt_string):
    dt_start = datetime.date(2017, 10, 30)
    dt_finish = datetime.date(2018, 3, 25)
    dt = datetime.datetime.strptime(dt_string, '%d/%m/%Y')
    delta1 = dt.date() - dt_start
    delta2 = dt_finish - dt_start
    offset = 40 + (delta1/delta2)*876
    return offset


data = """
Head Office Deployment|30/10/2017
Tier 1 Deployment|14/12/2017
Deployment Phase 2|17/01/2018
Deployment Phase 3|02/02/2018
Deployment Phase 4|13/02/2018
Deployment Phase 5|22/02/2018
Project Closeout|02/03/2018
""".splitlines()

data = [x.split('|') for x in data if len(x) > 0]

pp = win32com.client.gencache.EnsureDispatch('Powerpoint.Application')
pp.Visible = True

deck = pp.Presentations.Open(r'C:\Users\ric\Desktop\20171109 - Protect the Network - Project Update!.pptx')
slide = deck.Slides(1)

tbl = slide.Shapes("Table 13").Table

for idx, row in enumerate(data):
    tsk_name, finish = row
    tbl.Cell(idx+2,1).Shape.TextFrame.TextRange.Text = idx+1
    tbl.Cell(idx+2,2).Shape.TextFrame.TextRange.Text = tsk_name
    tbl.Cell(idx+2,3).Shape.TextFrame.TextRange.Text = finish
    tbl.Cell(idx+2,4).Shape.TextFrame.TextRange.Text = finish


display_dates = [
                '30/10/2017',
                '13/11/2017',
                '27/11/2017',
                '11/12/2017',
                '25/12/2017',
                '08/01/2018',
                '22/01/2018',
                '05/02/2018',
                '19/02/2018',
                '05/03/2018',
                '19/03/2018',
                ]

tbl = slide.Shapes("Table 154").Table
for idx, dt in enumerate(display_dates):
    tbl.Cell(1, idx+1).Shape.TextFrame.TextRange.Text = dt


dates = [x[1] for x in data]

for idx, dt in enumerate(dates):
    slide.Shapes(f'{slide_objects[idx]}').Left = calculateOffset(dt)
    slide.Shapes(f'{slide_objects[idx]}').TextFrame.TextRange.Text = idx+1

'''
Table 13      ID
Table 154     MON 8/5
Table 11      Planned Completion:  53%
'''