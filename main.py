import openpyxl
import matplotlib.pyplot as plt
import plotly.graph_objects as go
import pandas as pd
import os

# Create storage for data

# Open excel docs
wb = openpyxl.load_workbook('example.xlsx')
sheet = wb['Counts']
sheet2 = wb['Smartpen Check-out']

# Get data from spreadsheet
num_pen_tot = int(sheet['B1'].value)
num_pen_out = int(sheet['B2'].value)
num_pen_in  = int(num_pen_tot - num_pen_out)

num_notebok_tot  = int(sheet['B5'].value)
num_notebook_out = int(sheet['B6'].value)
num_notebook_in  = int(num_notebok_tot - num_notebook_out)

num_headset_tot  = int(sheet['B9'].value)
num_headset_out  = int(sheet['B10'].value)
num_headset_in   = int(num_headset_tot - num_headset_out)

# Scraping for date, person, item data
# Tally the number of each item for each month
class Entry:
    def __init__(self, row):
        self.row = row
        self.pen = 0
        self.notebook = 0
        self.headset = 0
        self.techie = ""
        self.month = 0

    def get_counts(self):
        r = str(self.row)

        # Pen check
        if(sheet2['L' + r].value != None):
            self.pen = int(sheet2['L' + r].value)
        # Notebook check
        if(sheet2['M' + r].value != None):
            self.notebook = int(sheet2['M' + r].value)
        # Headset check
        if(sheet2['N' + r].value != None):
            self.headset = int(sheet2['N' + r].value)
        # Techie
        if(sheet2['O' + r].value != None and sheet2['O' + r].value != 'N/A'):
            t = str(sheet2['O' + r].value)
            self.techie = t.strip()
        # Date
        if(sheet2['P' + r].value != None and sheet2['P' + r].value != 'N/A' and sheet2['P' + r].value != '#VALUE!'):
            x = sheet2['P' + r].value
            self.month = int(x.month)
            
i = 2
totalPen = 0
totalNotebook = 0
totalHeadset = 0
months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
monthC = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
techies = ['Antonio', 'Becky', 'Chris', 'Matthew', 'Robert', 'Sam', 'Sarah', 'Stephanie', 'Samantha']
techieC = [0, 0, 0, 0, 0, 0, 0, 0]

while i < 1000:
    ent = Entry(i)
    ent.get_counts()

    totalPen += ent.pen
    totalNotebook += ent.notebook
    totalHeadset += ent.headset
    if(ent.month != 0):
        monthC[ent.month - 1] += 1
    
    if ent.techie in techies:
        if(ent.techie == 'Antonio'):
            techieC[0] += 1
        elif(ent.techie == 'Becky'):
            techieC[1] += 1
        elif(ent.techie == 'Chris'):
            techieC[2] += 1
        elif(ent.techie == 'Matthew'):
            techieC[3] += 1
        elif(ent.techie == 'Robert'):
            techieC[4] += 1
        elif(ent.techie == 'Sam' or ent.techie == 'Samantha'):
            techieC[5] += 1
        elif(ent.techie == 'Sarah'):
            techieC[6] += 1
        elif(ent.techie == 'Stephanie'):
            techieC[7] += 1

    i += 1

# Make graphs using data

# Line graph showing month distribution
fig3 = go.Figure()
fig3.add_trace(
    go.Scatter(
        x = months,
        y = monthC,
        mode = 'lines',
        name = 'lines',
        line_color = 'rgb(0,176,246)',
    )
)

fig3.update_layout(
    title = 'Check-Out Distribution by Month',
    font = dict(color='#909090'),
    xaxis = dict(
        title = 'Month',
        titlefont = dict(
            family = 'Arial, sans-serif',
            size = 12,
            color = '#909090'
        ),
        showticklabels = True,
        tickfont = dict(
            family = 'Arial, sans-serif',
            size = 12,
            color = '#909090'
        ),
    ),
    yaxis = dict(
        range = [0, 30],
        title = 'Number Checked-Out',
        titlefont = dict(
            family = 'Arial, sans-serif',
            size = 12,
            color = '#909090',
        ),
        showticklabels = True,
        tickfont = dict(
            family = 'Arial, sans-serif',
            size = 12,
            color = '#909090'
        )
    ),
)

# Bar graph showing techie
color = ['#004C6D', '#145C80', '#246D94', '#327EA9', '#4090BE', '#4EA2D3', '#5CB4E9', '#6AC7FF']
trace = go.Bar(
    y = techies,
    x = techieC,
    orientation = 'h',
    marker = dict(
        color = color,
    )
)

data = [trace]

layout = go.Layout(
    title = '2019 Techie Check-Out Rates',
    font = dict(color='#909090'),
    xaxis = dict(
        title = 'Number Checked-Out',
        titlefont = dict(
            family = 'Arial, sans-serif',
            size = 12,
            color = '#909090'
        ),
        showticklabels = True,
        tickfont = dict(
            family = 'Arial, sans-serif',
            size = 12,
            color = '#909090'
        ),
    ),
    yaxis = dict(
        range = [-1, 8],
        title = 'Techie',
        titlefont = dict(
            family = 'Arial, sans-serif',
            size = 12,
            color = '#909090',
        ),
        showticklabels = True,
        tickfont = dict(
            family = 'Arial, sans-serif',
            size = 12,
            color = '#909090'
        )
    ),
)

fig2 = go.Figure(data = data, layout = layout)

# Bar graph showing # given out per year
color = ['#009ACD', '#ADD8E6', '#63D1F4']
trace = go.Bar(
    x = ['Pen', 'Notebook', 'Headset'],
    y = [num_pen_out, num_notebook_out, num_headset_out],
    marker = dict(
        color = color
    )
)

data = [trace]

layout = go.Layout(
    title = '2019 Inventory',
    font = dict(color='#909090'),
    xaxis = dict(
        title = 'Type',
        titlefont = dict(
            family = 'Arial, sans-serif',
            size = 12,
            color = '#909090'
        ),
        showticklabels = True,
        tickfont = dict(
            family = 'Arial, sans-serif',
            size = 12,
            color = '#909090'
        ),
    ),
    yaxis = dict(
        range = [0, 225],
        title = 'Number of Checked-Out',
        titlefont = dict(
            family = 'Arial, sans-serif',
            size = 12,
            color = '#909090',
        ),
        showticklabels = True,
        tickfont = dict(
            family = 'Arial, sans-serif',
            size = 12,
            color = '#909090'
        )
    )
)

fig1 = go.Figure(data = data, layout = layout)

# Line graph showing month progression
# Animated bar graph showing progression

# Save graphs as pdf
fig1.write_image("graphs/nums.pdf")
fig2.write_image("graphs/techies.pdf")
fig3.write_image("graphs/months.pdf")

print('Success')

# Opens graphs in a local browser
fig1.show()
fig2.show()
fig3.show()