import openpyxl
import matplotlib.pyplot as plt
import plotly.graph_objects as go
import pandas as pd
import os

# Create storage for data

# Open excel docs
wb = openpyxl.load_workbook('example.xlsx')
sheet = wb['Counts']

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

# Scraping for date data
month = ['May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']


# Make graphs using data

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
fig1.show()


# Line graph showing month progression
# Animated bar graph showing progression

# Save graphs as pdf
fig1.write_image("graphs/fig1.pdf")
