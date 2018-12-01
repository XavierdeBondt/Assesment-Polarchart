# Creating the charts
import plotly
import plotly.graph_objs as go

# Importing common variables calculated by the excelsheet input.xlsm
import common

# Creation of the outer-pie/donut 
bigtrace = go.Pie(labels=common.biglabels, values=common.bigvalues,
               sort=False,
               visible=True,
               showlegend=False,
               textinfo='label', 
               hole = 0.96,
               domain = {"x": [0, 1],
                         "y": [0, 1]},
               direction = 'clockwise',
               outsidetextfont = dict(size = 18),
               marker=dict(colors=common.bigcolors, 
                           line=dict(color='#000000', width=1)))

# Creation of the inner-pie/donut
smalltrace = go.Pie(labels=common.smalllabels, values=common.smallvalues, text=common.smalltexts,
               sort=False,
               showlegend=False,
               textinfo='text', 
               textposition='inside',
               hole = 0.90,
               domain = {"x": [0.02, 0.98],
                         "y": [0.02, 0.98]},
               direction = 'clockwise',
               insidetextfont = dict(size = 10),
               marker=dict(colors=common.smallcolors,
                           line=dict(color='#000000', width=1)))

# Creation of the polar-data (max 6)
polardata = [
    go.Scatterpolar(
        r = common.rlist[0],
        theta = common.tlist[0],
        mode = 'lines',
        fill = 'toself',
        fillcolor = common.bigcolors[0],
        line =  dict(
            width = 1,
            color = 'black'
        )
    ),
    go.Scatterpolar(
        r = common.rlist[1],
        theta = common.tlist[1],
        mode = 'lines',
        fill = 'toself',
        fillcolor = common.bigcolors[1],
        line =  dict(
            width = 1,
            color = 'black'
        )
    ),
    go.Scatterpolar(
        r = common.rlist[2],
        theta = common.tlist[2],
        mode = 'lines',
        fill = 'toself',
        fillcolor = common.bigcolors[2],
        line =  dict(
            width = 1,
            color = 'black'
        )
    ),
    go.Scatterpolar(
        r = common.rlist[3],
        theta = common.tlist[3],
        mode = 'lines',
        fill = 'toself',
        fillcolor = common.bigcolors[3],
        line =  dict(
            width = 1,
            color = 'black'
        )
    ),
    go.Scatterpolar(
        r = common.rlist[4],
        theta = common.tlist[4],
        mode = 'lines',
        fill = 'toself',
        fillcolor = common.bigcolors[4],
        line =  dict(
            width = 1,
            color = 'black'
        )
    ),
    go.Scatterpolar(
        r = common.rlist[5],
        theta = common.tlist[5],
        mode = 'lines',
        fill = 'toself',
        fillcolor = common.bigcolors[5],
        line =  dict(
            width = 1,
            color = 'black'
        )
    )
]

# Global layout
layout = go.Layout(
    title = common.title,
    polar = dict(
        radialaxis = dict(
            visible = True,
            range = [0,100],
            angle=common.radialangle,
            ticksuffix='%',
            tickmode="array",
            tickvals=[0,25,50,75,100] # It will always be percentages
        ),
        domain = dict(
            x = [0.10, 0.90],
            y = [0.10, 0.90]
        ),
        angularaxis=dict(
            visible=True,
            rotation=90,
            direction="clockwise",
            ticklen = 25,
            tickmode="array",
            tickvals=common.angles,
            showticklabels=False,
            linewidth=1
        )
    ),
    annotations=[
        dict(
            x=0.50,
            y=0.077,
            showarrow=False,
            font = dict(
                size=9,
                color='#D3D3D3'
            ),
            text='© 2018 XavierdeBondt',
            xref='paper',
            yref='paper'
        ),
        dict(
            x=0.5,
            y=1.08,
            showarrow=False,
            text=common.author,
            xref='paper',
            yref='paper'
        ),
        dict(
            x=0.5,
            y=1.05,
            showarrow=False,
            text=common.date,
            xref='paper',
            yref='paper'

        )],
    showlegend = False
)

# Combining both graphs into one figure
data = polardata + [smalltrace, bigtrace]
fig = go.Figure(data, layout=layout)

from tkinter import filedialog
import plotly.io as pio
filename = filedialog.asksaveasfilename()
if filename is "":
    raise SystemExit
# Write to jpg and html file!
#pio.write_image(fig, filename + ".jpeg", width=1100, height=700, scale=1)
import ntpath
plotly.offline.plot(fig, filename=filename + ".html", 
                    image="jpeg", image_filename=ntpath.basename(filename), image_width=1100, image_height=700)
