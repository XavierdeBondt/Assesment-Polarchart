import xlrd

# Path of excel input file
loc = ("input.xlsm")

# Open workbook's first sheet
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)    # Note that we take the Sheet index to be 0

# How the title, date and author should be given:
row_title = 0
column_title = 1
row_author = 1
column_author = 1
row_date = 2
column_date = 1

# We use datetime to format our date because of the Excel 1900-'error'.
import datetime
title = sheet.cell_value(row_title, column_title)
unformatted_date = datetime.date(1900,1,1) + datetime.timedelta(sheet.cell_value(row_date, column_date) - 2)
author = sheet.cell_value(row_author, column_author)
date = '{:%d/%m/%Y}'.format(unformatted_date)


# How the excel sheet should be built:
row_firstvalues = 5
column_category = 0
column_color = 1
column_slicename = 2
column_sliceamount = 3
column_slicepercentage = 4

# Go through second column to get all HEX-code values of the colors
indexcounter = []
bigcolors = []
for i in range(row_firstvalues, sheet.nrows):
    inputcolor = sheet.cell_value(i,column_color)
    if ('#' + inputcolor) not in bigcolors:
        bigcolors.append('#' + inputcolor) 
        indexcounter.append(i-row_firstvalues)    # Meanwhile we take note of where these colors start... (**) 

# Amount of big parts are given by the amount of colors
bigparts = len(bigcolors)            

# Convert index counter to a useable amount of slices per big part (**)
slicecount = []
for i in range(bigparts):
    if i == (bigparts -1):
        slicecount.append(sheet.nrows - indexcounter[i] - row_firstvalues)
    else:
        slicecount.append(indexcounter[i+1] - indexcounter[i])

# Read out Big part labels
biglabels = []
for i in range(row_firstvalues, sheet.nrows):
    inputlabel = sheet.cell_value(i,column_category)
    if inputlabel not in biglabels:
        biglabels.append(inputlabel)

# Use slicecount for the big parts
bigvalues = []
for i in range(bigparts):
    bigvalues.append(slicecount[i])

# Get all the small labels
smalllabels = []
smalltexts=[]
for i in range(row_firstvalues, sheet.nrows):
    smalltexts.append(sheet.cell_value(i,column_slicename))
    smalllabels.append(sheet.cell_value(i,column_slicename)+str(i))

# Converts to a color scheme for the small donut
smallcolors = []
for j in range(bigparts):
    for i in range(slicecount[j]):
        smallcolors.append(bigcolors[j])

# Read percentages from file
percentages = []
for i in range(row_firstvalues, sheet.nrows):
    percentages.append(sheet.cell_value(i,column_slicepercentage) * 100)

# Convert percentages to a 2D-array
rlist = []
for j in range(bigparts):
    values = []
    minimum = indexcounter[j]
    if j == (bigparts - 1):
        maximum = sheet.nrows  - row_firstvalues
    else :
        maximum = indexcounter[j+1]
    # Add zero's in front and behind so we can use Scatterpolar
    for i in range(minimum, maximum):
        values.append(0)
        for k in range(0,3): 
            # Triple this percentage we have three points
            values.append(percentages[i])
    values.append(0)
    rlist.append(values)

# Read slice values and convert them into percentages of 360 degrees
slicepercentages = []
smallvalues = []
slicesum = 0
for i in range(row_firstvalues, sheet.nrows):
    value = sheet.cell_value(i,column_sliceamount)
    slicesum = slicesum + value
    smallvalues.append(value) 

for i in range(len(smallvalues)):
    slicepercentages.append(smallvalues[i] / slicesum)


# Percentage of 360 degrees to angles
angles = []
angle = 0
for i in range(len(slicepercentages)):
    angle = angle + slicepercentages[i]*360
    angles.append(angle)
# Create radialangle
radialangle = 90 - int(slicesum / 4)*angles[0]
# Convert angles to a 2D-array
tlist = []
for j in range(bigparts):
    values = []
    minimum = indexcounter[j] 
    
    if j == (bigparts - 1):
        maximum = sheet.nrows  - row_firstvalues
    else :
        maximum = indexcounter[j+1] 
    
    # Add zero's in front and behind so we can use Scatterpolar
    for i in range(minimum, maximum):
        values.append(0)
        # Take the two points before the angle
        values.append(angles[i]-angles[0])
        values.append(angles[i]-(angles[0]/2))
        values.append(angles[i])
    values.append(0)
    tlist.append(values)


# Cheat so we can create up to 6 categories
if(bigparts < 6):
    for i in range(bigparts , 6):
        bigcolors.append(bigcolors[0])
        rlist.append([])
        tlist.append([])

print(percentages)
print(rlist)