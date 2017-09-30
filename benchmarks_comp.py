# import relevant Python modules
import os
import numpy as np
import glob
import math
import csv
import os.path
import xlrd
import xlsxwriter
from xlsxwriter.workbook import Workbook

#NOTE: code will not work if there are more than 56 swimmers
#NOTE: code has issues making borders around the 25 time cells for the 100k at >52 swimmers

#define directory/file containing the data
garnet = 1
white = 2
#input group and week
group = garnet
week = 1
if week == 1:
    file_location = 'section_assignments_'+'wk1'+'.xlsx'
if week == 2:
    file_location = 'section_assignments_'+'wk2'+'.xlsx'
if week == 3:
    file_location = 'section_assignments_'+'wk3'+'.xlsx'
workbook = xlrd.open_workbook(file_location)
if group == garnet:
    sheet = workbook.sheet_by_index(garnet)
if group == white:
    sheet = workbook.sheet_by_index(white)

#sets up number of rows and columns  
rows = np.zeros(sheet.nrows-1)
columns = np.zeros(sheet.ncols-1)

#read data from sheet
cell = sheet.cell_value

#trying to set it up to be fine with rows
#r is number of rows (number of swimmers in group)
#remainder is how many people won't be in full heats
#waves is how many waves of swimmers (different than heats)
#full_heats is number of people in full heats
#stray_kids is excel sheet row number associated with the kids not in full heats
kids = len(rows)+1
remainder = (kids % 8)
waves = (kids / 8)
full_heats = (kids - remainder)
remainder_array = np.arange(1,remainder+1)
stray_kids = np.arange(full_heats+1,full_heats+remainder+1)

#setting up arrays of 8
lane = [0 for x in range(8)]
index = np.arange(1,full_heats,len(lane))
index = np.array(index)
counting = [0,1,2,3,4,5,6,7]

#setting up indexing for the different lanes
for i in range(0,len(lane)):
    lane[i] = index + counting[i]

#creating lane assignments
assignments = []
for i in range(0,len(lane)):
    list = []
    for j in lane[i]:
        list.append((cell(j-1,1)) + ' ' + cell(j-1,2))
    assignments.append(list)

#accounting for stray kids
for i in range(0,remainder):
    assignments[i].append(cell(stray_kids[i]-1,1) + ' ' + cell(stray_kids[i]-1,2))

#setting up number of heats, number_heats is total number of heats (not necessarily all full)
#if the remainder (kids not in full heats) is 5 or above, make a new heat; if there's only 1-4 kids remaindering then make a heat with three people in it for the last heat
if remainder > 4:
    number_waves = (waves - (remainder/(len(lane))) + 1)
else:
    number_waves = (waves - (remainder/(len(lane))))
if isinstance(number_waves/2,int) == True:
    number_heats = (number_waves / 2)
else:
    number_heats = (int(round(number_waves / 2 + 0.1)))
number_half_heats = (number_waves - number_heats)

#filling out not-full heats with 'nobody'
#if there are one or more heats with three people in it, the lanes with only two people in it will have a 'none' instead of a 'nobody', just to differentiate
#if there are some lanes that are totally empty, the heat sheet will have 'empty' in those spots
for i in range(0,len(assignments)):
    if len(assignments[i]) % 2 == 1:
        assignments[i].append('nobody')
    if len(assignments[i]) == 8:
        assignments[i].remove('nobody')
    if len(assignments[i]) == 6 and assignments[i][5] == 'nobody':
        assignments[i].remove('nobody')
for i in range(0,len(assignments)):
    if len(assignments[i]) == 7:
        for j in range(i+1,len(assignments)):
            assignments[j].append('none')
    if len(assignments[i]) == 5:
        for j in range(i+1,len(assignments)):
            assignments[j].append('none')
    if len(assignments[i]) == 8:
        assignments[i].remove('none')
    if len(assignments[i]) == 6 and assignments[i][5] == 'none':
        assignments[i].remove('none')
    if len(assignments[i]) == len(assignments[0]) - 2:
        assignments[i].append('empty')
    if assignments[i][len(assignments[i])-1] == 'empty':
        assignments[i].append('empty')

#arrays for heats
heat_1 = []
for i in range(0,len(lane)):
    list = []
    for j in range(0,2):
        list.append(assignments[i][j])
    heat_1.append(list)

heat_2 = []
for i in range(0,len(lane)):
    list = []
    if len(assignments[i]) == 4:
        for j in range(2,4):
            list.append(assignments[i][j])
        heat_2.append(list)
    if len(assignments[i]) == 5 and assignments[i][4] == 'nobody':
        for j in range(2,5):
            list.append(assignments[i][j])
        heat_2.append(list)
    if len(assignments[i]) == 5 and assignments[i][4] == 'none':
        for j in range(2,5):
            list.append(assignments[i][j])
        heat_2.append(list)

heat_3 = []
if number_heats > 2:
    for i in range(0,len(lane)):
        list = []
        if len(assignments[i]) == 7:
            for j in range(4,7):
                list.append(assignments[i][j])
            heat_3.append(list)
        else:
            for j in range(4,len(assignments[i])):
                list.append(assignments[i][j])
            heat_3.append(list)
        if len(heat_3[i]) == 1:
            heat_3[i].append('blank')
            for j in range(0,len(lane)):
                assignments[j].append('blank')
else:
    print('no heat 3')

#'heat' will be all three heats in a single lists
heat = []
heat.append(heat_1)
heat.append(heat_2)
heat.append(heat_3)


#writing to excel file
#labeling files properly based on which week and group
if group == garnet:
    if week == 1:
        workbook = xlsxwriter.Workbook(file_location.replace('section_assignments_wk1','benchmark_assignments_wk1_garnet'))
    if week == 2:
        workbook = xlsxwriter.Workbook(file_location.replace('section_assignments_wk2','benchmark_assignments_wk2_garnet'))
    if week == 3:
        workbook = xlsxwriter.Workbook(file_location.replace('section_assignments_wk3','benchmark_assignments_wk3_garnet'))
if group == white:
    if week == 1:
        workbook = xlsxwriter.Workbook(file_location.replace('section_assignments_wk1','benchmark_assignments_wk1_white'))
    if week == 2:
        workbook = xlsxwriter.Workbook(file_location.replace('section_assignments_wk2','benchmark_assignments_wk2_white'))
    if week == 3:
        workbook = xlsxwriter.Workbook(file_location.replace('section_assignments_wk3','benchmark_assignments_wk3_white'))

#first tab on excel sheet will be the sheet for the 100 kick benchmark
worksheet = workbook.add_worksheet('100k')

#applying borders to the cells
format = workbook.add_format()
format.set_border(style=1)

#writing column headings
column_number = 0
headings = ['Name','1st 25','2nd 25','3rd 25','4th 25','Total Time']
lane_labels = ['Lane 1','Lane 2','Lane 3','Lane 4','Lane 5','Lane 6','Lane 7','Lane 8']
labels = [[],[],[],[],[],[],[],[]]
for i in range(0,len(lane_labels)):
    labels[i] = [lane_labels[i],'1st 25','2nd 25','3rd 25','4th 25','Total Time']

#writing column headings
for i in range(0,len(headings)):
    for j in range(0,6*len(lane),6):
         worksheet.write(0,j+(column_number+i),headings[i], format)

for i in range(0,len(lane)):
    worksheet.write(0,i*6,labels[i][0], format)
        
#set columns to be wider
for i in range(0,48):
    if i%6 == 0:
        worksheet.set_column(i,i,20)
    else:
        worksheet.set_column(i,i,12)
#set rows to be wider
for i in range(1,len(assignments[0]*2)):
    worksheet.set_row(i,40)    

#placing labels for each heat
heat_headers = []
for i in range(0,48):
    if i%6 == 0:
        heat_headers.append(i)
store = ['Heat 1','Heat 2','Heat 3']
for i in range(0,3):
    for j in heat_headers:
        worksheet.write((i*4+1),j,store[i])
if number_heats == 2:
    for i in range(0,48):
        if i%6 == 0:
            worksheet.write(9,i,'no heat 3')
   
#writing names into excel
#formatting borders, etc.
column_blanks = []
for i in range(0,48):
    if i%6 != 0:
        column_blanks.append(i)
row_blanks = []
#if len(assignments[0]) == 7:
 #   row_blanks = [2,3,6,7,10,11,12]
#if len(assignments[0]) == 5:
 #   row_blanks = [2,3,6,7,8]
#if len(assignments[0]) != 5:
 #   row_blanks = [2,3,6,7,10,11]
#if len(assignments[0]) != 7:
 #   row_blanks = [2,3,6,7,10,11]
#if len(assignments[0]) == 7:
 #   row_blanks.append(12)
#if len(assignments[0]) == 5:
 #   row_blanks.append(8)
if number_heats == 2:
    if len(heat[1][0]) == 2:
        row_blanks = [2,3,6,7]
    if len(heat[1][0]) == 3:
        row_blanks = [2,3,6,7,8]
if number_heats == 3:
    if len(heat[2][0]) == 2:
        row_blanks = [2,3,6,7,10,11]
    if len(heat[2][0]) == 3:
        row_blanks = [2,3,6,7,10,11,12]

#properly mapping assignments entries to excel sheet
if len(assignments[0]) == 7 and len(assignments[1]) == 6:
    assignments[0].remove('blank')
for i in range(0,len(assignments[0])):
    for k in range(0,len(lane)):
        if i%2 == 1:
            worksheet.write_string(i*2+1, 6*k, assignments[k][i], format)
        if i == 0:
            worksheet.write_string(i*2+2, 6*k, assignments[k][i], format)
        if i == 2:
            worksheet.write_string(i*2+2, 6*k, assignments[k][i], format)
        if i == 4:
            if len(assignments[0]) == 5:
                worksheet.write_string(i*2, 6*k, assignments[k][i], format)
            else:
                worksheet.write_string(i*2+2, 6*k, assignments[k][i], format)
        if i == 6:
            worksheet.write_string(i*2, 6*k, assignments[k][i], format)
            
#inserting blank bordered cells for filling out with times           
for i in column_blanks:
    for j in row_blanks:
        worksheet.write_blank(j,i,'',format)
#if only 2 heats, deleting bordered cells from heat three rows
if number_heats == 2:
    for i in range(0,48):
        for j in range(9,12):
            format = workbook.add_format()
            format.set_border(style=0)
            worksheet.write_blank(j,i,'',format)

#adding in the header to differentiate benchmark sets
header = '&C100 Kick Benchmark'
worksheet.set_header(header)



#here comes a ton of repeat code...until i can make it more elegant, this will have to do for creating multiple sheets for the multiple benchmarks...

#second tab on excel sheet will be the sheet for the 100 swim benchmark
worksheet = workbook.add_worksheet('100s')

#applying borders to the cells
format = workbook.add_format()
format.set_border(style=1)

#writing column headings
column_number = 0
headings = ['','1st 25','2nd 25','3rd 25','4th 25','Total Time']
lane_labels = ['Lane 1','Lane 2','Lane 3','Lane 4','Lane 5','Lane 6','Lane 7','Lane 8']
labels = [[],[],[],[],[],[],[],[]]
for i in range(0,len(lane_labels)):
    labels[i] = [lane_labels[i],'1st 25','2nd 25','3rd 25','4th 25','Total Time']

#writing column headings
for i in range(0,len(headings)):
    for j in range(0,6*len(lane),6):
         worksheet.write(0,j+(column_number+i),headings[i], format)

for i in range(0,len(lane)):
    worksheet.write(0,i*6,labels[i][0], format)
        
#set columns to be wider
for i in range(0,48):
    if i%6 == 0:
        worksheet.set_column(i,i,20)
    else:
        worksheet.set_column(i,i,12)
#set rows to be wider
for i in range(1,len(assignments[0]*2)):
    worksheet.set_row(i,40)    

#placing labels for each heat
heat_headers = []
for i in range(0,48):
    if i%6 == 0:
        heat_headers.append(i)
store = ['Heat 1','Heat 2','Heat 3']
for i in range(0,3):
    for j in heat_headers:
        worksheet.write((i*4+1),j,store[i])
if number_heats == 2:
    for i in range(0,48):
        if i%6 == 0:
            worksheet.write(9,i,'no heat 3')
   
#writing names into excel
#formatting borders, etc.
column_blanks = []
for i in range(0,48):
    if i%6 != 0:
        column_blanks.append(i)
row_blanks = []
if len(assignments[0]) == 7:
    row_blanks = [2,3,6,7,10,11,12]
if len(assignments[0]) == 5:
    row_blanks = [2,3,6,7,8]
if len(assignments[0]) != 5:
    row_blanks = [2,3,6,7,10,11]
if len(assignments[0]) != 7:
    row_blanks = [2,3,6,7,10,11]
if len(assignments[0]) == 7:
    row_blanks.append(12)
if len(assignments[0]) == 5:
    row_blanks.append(8)

#properly mapping assignments entries to excel sheet
if len(assignments[0]) == 7 and len(assignments[1]) == 6:
    assignments[0].remove('blank')
for i in range(0,len(assignments[0])):
    for k in range(0,len(lane)):
        if i%2 == 1:
            worksheet.write_string(i*2+1, 6*k, assignments[k][i], format)
        if i == 0:
            worksheet.write_string(i*2+2, 6*k, assignments[k][i], format)
        if i == 2:
            worksheet.write_string(i*2+2, 6*k, assignments[k][i], format)
        if i == 4:
            if len(assignments[0]) == 5:
                worksheet.write_string(i*2, 6*k, assignments[k][i], format)
            else:
                worksheet.write_string(i*2+2, 6*k, assignments[k][i], format)
        if i == 6:
            worksheet.write_string(i*2, 6*k, assignments[k][i], format)
            
#inserting blank bordered cells for filling out with times           
for i in column_blanks:
    for j in row_blanks:
        worksheet.write_blank(j,i,'',format)
#if only 2 heats, deleting bordered cells from heat three rows
if number_heats == 2:
    for i in range(0,48):
        for j in range(9,12):
            format = workbook.add_format()
            format.set_border(style=0)
            worksheet.write_blank(j,i,'',format)

#adding in the header to differentiate benchmark sets
header = '&C100 Swim Benchmark'
worksheet.set_header(header)



#third tab on excel sheet will be the sheet for the 200 broken swim benchmark
worksheet = workbook.add_worksheet('200s')

#applying borders to the cells
format = workbook.add_format()
format.set_border(style=1)

#writing column headings
column_number = 0
headings = ['Name','1st 50','2nd 50','3rd 50','4th 50','Total Time']
lane_labels = ['Lane 1','Lane 2','Lane 3','Lane 4','Lane 5','Lane 6','Lane 7','Lane 8']
labels = [[],[],[],[],[],[],[],[]]
for i in range(0,len(lane_labels)):
    labels[i] = [lane_labels[i],'1st 50','2nd 50','3rd 50','4th 50','Total Time']

#writing column headings
for i in range(0,len(headings)):
    for j in range(0,6*len(lane),6):
         worksheet.write(0,j+(column_number+i),headings[i], format)

for i in range(0,len(lane)):
    worksheet.write(0,i*6,labels[i][0], format)
        
#set columns to be wider
for i in range(0,48):
    if i%6 == 0:
        worksheet.set_column(i,i,20)
    else:
        worksheet.set_column(i,i,12)
#set rows to be wider
for i in range(1,len(assignments[0]*2)):
    worksheet.set_row(i,40)    

#placing labels for each heat
heat_headers = []
for i in range(0,48):
    if i%6 == 0:
        heat_headers.append(i)
store = ['Heat 1','Heat 2','Heat 3']
for i in range(0,3):
    for j in heat_headers:
        worksheet.write((i*4+1),j,store[i])
if number_heats == 2:
    for i in range(0,48):
        if i%6 == 0:
            worksheet.write(9,i,'no heat 3')
   
#writing names into excel
#formatting borders, etc.
column_blanks = []
for i in range(0,48):
    if i%6 != 0:
        column_blanks.append(i)
row_blanks = []
if len(assignments[0]) == 7:
    row_blanks = [2,3,6,7,10,11,12]
if len(assignments[0]) == 5:
    row_blanks = [2,3,6,7,8]
if len(assignments[0]) != 5:
    row_blanks = [2,3,6,7,10,11]
if len(assignments[0]) != 7:
    row_blanks = [2,3,6,7,10,11]
if len(assignments[0]) == 7:
    row_blanks.append(12)
if len(assignments[0]) == 5:
    row_blanks.append(8)

#properly mapping assignments entries to excel sheet
if len(assignments[0]) == 7 and len(assignments[1]) == 6:
    assignments[0].remove('blank')
for i in range(0,len(assignments[0])):
    for k in range(0,len(lane)):
        if i%2 == 1:
            worksheet.write_string(i*2+1, 6*k, assignments[k][i], format)
        if i == 0:
            worksheet.write_string(i*2+2, 6*k, assignments[k][i], format)
        if i == 2:
            worksheet.write_string(i*2+2, 6*k, assignments[k][i], format)
        if i == 4:
            if len(assignments[0]) == 5:
                worksheet.write_string(i*2, 6*k, assignments[k][i], format)
            else:
                worksheet.write_string(i*2+2, 6*k, assignments[k][i], format)
        if i == 6:
            worksheet.write_string(i*2, 6*k, assignments[k][i], format)
            
#inserting blank bordered cells for filling out with times           
for i in column_blanks:
    for j in row_blanks:
        worksheet.write_blank(j,i,'',format)
#if only 2 heats, deleting bordered cells from heat three rows
if number_heats == 2:
    for i in range(0,48):
        for j in range(9,12):
            format = workbook.add_format()
            format.set_border(style=0)
            worksheet.write_blank(j,i,'',format)

#adding in the header to differentiate benchmark sets
header = '&C200 Broken Benchmark'
worksheet.set_header(header)



#fourth tab on excel sheet will be the sheet for the 500 swim benchmark
worksheet = workbook.add_worksheet('500s')

#applying borders to the cells
format = workbook.add_format()
format.set_border(style=1)

#writing column headings
column_number = 0
headings = ['Name','','','','','Total Time']
lane_labels = ['Lane 1','Lane 2','Lane 3','Lane 4','Lane 5','Lane 6','Lane 7','Lane 8']
labels = [[],[],[],[],[],[],[],[]]
for i in range(0,len(lane_labels)):
    labels[i] = [lane_labels[i],'1st 50','2nd 50','3rd 50','4th 50','Total Time']

#writing column headings
for i in range(0,len(headings)):
    for j in range(0,6*len(lane),6):
         worksheet.write(0,j+(column_number+i),headings[i], format)

for i in range(0,len(lane)):
    worksheet.write(0,i*6,labels[i][0], format)
        
#set columns to be the right width
narrow_columns = []
for i in range(0,48):
    if (i+1)%6 !=0 and i%6 !=0:
        narrow_columns.append(i)
for i in range(0,48):
    if i%6 == 0:
        worksheet.set_column(i,i,20)
    if (i+1)%6 == 0:
        worksheet.set_column(i,i,30)
    else:
        worksheet.set_column(1,i,20)
for i in narrow_columns:
    worksheet.set_column(i,i,5)
#set rows to be wider
for i in range(1,len(assignments[0]*2)):
    worksheet.set_row(i,40)

#placing labels for each heat
heat_headers = []
for i in range(0,48):
    if i%6 == 0:
        heat_headers.append(i)
store = ['Heat 1','Heat 2','Heat 3']
for i in range(0,3):
    for j in heat_headers:
        worksheet.write((i*4+1),j,store[i])
if number_heats == 2:
    for i in range(0,48):
        if i%6 == 0:
            worksheet.write(9,i,'no heat 3')

#writing names into excel
#formatting borders, etc.
column_blanks = []
for i in range(0,48):
    if i%6 != 0:
        column_blanks.append(i)
row_blanks = []
if len(assignments[0]) == 7:
    row_blanks = [2,3,6,7,10,11,12]
if len(assignments[0]) == 5:
    row_blanks = [2,3,6,7,8]
if len(assignments[0]) != 5:
    row_blanks = [2,3,6,7,10,11]
if len(assignments[0]) != 7:
    row_blanks = [2,3,6,7,10,11]
if len(assignments[0]) == 7:
    row_blanks.append(12)
if len(assignments[0]) == 5:
    row_blanks.append(8)

#properly mapping assignments entries to excel sheet
if len(assignments[0]) == 7 and len(assignments[1]) == 6:
    assignments[0].remove('blank')
for i in range(0,len(assignments[0])):
    for k in range(0,len(lane)):
        if i%2 == 1:
            worksheet.write_string(i*2+1, 6*k, assignments[k][i], format)
        if i == 0:
            worksheet.write_string(i*2+2, 6*k, assignments[k][i], format)
        if i == 2:
            worksheet.write_string(i*2+2, 6*k, assignments[k][i], format)
        if i == 4:
            if len(assignments[0]) == 5:
                worksheet.write_string(i*2, 6*k, assignments[k][i], format)
            else:
                worksheet.write_string(i*2+2, 6*k, assignments[k][i], format)
        if i == 6:
            worksheet.write_string(i*2, 6*k, assignments[k][i], format)
            
#inserting blank bordered cells for filling out with times           
for i in column_blanks:
    for j in row_blanks:
        worksheet.write_blank(j,i,'',format)
#if only 2 heats, deleting bordered cells from heat three rows
if number_heats == 2:
    for i in range(0,48):
        for j in range(9,12):
            format = workbook.add_format()
            format.set_border(style=0)
            worksheet.write_blank(j,i,'',format)

#if only 2 heats, deleting bordered cells from heat three rows
if number_heats == 2:
    for i in range(0,48):
        for j in range(9,12):
            format = workbook.add_format()
            format.set_border(style=0)
            worksheet.write_blank(j,i,'',format)

#getting rid of the borders on the blank cells
row_blanks.append(0)
for i in narrow_columns:
    for j in row_blanks:
        format = workbook.add_format()
        format.set_border(style=0)
        worksheet.write_blank(j,i,'',format)

#adding in the header to differentiate benchmark sets
header = '&C500 Swim Benchmark'
worksheet.set_header(header)


workbook.close()
