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
#NOTE: code "greener" will include the heat sheets, "green" simply will not

#define directory/file containing the data
garnet = 1
white = 2
#input group and week
group = white
week = 3

#file location
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
    
#heat_2 issue still needs help...i think...not sure...need to check for every section
heat_2 = []
for i in range(0,len(lane)):
    list = []
    if len(assignments[i]) == 4 or 7:
        for j in range(2,4):
            list.append(assignments[i][j])
        heat_2.append(list)
    if len(assignments[i]) == 5:
        if number_heats == 3:
            list = []
            for j in range(2,4):
                list.append(assignments[i][j])
            heat_2.append(list)
            heat_2.remove(heat_2[i])
        else:
            if remainder < 5:
                list = []
                for j in range(2,5):
                    list.append(assignments[i][j])
                heat_2.append(list)
                heat_2.remove(heat_2[i])
            if remainder > 4:
                list = []
                for j in range(2,4):
                    list.append(assignments[i][j])
                heat_2.append(list)
                heat_2.remove(heat_2[i])

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

#'heat' will be all three heats in a single array
heat = []
heat.append(heat_1)
heat.append(heat_2)
heat.append(heat_3)


#writing to excel file
#labeling files properly based on which week and group
if group == garnet:
    if week == 1:
        workbook = xlsxwriter.Workbook(file_location.replace('section_assignments_wk1','benchmark_greener_wk1_garnet'))
    if week == 2:
        workbook = xlsxwriter.Workbook(file_location.replace('section_assignments_wk2','benchmark_greener_wk2_garnet'))
    if week == 3:
        workbook = xlsxwriter.Workbook(file_location.replace('section_assignments_wk3','benchmark_greener_wk3_garnet'))
if group == white:
    if week == 1:
        workbook = xlsxwriter.Workbook(file_location.replace('section_assignments_wk1','benchmark_greener_wk1_white'))
    if week == 2:
        workbook = xlsxwriter.Workbook(file_location.replace('section_assignments_wk2','benchmark_greener_wk2_white'))
    if week == 3:
        workbook = xlsxwriter.Workbook(file_location.replace('section_assignments_wk3','benchmark_greener_wk3_white'))

#first tab will be the the benchmarks grids (second will be heat sheet)
worksheet = workbook.add_worksheet('green')

#applying borders to the cells
format = workbook.add_format()
format.set_border(style=1)
format2 = workbook.add_format()
format2.set_border(style=1)
format2.set_bold()
format2.set_font_size(17)

#writing column headings
column_number = 0
headings = ['Name','Set','Str','1st 25/50','2nd 25/50','3rd 25/50','4th 25/50','Total Time']
lane_labels = ['Lane 1','Lane 2','Lane 3','Lane 4','Lane 5','Lane 6','Lane 7','Lane 8']
labels = [[],[],[],[],[],[],[],[]]
for i in range(0,len(lane_labels)):
    labels[i] = [lane_labels[i],'Set','Str','1st 25/50','2nd 25/50','3rd 25/50','4th 25/50','Total Time']

#writing column headings
for i in range(0,len(headings)):
    for j in range(0,8*len(lane),8):
        worksheet.write(0,(j+i),headings[i], format)

for i in range(0,len(lane)):
    worksheet.write(0,i*8,labels[i][0], format2)
        
#set columns to be wider
for i in range(0,64):
    if i%8 == 0:
        worksheet.set_column(i,i,16.25)
    else:
        worksheet.set_column(i,i,11)
mark_column = []
#for i in range(0,8):
#mark_column.append(1)
#mark_column.append(2)
for i in range(0,64):
    if i%8 == 1 or i%8 == 2:
        mark_column.append(i)
for i in mark_column:
    worksheet.set_column(i,i,4)
#set rows to be wider
if number_heats == 2:
    if len(heat[1][0]) == 3:
        for i in range(1,23):
            worksheet.set_row(i,21.75)
    if len(heat[1][0]) != 3:
        for i in range(1,19):
            worksheet.set_row(i,21.75)
if number_heats == 3:
    if len(heat[2][0]) == 2:
        for i in range(1,28):
            worksheet.set_row(i,21.75)
    if len(heat[2][0]) == 3:
        for i in range(1,32):
            worksheet.set_row(i,21.75)

#placing labels for each heat
heat_headers = []
for i in range(0,64):
    if i%8 == 0:
        heat_headers.append(i)
store = ['Heat 1','Heat 2','Heat 3']
for i in range(0,3):
    for j in heat_headers:
        if i == 0:
            worksheet.write(1,j,store[i])
        else:
            worksheet.write(i*10-(i-1),j,store[i])
   
#writing names into excel
#formatting borders, etc.
column_blanks = []
for i in range(0,64):
    if i%8 != 0:
        column_blanks.append(i)
row_blanks = []
if number_heats == 2:
    if len(heat[1][0]) != 3:
        row_blanks = [2,3,4,5,6,7,8,9,11,12,13,14,15,16,17,18]
    if len(heat[1][0]) == 3:
        row_blanks = [2,3,4,5,6,7,8,9,11,12,13,14,15,16,17,18,19,20,21,22]
if number_heats == 3:
    if len(heat[2][0]) == 2:
        row_blanks = [2,3,4,5,6,7,8,9,11,12,13,14,15,16,17,18,20,21,22,23,24,25,26,27]
    if len(heat[2][0]) == 3:
        row_blanks = [2,3,4,5,6,7,8,9,11,12,13,14,15,16,17,18,20,21,22,23,24,25,26,27,28,29,30,31]

#properly mapping assignments entries to excel sheet
if len(assignments[0]) == 7 and len(assignments[1]) == 6:
    assignments[0].remove('blank')
for i in range(0,len(assignments[0])):
    for k in range(0,len(lane)):
        if i == 0:
            worksheet.write_string(2, 8*k, assignments[k][i], format)
        if i == 1:
            worksheet.write_string(6, 8*k, assignments[k][i], format)
        if i == 2:
            worksheet.write_string(11, 8*k, assignments[k][i], format)
        if i == 3:
            worksheet.write_string(15, 8*k, assignments[k][i], format)
        if i == 4:
            if len(assignments[0]) == 5:
                worksheet.write_string(19, 8*k, assignments[k][i], format)
            else:
                worksheet.write_string(20, 8*k, assignments[k][i], format)
        if i == 5:
            worksheet.write_string(24, 8*k, assignments[k][i], format)
        if i == 6:
            worksheet.write_string(28, 8*k, assignments[k][i], format)
            
#inserting blank bordered cells for filling out with times
if len(assignments[0]) == 5:
    row_blanks.append(19)
for i in column_blanks:
    for j in row_blanks:
        worksheet.write_blank(j,i,'',format)
#if only 2 heats, deleting bordered cells from heat three rows
if number_heats == 2:
    if len(heat[1][0]) != 3:
        for i in range(0,64):
            for j in range(19,32):
                format = workbook.add_format()
                format.set_border(style=0)
                worksheet.write_blank(j,i,'',format)        
    else:
        for i in range(0,64):
            for j in range(23,32):
                format = workbook.add_format()
                format.set_border(style=0)
                worksheet.write_blank(j,i,'',format)

#properly mapping mark type to excel sheet

#marking the type of benchmark set
mark_type = ['100k','100s','200s','500s']
format = workbook.add_format()
format.set_border(style=1)
for i in range(0,len(mark_type)):
    for j in range(0,len(lane)):
        worksheet.write_string(i+2,j*8+1,mark_type[i],format)
        worksheet.write_string(i+6,j*8+1,mark_type[i],format)
        if number_heats == 2:
            if len(heat[1][0]) != 3:
                worksheet.write_string(i+11,j*8+1,mark_type[i],format)
                worksheet.write_string(i+15,j*8+1,mark_type[i],format)
            if len(heat[1][0]) == 3:
                worksheet.write_string(i+11,j*8+1,mark_type[i],format)
                worksheet.write_string(i+15,j*8+1,mark_type[i],format)
                worksheet.write_string(i+19,j*8+1,mark_type[i],format)
        if number_heats == 3:
            if len(heat[2][0]) != 3:
                worksheet.write_string(i+11,j*8+1,mark_type[i],format)
                worksheet.write_string(i+15,j*8+1,mark_type[i],format)                
                worksheet.write_string(i+20,j*8+1,mark_type[i],format)
                worksheet.write_string(i+24,j*8+1,mark_type[i],format)
            if len(heat[2][0]) == 3:
                worksheet.write_string(i+11,j*8+1,mark_type[i],format)
                worksheet.write_string(i+15,j*8+1,mark_type[i],format)
                worksheet.write_string(i+20,j*8+1,mark_type[i],format)
                worksheet.write_string(i+24,j*8+1,mark_type[i],format)
                worksheet.write_string(i+28,j*8+1,mark_type[i],format)

#adding in the header for the benchmark sets
worksheet.set_margins(bottom=0.5)
worksheet.set_margins(top=0.75)
if group == 1:
    header = '&CBCSC Benchmarks Garnet'
    worksheet.set_header(header)
if group == 2:
    header = '&CBCSC Benchmarks White'
    worksheet.set_header(header)

#creating formulas for the cells to add up total times
calc_rows = [2,3,4,6,7,8,11,12,13,15,16,17]
if number_heats == 2:
    if len(heat[1][0]) != 3:
        for i in []:
            calc_rows.append()
    if len(heat[1][0]) == 3:
        for i in [19,20,21]:
            calc_rows.append(i)
if number_heats == 3:
    if len(heat[2][0]) != 3:
        for i in [20,21,22,24,25,26]:
            calc_rows.append(i)
    if len(heat[2][0]) == 3:
        for i in [20,21,22,24,25,26,28,29,30]:
            calc_rows.append(i)
calc_columns = [6,13,20,27,34,41,48,55]
add = '=SUM'
numbers = [3,4,5,7,8,9,12,13,14,16,17,18]
if number_heats == 2:
    if len(heat[1][0]) == 3:
        for i in [20,21,22]:
            numbers.append(i)
if number_heats == 3:
    if len(heat[2][0]) == 3:
        for i in [21,22,23,25,26,27,29,30,31]:
            numbers.append(i)
    if len(heat[2][0]) != 3:
        for i in [21,22,23,25,26,27]:
            numbers.append(i)

calc_1 = [add+'(C'+str(i)+':F'+str(i)+')/86400' for i in numbers]
calc_2 = [add+'(J'+str(i)+':M'+str(i)+')/86400' for i in numbers]
calc_3 = [add+'(Q'+str(i)+':T'+str(i)+')/86400' for i in numbers]
calc_4 = [add+'(X'+str(i)+':AA'+str(i)+')/86400' for i in numbers]
calc_5 = [add+'(AE'+str(i)+':AH'+str(i)+')/86400' for i in numbers]
calc_6 = [add+'(AL'+str(i)+':AO'+str(i)+')/86400' for i in numbers]
calc_7 = [add+'(AS'+str(i)+':AV'+str(i)+')/86400' for i in numbers]
calc_8 = [add+'(AZ'+str(i)+':BC'+str(i)+')/86400' for i in numbers]

calc = [calc_1,calc_2,calc_3,calc_4,calc_5,calc_6,calc_7,calc_8]

#the first index for calc gets you to calc_1/2/3/4_all...etc., that is gets you all the correct summation cells for each given lane (1-8)
#the second index MUST BE [0] idk why there's an extra layer in there but oh well...
#len(calc[0][0]) should = p should = len(assignments[0])
#the third index gets you to the person in each lane...example: calc[0][0][1] would give you ['=SUM(C7:F7)', '=SUM(C8:F8)', '=SUM(C9:F9)'], which would be the formulas for the second person in the first lane. THIS INDEX SHOULD BE REGULATED BASED ON HOW MANY PEOPLE ARE IN THE HEAT...i.e. set up conditionals (set up as the p values below)
#the fourth and final index gets you to the individual set...this should loop over three

if number_heats == 2:   #this is where I set up the p conditionals mentioned above
    if len(heat[1][0]) != 3:
        p = 4
    if len(heat[1][0]) == 3:
        p = 5
if number_heats == 3:
    if len(heat[2][0]) != 3:
        p = 6
    if len(heat[2][0]) == 3:
        p = 7
#below sets up the cells in which the total times will come through after fomula
letters = ['H','P','X','AF','AN','AV','BD','BL']
calc_cells_all = [letters[i]+str(numbers[j]) for i in range(len(letters)) for j in range(len(numbers))]
calc_cells_H = []
calc_cells_P = []
calc_cells_X = []
calc_cells_AF = []
calc_cells_AN = []
calc_cells_AV = []
calc_cells_BD = []
calc_cells_BL = []
for i in range(len(calc_cells_all)):
    if calc_cells_all[i][0] == 'H':
        calc_cells_H.append(calc_cells_all[i])
    if calc_cells_all[i][0] == 'P':
        calc_cells_P.append(calc_cells_all[i])
    if calc_cells_all[i][0] == 'X':
        calc_cells_X.append(calc_cells_all[i])
    if calc_cells_all[i][1] == 'F':
        calc_cells_AF.append(calc_cells_all[i])
    if calc_cells_all[i][1] == 'N':
        calc_cells_AN.append(calc_cells_all[i])
    if calc_cells_all[i][1] == 'V':
        calc_cells_AV.append(calc_cells_all[i])
    if calc_cells_all[i][1] == 'D':
        calc_cells_BD.append(calc_cells_all[i])
    if calc_cells_all[i][1] == 'L':
        calc_cells_BL.append(calc_cells_all[i])
#the calc_cells array is an array that holds all the cells to be filled out by formulas, separate arrays per lane...so it's an array of 8 arrays
calc_cells = [calc_cells_H,calc_cells_P,calc_cells_X,calc_cells_AF,calc_cells_AN,calc_cells_AV,calc_cells_BD,calc_cells_BL]

#set number formatting for the calc_cells
format1 = workbook.add_format()
format1.set_num_format('m:ss.0')
format1.set_border(style=1)

#worksheet writing below
for i in range(0,len(lane)): #ranges over lanes...8 iterates
    for j in range(0,len(calc_cells[0])):  #ranges pvr nmber calc_cells/lane (3xkids/lane)
        worksheet.write_formula(calc_cells[i][j],calc[i][j],format1)

#greying out the cells for the 500 benchmark that are not necessary (25s/50s splits)
grey = [2,3,4,5,6]
grey_columns = []
for i in range(0,len(lane)):
    for j in grey:
        grey_columns.append(i*8+j)
grey_rows = [5,9,14,18]
if number_heats == 2:
    if len(heat[1][0]) != 3:
        grey_rows = [5,9,14,18]
    if len(heat[1][0]) == 3:
        grey_rows.append(22)
if number_heats == 3:
    if len(heat[2][0]) != 3:
        for i in [23,27]:
            grey_rows.append(i)
    if len(heat[2][0]) == 3:
        for i in [23,27,31]:
            grey_rows.append(i)
format = workbook.add_format()
format.set_pattern(1)
format.set_bg_color('gray')
format.set_top()
format.set_bottom()
for r in grey_rows:
    for c in grey_columns:
        worksheet.write_string(r,c,'',format)

#makes the 500 total time in the right format
for r in grey_rows:
    worksheet.write_blank(r,7,'',format1)


#spits out a heat sheet for each group
#second tab - this one is for the heat sheet
worksheet = workbook.add_worksheet('heatsheet')
format = workbook.add_format()
format.set_bold()
format.set_font_size(12)
format.set_border(style=1)
format3 = workbook.add_format()
format3.set_border(style=1)
format3.set_font_size(13)
format2.set_border(style=1)
format2.set_font_size(17)
format2.set_bold()
heat_labels = ['Heat 1','Heat 2','Heat 3']
lanes = ['Lane 1:','Lane 2:','Lane 3:','Lane 4:','Lane 5:','Lane 6:','Lane 7:','Lane 8:']
heat_rows =  [0,10,20]
lane_rows = [[1,2,3,4,5,6,7,8],[11,12,13,14,15,16,17,18],[21,22,23,24,25,26,27,28]]
for i in range(number_heats):   #sets heat labels
    worksheet.write_string(heat_rows[i],0,heat_labels[i],format2)
for i in range(number_heats):   #sets lane labels
    for j in range(0,len(lanes)):
        worksheet.write_string(lane_rows[i][j],0,lanes[j],format)
for i in range(number_heats):
    for j in range(len(lanes)):
        worksheet.write_string(lane_rows[i][j],1,heat[i][j][0],format3)
        worksheet.write_string(lane_rows[i][j],2,heat[i][j][1],format3)
        if len(heat[1][0]) == 3:
            worksheet.write_string(lane_rows[1][j],3,heat[1][j][2],format3)
        if number_heats == 3:
            if len(heat[2][0]) == 3:
                worksheet.write_string(lane_rows[2][j],3,heat[2][j][2],format3)

#set column width
for i in [1,2,3]:
    worksheet.set_column(i,i,22)

#adding in the header for the heatsheets
worksheet.set_margins(bottom=0.5)
worksheet.set_margins(top=0.75)
if group == 1:
    header = '&CBCSC Benchmark Heat Sheets Garnet Group Week ' + str(week)
    worksheet.set_header(header)
if group == 2:
    header = '&CBCSC Benchmark Heat Sheets White Group Week ' + str(week)
    worksheet.set_header(header)

workbook.close()

if group == garnet:
    print('garnet')
if group == white:
    print('white')
print(week)
