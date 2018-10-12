#!/usr/bin/env python
# -*- coding: utf-8 -*-
import xlrd
import xlwt
import os
#from math import ceil
#from xlutils.copy import copy
#import re

def writeROW(i, data):    
    for cell in range(len(data)):        
        ws.write(i, cell, data[cell])

f01 = [u'MER-0' + str(x) for x in range(10,30)]
f02 = [u'MER-' + str(x) for x in range(400,456)] 
filter01 = [u'LER-014'] + f01 + f02 + [u'MER-501']
print filter01
filter02 = [u'UNGG']

pathos = os.getcwd()
xlsdir = pathos + '\\xls'
wb = xlwt.Workbook()
ws = wb.add_sheet(u'Result')
i  = 1
for xls in os.listdir(xlsdir):    
    if xls[-4:] == '.xls':
        print '***********************',xls   
        rb = xlrd.open_workbook(xlsdir + '\\' + xls)
        sheet = rb.sheet_by_index(0)
        count = 0
        ncols = sheet.ncols
        for row in range(sheet.nrows):
            count += 1
            #if count < 5416 or count>5430:
            #    continue
            rawrow = []
            for col in range(ncols):
                cell = sheet.cell_value(row, col)
                rawrow.append(cell)
            #print rawrow
            if (unicode(rawrow[14].strip()) in filter02) and (unicode(rawrow[30].strip()) in filter01):
                writeROW(i, rawrow)
                i+=1
            #else:
            #    print rawrow[14], rawrow[30]
            #    print rawrow[14].strip() in filter02, rawrow[30] in filter01                
            if count%100==0:
                print count
        
wb.save('result.xls')
print u'Обработка закончена'
