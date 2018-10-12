#!/usr/bin/env python
# -*- coding: utf-8 -*-
import xlrd
import xlwt
import os
#import re

def writeROW(i, data, ws):
    for cell in range(len(data)):        
        ws.write(i, cell, data[cell])

pathos = os.getcwd()

matchtablexls = pathos + '\\WorkPlace\\MatchTable.xls'
rb = xlrd.open_workbook(matchtablexls)
sheet = rb.sheet_by_index(0)
matchtable = {}
old_names = []
for row in range(sheet.nrows):
    old = sheet.cell_value(row, 0)
    new = sheet.cell_value(row, 1)
    old_names.append(old)
    matchtable[old] = new
#print matchtable
#print len(matchtable)

xlsdir = pathos + '\\WorkPlace\\xls'
i = 1
for xls in os.listdir(xlsdir):
    if xls.split('.')[-1] == 'xls' and xls[:4] != 'new_':
        wb_new = xlwt.Workbook()
        ws_new = wb_new.add_sheet(xls.split('.')[1])
        wb_rem = xlwt.Workbook()
        ws_rem = wb_rem.add_sheet(xls.split('.')[1])
        print '***********************',xls   
        rb = xlrd.open_workbook(xlsdir + '\\' + xls)
        sheet = rb.sheet_by_index(0)
        ncols = sheet.ncols
        rawrow = []
        for col in range(ncols):
            cell = sheet.cell_value(0, col)
            rawrow.append(cell)
        writeROW(0, rawrow, ws_new)
        writeROW(0, rawrow, ws_rem)

        for row in range(1, sheet.nrows):
            addrowflag = False
            rawrow_new = []
            rawrow_rem = []
            for col in range(ncols):
                cell_rem = sheet.cell_value(row, col)
                cell_new = cell_rem
                if cell_new in old_names:
                    cell_new = matchtable[cell_new]
                    addrowflag = True
                rawrow_rem.append(cell_rem)
                rawrow_new.append(cell_new)
            if addrowflag:
                writeROW(i, rawrow_new, ws_new)
                writeROW(i, rawrow_rem, ws_rem)
                i += 1

        wb_new.save(xlsdir + '\\new_' + xls)
        wb_rem.save(xlsdir + '\\rem_' + xls)
        

print u'Обработка закончена'
