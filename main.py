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
dspldir = pathos + '\\WorkPlace\\Displays'
if not os.path.exists(dspldir + '\\Result'):
    os.mkdir('WorkPlace\\Displays\\Result')

for xls in os.listdir(xlsdir):
    i = 1
    if xls.split('.')[-1] == 'xls' and xls[:4] != 'new_' and xls[:4] != 'rem_':
        wb_new = xlwt.Workbook()
        ws_new = wb_new.add_sheet(xls.split('.')[1])
        wb_rem = xlwt.Workbook()
        ws_rem = wb_rem.add_sheet(xls.split('.')[1])
        print u'***********************',xls
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

for xml in os.listdir(dspldir):
    if xml.split('.')[-1] == 'xml':
        print u'***********************', xml
        fu = open(dspldir + '\\' + xml, 'r')
        donewfileflag = False
        Lines = []
        for line in fu:
            if '<itemId>' not in line:
                #if '<value>' in line and '</value>' in line:
                #    oldvalue = line.strip()[7:-8]
                #    if ItemNewOld.get(oldvalue):
                #        line = line.replace(oldvalue, ItemNewOld[oldvalue])
                #if '<itemName>' in line and '</itemName>' in line:
                #    oldvalue = line.strip()[10:-11]
                #    if ItemNewOld.get(oldvalue):
                #        line = line.replace(oldvalue, ItemNewOld[oldvalue])
                line = line.decode('utf8')
                for oldname in old_names:
                    #print type(oldname), type (line)
                    #print oldname, line
                    if oldname in line:
                        line = line.replace(oldname, matchtable[oldname])
                        donewfileflag = True
                Lines.append(line)
        fu.close()
        if donewfileflag:
            fr = open('WorkPlace\\Displays\\Result\\' + xml, 'w')
            for line in Lines:
                fr.write(line.encode('utf8'))
            fr.close()

print u'Обработка закончена'
