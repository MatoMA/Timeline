#!/usr/bin/env python
from xlrd import open_workbook

wb = open_workbook('samples.xlsx')

sheets = wb.sheets()

for s in sheets:
    print 'Sheet:', s.name

sheet1 = sheets[0]
for row in range(sheet1.nrows):
    for col in range(sheet1.ncols):
        print sheet1.cell(row, col).value
    print

    #for row in range(s.nrows):
        #values = []
        #for col in range(s.ncols):
            #values.append(s.cell(row, col).value)
            #print ',', values
        #print
