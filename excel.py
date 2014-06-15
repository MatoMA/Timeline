#!/usr/bin/env python
import webbrowser, os, fileinput
from xlrd import open_workbook, xldate_as_tuple
from Tkinter import *

funds = dict()

#--------Data structure---------------------
class Event:
    def __init__(self, name, comment, date):
        self.name = name
        self.comment = comment
        self.date = date

    def toString(self):
        s = ""
        s = s + '{ "startDate":"' + \
                str(self.date[0]) + ',' + str(self.date[1]) + ',' + str(self.date[2]) + \
                '", ' + \
                '"endself.date":"' + \
                str(self.date[0]) + ',' + str(self.date[1]) + ',' + str(self.date[2]) + \
                '", ' + \
                '"headline":"' + \
                self.name + \
                '", ' + \
                '"text":"' + \
                '<p>' + self.comment + r'</p>' + \
                '", ' + '},'

        return s

class Fund:
    def __init__(self, name):
        self.name = name
        self.events = []

    def addEvent(self, event):
        self.events.append(event)

    def printSelf(self):
        print self.name
        print "----"
        for event in self.events:
            print "Event Type: ", event.name
            print "Event Comment: ", event.comment
            print "TDC Date: ", event.date
            print "---------------"

    def showTimeline(self):
        template = open('template.html', 'r')
        if os.path.isfile('index.html'):
            os.remove('index.html')
        indexFile = open('index.html', 'w')

        for line in template:
            if line.find("FundName"):
                line = line.replace("FundName", self.name)
            if line.find("EventDates"):
                eventString = ""
                for event in self.events:
                    eventString = eventString + event.toString()
                line = line.replace("EventDates", eventString)
            indexFile.write(line.encode('utf8'))
        template.close()
        indexFile.close()

        filepath = os.path.realpath('index.html')
        webbrowser.open('file://'+filepath)

#-------------------------------------------

#---------Get data from excel---------------
def extractData(filename):
    wb = open_workbook(filename)
    sheets = wb.sheets()
    sheet = sheets[0]

    umbrellaCol = 2
    fundCol = umbrellaCol + 1
    eventCol = fundCol + 1
    commentCol = eventCol + 1
    TDCDateCol = commentCol + 1

    for row in range(sheet.nrows)[1:]:
        firstColValue = sheet.cell(row, 0).value
        if firstColValue == "Local" or firstColValue == "Global":
            umbrellaName = sheet.cell(row, umbrellaCol).value
            fundName = sheet.cell(row, fundCol).value
            eventName = sheet.cell(row, eventCol).value
            comment = sheet.cell(row, commentCol).value
            date = sheet.cell(row, TDCDateCol).value

            if not isinstance(date, basestring):
                date = xldate_as_tuple(date, 0)
                event = Event(eventName, comment, date)

                if umbrellaName != '-' and umbrellaName != '':
                    if umbrellaName not in funds.keys():
                        funds[umbrellaName] = Fund(umbrellaName)
                    funds[umbrellaName].addEvent(event)

                if fundName != '-' and fundName != '':
                    if fundName not in funds.keys():
                        funds[fundName] = Fund(fundName)
                    funds[fundName].addEvent(event)

#-------------------------------------------

#----------------GUI------------------------
def lb_click_callback(event):
    index = event.widget.curselection()
    fundName = event.widget.get(index)
    funds[fundName].showTimeline()

def tk():
    root = Tk()

    #Listbox configuration
    lb = Listbox(root, width=100, height=50, selectmode=SINGLE)
    i = 1
    sortedFundNames = sorted(funds.keys())
    for fund in sortedFundNames:
        lb.insert(i, fund)
        i = i + 1
    lb.bind("<ButtonRelease-1>", lb_click_callback)
    lb.pack()

    root.mainloop()
#-------------------------------------------

def main():
    filename = 'samples.xlsx'
    funds = extractData(filename)
    tk()

if __name__ == '__main__':
    main()

