from PyQt5 import QtCore, QtGui
from PyQt5.QtWidgets import QApplication, QPushButton, QDialog, QLineEdit, QLabel, QHBoxLayout, QVBoxLayout, QSpacerItem, QMainWindow, QTabWidget, QWidget, QMessageBox

import pyqtgraph as pg

import os, sys, glob, random, datetime, re

import xlrd

import sqlite3





pg.setConfigOptions(antialias=True)

scriptPath = os.path.dirname(os.path.realpath(__file__))

XLS_START_ROW = 8

class MainWindow(QDialog):
        
    def __init__(self, parent):
        super().__init__()

        self.parent = parent

        ## Widgets
        self.setWindowTitle("Fidelity Funds")        
        self.setGeometry(300,300,1000,700)
        
        self.btn01 = QPushButton("Load Files", self)
        self.btn01.pressed.connect(self.parent.loadNewRecords)

        self.btn02 = QPushButton("Clear All", self)
        self.btn02.pressed.connect(self.parent.clearDatabase)
        
        self.progressLabel = QLabel()
     
        self.pW_funds_separated = pg.PlotWidget()
        self.pW_funds_combined = pg.PlotWidget()

        self.tabs = QTabWidget()
        self.tabs.setTabsClosable(True)

        self.tabs.addTab(self.pW_funds_separated,"Separated")
        self.tabs.addTab(self.pW_funds_combined,"Combined")
        self.tabs.currentChanged.connect(self.PrintTabIndex)

        ## Layout
        buttonsHBox = QHBoxLayout()
        buttonsHBox.addWidget(self.btn01)
        buttonsHBox.addWidget(self.btn02)
        buttonsHBox.addStretch()
        buttonsHBox.addWidget(self.progressLabel)
        
        layout = QVBoxLayout()
        layout.addLayout(buttonsHBox)
        layout.addWidget(self.tabs)
        
        self.setLayout(layout)
        self.show()

    def PrintTabIndex(self):
        print("Current Tab: " + str(self.tabs.currentIndex()))


class FundRecord():
    
    def __init__(self, record):
        self.holdingName = str(record[1])
        
        re_date = re.search(r'(^\d\d\/\d\d\/\d\d\d\d$)', record[4])
        if re_date != None:
            stringDate = re_date.group(1)
            self.date = (datetime.datetime.strptime(stringDate, "%d/%m/%Y").date() - datetime.date(2000,1,1)).days
        else:
            self.date = 0

        re_value = re.search(r'^([£]?)([\S]+)$', record[7])
        if re_value != None:
            self.value = float(re_value.group(2).replace(',',''))
        else:
            self.value = 0

        #self.printInfo()

    def printInfo(self):
        print('Holding Name: ' + self.holdingName)
        print('Valued £' + str(self.value) + ' on ' + str(self.date))

class FundReport():

    def __init__(self, fileName, parent):
        self.parent = parent
        self.fundRecords = []

        self.loadXls(fileName)

    def loadXls(self, fileName):
        filePath = os.path.join(scriptPath,fileName)
        book = xlrd.open_workbook(filePath)
        sheet = book.sheet_by_index(0)
        
        for i in range(XLS_START_ROW, sheet.nrows-2):
            row = sheet.row(i)
            row_as_strings = [None] * len(row)
            for index, item in enumerate(row):
                row_as_strings[index] = str(item.value)

            self.fundRecords.append(FundRecord(row_as_strings))

    def saveToDb(self):

        cur = self.parent.conn.cursor()
        
        for fundRecord in self.fundRecords:
            cur.execute('SELECT COUNT(*) FROM fundRecords WHERE holding = ? AND date = ?', (fundRecord.holdingName, fundRecord.date, ))
            duplicateCount = cur.fetchone()[0]
            if duplicateCount == 0:
                cur.execute('INSERT INTO fundRecords (holding, date, value) VALUES (?, ?, ?);', (fundRecord.holdingName, fundRecord.date, fundRecord.value) )
            #else:
                #print('Record of ' + fundRecord.holdingName + ' on ' + str(fundRecord.date) + ' already in database.')
            self.parent.conn.commit()


class App():

    class FundHistory():
    
        def __init(self):
            print('a')

    def __init__(self, qApp):
        
        self.qApp = qApp
        self.mainWindow = MainWindow(self)
        self.connectToDb()
        self.plotFunds()

    def connectToDb(self):
        try:
            self.conn = sqlite3.connect(os.path.join(scriptPath,'sqliteDb'))

        except:
            def msgButtonPressed(btn):
                if btn.text() == 'Retry':
                    self.connectToDb()
                else:
                    sys.exit()

            msg = QMessageBox(text="Could not connect to DB", windowTitle="Error", icon=QMessageBox.Warning)
            msg.setStandardButtons(QMessageBox.Retry | QMessageBox.Close)
            msg.buttonClicked.connect(msgButtonPressed)
            msg.exec_()
            return None

    def loadNewRecords(self):
        for fname in glob.glob(os.path.join(scriptPath,"data/*.xls")):
            self.mainWindow.progressLabel.setText(fname)
            self.qApp.processEvents()
            report = FundReport(fname, self)
            report.saveToDb()
            del report
            self.mainWindow.progressLabel.clear()
            
        self.plotFunds()
            
    def plotFunds(self):
        cur = self.conn.cursor()
        cur.execute('SELECT DISTINCT holding FROM fundRecords EXCEPT SELECT holding FROM fundRecords WHERE holding=?', ('Cash',))
        holdingNames = [i[0] for i in cur.fetchall()]
        
        self.mainWindow.pW_funds_separated.clear()
        
        for holdingName in holdingNames:
            cur.execute('SELECT date, value FROM fundRecords WHERE holding = ? ORDER BY date', (holdingName, ))
            result = cur.fetchall()
            dates  = [i[0] for i in result]
            values = [i[1] for i in result]
            self.mainWindow.pW_funds_separated.plot(x=dates, y=values)
        
    def clearDatabase(self):
        cur = self.conn.cursor()
        cur.execute('DELETE FROM fundRecords')
        
        self.plotFunds()
        

if __name__ == '__main__':
    
    qApp = QApplication(sys.argv)
    testApp = App(qApp)
    sys.exit(qApp.exec_())
    
