# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'PyQt_KML.ui'
#
# Created: Tue Dec 10 16:30:04 2013
#      by: PyQt4 UI code generator 4.10.3
#
# WARNING! All changes made in this file will be lost!
import sip
sip.setapi('QVariant', 2)

from PyQt4 import QtCore, QtGui
from functools import partial

try:
    _fromUtf8 = QtCore.QString.fromUtf8
except AttributeError:
    def _fromUtf8(s):
        return s

try:
    _encoding = QtGui.QApplication.UnicodeUTF8
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig, _encoding)
except AttributeError:
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig)
import difflib
import sys, os
sys.path.append(os.path.abspath(os.curdir))
from RFDB import RFDataBase
from RFKml import RFKml
import coloreditorfactory


tableData = ['Index','eNodeB Id','Latitude','Longitude','Sector Id','Cell Id','EARFCN','Azimuth','Type','PCI']

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName(_fromUtf8("MainWindow"))
        MainWindow.resize(600, 505)
        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Minimum, QtGui.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(MainWindow.sizePolicy().hasHeightForWidth())
        
        MainWindow.setSizePolicy(sizePolicy)
        MainWindow.setMinimumSize(QtCore.QSize(600, 450))
        MainWindow.setAutoFillBackground(True)
        self.centralwidget = QtGui.QWidget(MainWindow)
        self.centralwidget.setObjectName(_fromUtf8("centralwidget"))

        self.tabWidget = QtGui.QTabWidget(self.centralwidget)
        #self.tabWidget.setEnabled(True)
        self.tabWidget.setGeometry(QtCore.QRect(0, 0, 600, 420))


        #initial Object
        self.RFDB = RFDataBase()
        self.rf_kml = RFKml()

        self.tabWidget.addTab(self.tab_loading(), _fromUtf8(""))
        self.tab_setting()
        self.tab_calculating()

        self.tabWidget.currentChanged.connect(partial(self.Option_Analysis,2))
        
        MainWindow.setCentralWidget(self.centralwidget)        
        self.menubar = QtGui.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 600, 18))
        self.menubar.setObjectName(_fromUtf8("menubar"))
        self.menuFile = QtGui.QMenu(self.menubar)
        self.menuFile.setObjectName(_fromUtf8("menuFile"))
        MainWindow.setMenuBar(self.menubar)
        
        #self.statusbar = QtGui.QStatusBar(MainWindow)
        #self.statusbar.setObjectName(_fromUtf8("statusbar"))
        #MainWindow.setStatusBar(self.statusbar)
        #self.actionOpen = QtGui.QAction(MainWindow)
        #self.actionOpen.setObjectName(_fromUtf8("actionOpen"))
        #self.menuFile.addAction(self.actionOpen)
        #self.menubar.addAction(self.menuFile.menuAction())

        self.retranslateUi(MainWindow)
        self.tabWidget.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(_translate("MainWindow", "Excel2KMZ", None))
        self.tabWidget.setTabText(0, _translate("MainWindow", "LoadExcel", None))
        
        
        self.qpb_add.setText(_translate("MainWindow", "Add", None))
        self.tabWidget.setTabText(1, _translate("MainWindow", "Setting", None))
        self.menuFile.setTitle(_translate("MainWindow", "File", None))
        #self.actionOpen.setText(_translate("MainWindow", "Open", None))
        self.qpb_calc.setText(_translate("MainWindow", "Calc", None))
        self.tabWidget.setTabText(2, _translate("MainWindow", "Calc", None))
        #testing
        
        #D:/RNP&RNO/ENGINEER/Parameters/RF Parameter/20131206/RF_MasterList20131206.xls
       
    def tab_loading(self):
        tab_load = QtGui.QWidget()
        tab_load.setObjectName(_fromUtf8("tab_load"))
        #setting for tab_load
        str_file = QtGui.QLineEdit(tab_load)
        str_file.setGeometry(QtCore.QRect(10, 10, 300, 20))
        str_file.setReadOnly(True)
        #str_file.setObjectName(_fromUtf8("str_file"))   

        cob_sheet = QtGui.QComboBox(tab_load)
        cob_sheet.setGeometry(QtCore.QRect(325, 10, 75, 20))
        #cob_sheet.currentIndexChanged.connect(self.updateTable)    
        
        qpb_browse = QtGui.QPushButton(tab_load)
        qpb_browse.setGeometry(QtCore.QRect(415, 10, 75, 20))
        qpb_browse.clicked.connect(partial(self.browseFile, tab_load, str_file, cob_sheet))
        qpb_browse.setText(_translate("MainWindow", "Browse", None))
        #qpb_browse.setObjectName(_fromUtf8("qpb_browse"))

        excelMatchtable = QtGui.QTableWidget(len(tableData),2, tab_load)
        excelMatchtable.setHorizontalHeaderLabels(["Keywords", "Column"])
        excelMatchtable.verticalHeader().setVisible(False)
        excelMatchtable.setGeometry(QtCore.QRect(10, 40, 561, 311))
        excelMatchtable.setEditTriggers(QtGui.QAbstractItemView.NoEditTriggers)
        #excelMatchtable.setObjectName(_fromUtf8("excelMatchtable"))
        #self.comboTittle = QtGui.QComboBox()
        for i, keywords in enumerate(tableData):
            kwItem = QtGui.QTableWidgetItem(keywords)
            vlItem = QtGui.QTableWidgetItem('...')
            excelMatchtable.setItem(i,0,kwItem)
            excelMatchtable.setItem(i,1,vlItem)
        
        excelMatchtable.resizeColumnToContents(1)
        excelMatchtable.horizontalHeader().setStretchLastSection(True)

        qpb_analysis = QtGui.QPushButton(tab_load)
        qpb_analysis.setGeometry(QtCore.QRect(505, 10, 75, 20))
        qpb_analysis.clicked.connect(partial(self.updateTable, cob_sheet, excelMatchtable))
        #qpb_analysis.setObjectName(_fromUtf8("qpb_analysis"))
        qpb_analysis.setText(_translate("MainWindow", "Analysis", None))


        return tab_load
    def Option_Analysis(self, i):
        if self.tabWidget.currentIndex() == 1:
            self.cell_item = dict()
            #change to tab_setting, do KPI statistics now
            self.cell_item["EARFCN"] = self.RFDB.DBStatistics("EARFCN")
            self.cell_item["Type"] = self.RFDB.DBStatistics("Type")
            self.cell_item["PCI"] = ['0', '1', '2']

            self.kml_item = dict()
            self.kml_item["Shape"] = ("Sector", "Circle")
            self.kml_item["Line Color"] = list()
            self.kml_item["Fill Color"] = list()

            self.rules = dict()

            self.setTable.clear()

            self.rules[("EARFCN", "Fill Color")] = dict().fromkeys(self.cell_item["EARFCN"], "blue")
            self.rules[("PCI", "Line Color")] = dict().fromkeys(self.cell_item["PCI"], "blue")
            self.rules[("Type", "Shape")] = dict(zip(self.cell_item["Type"],self.kml_item["Shape"]))


            for i, rule in enumerate(self.rules):
                #cell_type = rule[0]
                #kml_type = rule[1]
                print rule
                kml_Item = QtGui.QTableWidgetItem(rule[1])
                rf_Item = QtGui.QTableWidgetItem(rule[0])

                pub_options = QtGui.QPushButton('...')
                pub_options.clicked.connect(partial(self.tab_set_options, rule, self.rules[rule]))
                pub_remove = QtGui.QPushButton('Remove')
                pub_remove.clicked.connect(partial(self.tab_set_remove,i))

                self.setTable.setItem(i,0,rf_Item)
                #self.setTable.setColumnWidth(0,205)
                self.setTable.setItem(i,1,kml_Item)
                #self.setTable.setColumnWidth(1,205)
                self.setTable.setCellWidget(i,2,pub_options)
                #self.setTable.setColumnWidth(2,20)
                self.setTable.setCellWidget(i,3,pub_remove)
                #self.setTable.setColumnWidth(3,50)
            self.setTable.resizeColumnToContents(1)

    def updateTable(self, cob_sheet, excelMatchtable):
        #self.RFDB.readSheet
        if self.RFDB.readSheet(cob_sheet.currentIndex()):
            strTitle = self.RFDB.str_title
            self.map_title = dict()
        else:
            print "Read file Failed."
        for i_row in range(excelMatchtable.rowCount()):
            str_match = str(excelMatchtable.item(i_row, 0).text())
            comboTittle = QtGui.QComboBox()
            comboTittle.addItems(strTitle)
            try:
                index_match = strTitle.index(difflib.get_close_matches(str_match,strTitle)[0])
            except IndexError:
                index_match = -1            
            if index_match!=-1:
                comboTittle.setCurrentIndex(index_match)
                excelMatchtable.item(i_row, 0).setBackgroundColor(QtGui.QColor(0,255,13))
                self.map_title[str_match] = strTitle[index_match]
            excelMatchtable.setCellWidget(i_row, 1, comboTittle)
            comboTittle.currentIndexChanged.connect(partial(self.updateOptionTable, self.map_title, str_match, comboTittle))
                
        self.tabWidget.setTabEnabled(1,True)

    
    def browseFile(self, tab_load, str_file, cob_sheet):
        fileName = QtGui.QFileDialog.getOpenFileName(tab_load,"Open File",'.',"Excel File (*.xlsx *.xls);;All File (*)")    
        #if RFDB.readSheet:
        str_shnames = False
        if fileName:
            str_shnames = self.RFDB.readFile(fileName)
        if str_shnames:
            str_file.setText(fileName)
            self.tabWidget.setTabEnabled(1,False)
            #self.excelFile = excelAnalysis(str(str_file.text())) 
            cob_sheet.clear()     
            cob_sheet.addItems(str_shnames)
            #cob_sheet.setCurrentIndex(-1)
            #self.updateTable(strTitle)
        else:
            msgBox = QtGui.QMessageBox()
            msgBox.setWindowTitle("Warning")
            msgBox.setText("Open Excel Failed!")
            msgBox.exec_()

    def tab_setting(self):
        #setting for tab_set        
        self.tab_set = QtGui.QWidget()
        self.kml_rules = dict()
        label = QtGui.QLabel(self.tab_set)
        label.setGeometry(QtCore.QRect(20, 20, 150, 20))
        label.setText(_translate("MainWindow", "Distinguish Cell", None))
        
        qcb_type = QtGui.QComboBox(self.tab_set)
        qcb_type.setGeometry(QtCore.QRect(380, 20, 75, 20))

        label_2 = QtGui.QLabel(self.tab_set)
        label_2.setGeometry(QtCore.QRect(220, 20, 150, 20))
        label_2.setText(_translate("MainWindow", "according to different", None))
        
        qcb_value = QtGui.QComboBox(self.tab_set)
        qcb_value.setGeometry(QtCore.QRect(130, 20, 75, 20))
        
        self.qpb_add = QtGui.QPushButton(self.tab_set)
        self.qpb_add.setGeometry(QtCore.QRect(500, 20, 75, 20))

        KML_options = ['Shape','Line Color', 'Shape Color']
        RF_type = ['Type','Mod(PCI,3)','EARFCN']
        qcb_type.addItems(KML_options)
        qcb_value.addItems(RF_type)

        self.setTable = QtGui.QTableWidget(3,4,self.tab_set)
        self.setTable.horizontalHeader().setVisible(False)
        self.setTable.verticalHeader().setVisible(True)
        self.setTable.setGeometry(QtCore.QRect(40, 50, 500, 200))
        
        self.tabWidget.addTab(self.tab_set, _fromUtf8(""))
        #self.tabWidget.setTabEnabled(1,False)

    def tab_set_options(self, rule, rule_value):
        kml_type = rule[1]
        #print kml_type
        if kml_type.find("Color") != -1:
            self.set_options('Color', rule_value)
        elif kml_type.find('Shape') != -1:
            self.set_options('Shape', rule_value)
        else:
            pass

    def set_options(self, set_type, rule_value):
        if set_type == 'Color':            
            color_dialog = QtGui.QDialog()
            color_dialog.setModal(True)
            color_factory = coloreditorfactory.Window(color_dialog,rule_value)
            color_dialog.resize(275,270)
            color_dialog.setWindowTitle(set_type)
            color_factory.show()
            color_dialog.exec_()
        elif set_type == 'Shape':
            shape_dialog = QtGui.QDialog()
            shape_dialog.setModal(True)
            shape_dialog.resize(275,270)
            shape_dialog.setWindowTitle(set_type)

            shape_table = QtGui.QTableWidget(len(rule_value),2,shape_dialog)
            shape_table.horizontalHeader().setVisible(False)
            shape_table.verticalHeader().setVisible(False)
            shape_table.resize(270, 270)
            shape_table.setEditTriggers(QtGui.QAbstractItemView.NoEditTriggers)

            for i, rf_value in enumerate(rule_value):
                nameItem = QtGui.QTableWidgetItem(rf_value)
                #shapeItem = QtGui.QTableWidgetItem()
                comboShape = QtGui.QComboBox()
                comboShape.addItems(rule_value.values())
                comboShape.setCurrentIndex(i)
                shape_table.setItem(i, 0, nameItem)
                shape_table.setCellWidget(i, 1, comboShape)

                comboShape.currentIndexChanged.connect(partial(self.updateOptionTable, rule_value, rf_value, comboShape))

            #shape_table.resizeColumnToContents(True)
            #shape_table.horizontalHeader().setStretchLastSection(True)

            #shape_dialog.resize(275,270)
            shape_dialog.setWindowTitle(set_type)
            shape_dialog.exec_()


    def updateOptionTable(self, rule_value, rf_value, comboShape):
        new_value = {rf_value:str(comboShape.currentText())}
        rule_value.update(new_value)



    def tab_set_remove(self,i):
        #self.setTable.removeRow(i)
        pass
    def tab_set_add(self):
        pass

    def tab_calculating(self):
        self.tab_calc = QtGui.QWidget()
        self.tab_calc.setObjectName(_fromUtf8("tab_calc"))
        
        self.qpb_calc = QtGui.QPushButton(self.tab_calc)
        self.qpb_calc.setGeometry(QtCore.QRect(500,20,75,25))
        self.qpb_calc.setObjectName(_fromUtf8("pushButton"))
        
        self.textBrowser = QtGui.QTextBrowser(self.tab_calc)
        self.textBrowser.setGeometry(QtCore.QRect(30, 20, 450, 300))
        self.textBrowser.setObjectName(_fromUtf8("textBrowser"))
        self.tabWidget.addTab(self.tab_calc, _fromUtf8(""))

        

        self.qpb_calc.clicked.connect(self.calc)


    def calc(self):
        #try:
        
        self.rf_kml.createRules(self.rules)
        self.rf_kml.createCells(self.RFDB.DB, self.map_title)
            #tittle.append(self.excelMatchtable.item(i,1).currentText())
            #tittle[tableData[i]] = self.strTitle.index(self.excelMatchtable.cellWidget(i,1).currentText())
        #self.excelFile.matchTitle(tittle, 'eNodeB Id')
        #self.excelFile.generateKML()
        #except:
                #print 'Calc Information, Unexpected error: ', sys.exc_info()[0]
        

if __name__ == "__main__":
    import sys
    app = QtGui.QApplication(sys.argv)
    MainWindow = QtGui.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

