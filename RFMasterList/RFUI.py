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
from Logging import RFLogging
import coloreditorfactory

import logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')


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
        self.int_status = 0
        self.RFDB = RFDataBase()
        self.rf_kml = RFKml()
        #create a map between formatted keyword with Excel title {RFIndex:ExcelTitle}
        self.map_title = dict()
        self.map_title.fromkeys(self.RFDB.RFIndex)
        self.rules = dict()

        self.tabWidget.addTab(self.tab_loading(), _fromUtf8(""))
        self.tab_setting()
        self.tab_calculating()

                
        MainWindow.setCentralWidget(self.centralwidget)        
        self.menubar = QtGui.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 600, 18))
        self.menubar.setObjectName(_fromUtf8("menubar"))
        self.menuFile = QtGui.QMenu(self.menubar)
        self.menuFile.setObjectName(_fromUtf8("menuFile"))
        MainWindow.setMenuBar(self.menubar)

        self.retranslateUi(MainWindow)
        self.tabWidget.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(_translate("MainWindow", "Excel2KMZ", None))
        self.tabWidget.setTabText(0, _translate("MainWindow", "LoadExcel", None))        
        
        self.tabWidget.setTabText(1, _translate("MainWindow", "Setting", None))
        self.menuFile.setTitle(_translate("MainWindow", "File", None))
        #self.actionOpen.setText(_translate("MainWindow", "Open", None))
        self.qpb_calc.setText(_translate("MainWindow", "Calc", None))
        self.tabWidget.setTabText(2, _translate("MainWindow", "Calc", None))

       
    def tab_loading(self):
        tab_load = QtGui.QWidget()
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
        qpb_browse.clicked.connect(partial(self._browseFile, tab_load, str_file, cob_sheet))
        qpb_browse.setText(_translate("MainWindow", "Browse", None))
        #qpb_browse.setObjectName(_fromUtf8("qpb_browse"))
        
        tableData = self.RFDB.RFIndex
        tbl_mTitle = QtGui.QTableWidget(len(tableData),2, tab_load)
        tbl_mTitle.setHorizontalHeaderLabels(["Keywords", "Column"])
        tbl_mTitle.verticalHeader().setVisible(False)
        tbl_mTitle.setGeometry(QtCore.QRect(10, 40, 561, 311))
        tbl_mTitle.setEditTriggers(QtGui.QAbstractItemView.NoEditTriggers)

        for i, keywords in enumerate(tableData):
            kwItem = QtGui.QTableWidgetItem(keywords)
            tbl_mTitle.setItem(i,0,kwItem)
            #tbl_mTitle.setItem(i,1,QtGui.QTableWidgetItem('...'))
        
        tbl_mTitle.resizeColumnToContents(1)
        tbl_mTitle.horizontalHeader().setStretchLastSection(True)

        qpb_analysis = QtGui.QPushButton(tab_load)
        qpb_analysis.setGeometry(QtCore.QRect(505, 10, 75, 20))
        qpb_analysis.clicked.connect(partial(self._readTitle, cob_sheet, tbl_mTitle))
        qpb_analysis.setText(_translate("MainWindow", "Analysis", None))
        return tab_load
    
    def _readTitle(self, cob_sheet, tbl_mTitle):
        #self.RFDB.readSheet
        if self.int_status < 4:
            msgBox = QtGui.QMessageBox()
            msgBox.setWindowTitle("Warning")
            msgBox.setText("Please choose Excel First!")
            msgBox.exec_()
            return False          
        if self.RFDB.readTitle(cob_sheet.currentIndex()):
            strTitle = self.RFDB.str_title            
            #read title successfully
            self.int_status = 8
        else:
            msgBox = QtGui.QMessageBox()
            msgBox.setWindowTitle("Warning")
            msgBox.setText("Cannot find any keyword in title from selected Sheet \
            \n your title must contain at least one of below:\n \
            %s" ";".join(self.RFDB.RFIndex))
            msgBox.exec_()
            return False
        #auto match the title with index        
        for i_row in range(len(self.RFDB.RFIndex)):
            str_index = self.RFDB.RFIndex[i_row]
            comboTittle = QtGui.QComboBox()
            comboTittle.addItems(strTitle)
            try:
                str_matched = difflib.get_close_matches(str_index,strTitle)[0]
                index_match = strTitle.index(str_matched)
                self.map_title[str_index] = str_matched 
                #find the matched title
                comboTittle.setCurrentIndex(index_match)
                tbl_mTitle.item(i_row, 0).setBackgroundColor(QtGui.QColor(0,255,13))
            except IndexError:
                index_match = -1
            tbl_mTitle.setCellWidget(i_row, 1, comboTittle)
            comboTittle.currentIndexChanged.connect(partial(self._updateTableValue, str_index, comboTittle, self.map_title))
            
        self.tabWidget.setTabEnabled(1,True)
        self.int_status = 12

    def _browseFile(self, tab_load, str_file, cob_sheet):
        fileName = QtGui.QFileDialog.getOpenFileName(tab_load,"Open File",'.',"Excel File (*.xlsx *.xls)")    
        #if RFDB.readSheet:
        if not fileName:
            return False
        str_shnames = self.RFDB.readFile(fileName)
        if str_shnames:
            str_file.setText(fileName)
            cob_sheet.clear()
            self.int_status = 4
            cob_sheet.addItems(str_shnames)

        else:
            msgBox = QtGui.QMessageBox()
            msgBox.setWindowTitle("Warning")
            msgBox.setText("Open Excel Failed!")
            msgBox.exec_()
            return False    
    
    def tab_setting(self):
        #setting for tab_set        
        self.tab_set = QtGui.QWidget()
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
        
        KML_options = self.rf_kml.supportOptions[1]
        RF_type = self.rf_kml.supportOptions[0]
        qcb_type.addItems(KML_options)
        qcb_value.addItems(RF_type)
        #need to update this value after changed default rule number
        setTable = QtGui.QTableWidget(1,4,self.tab_set)
        setTable.horizontalHeader().setVisible(False)
        setTable.verticalHeader().setVisible(True)
        setTable.setGeometry(QtCore.QRect(40, 50, 500, 200))
        self.tabWidget.addTab(self.tab_set, _fromUtf8(""))
        self.tabWidget.currentChanged.connect(partial(self._defaultOptions, setTable))
        
        qpb_add = QtGui.QPushButton(self.tab_set)
        qpb_add.setGeometry(QtCore.QRect(500, 20, 75, 20))
        qpb_add.setText(_translate("MainWindow", "Add", None))
        qpb_add.clicked.connect(partial(self.updateRuleTable, setTable))        
        
        #self.tabWidget.setTabEnabled(1,False)
    
    def _defaultOptions(self, setTable):
        if self.tabWidget.currentIndex() == 1 and self.int_status == 12:
            #build default rules
            rf_item = dict()
            self.rf_log = RFLogging()
            
            #read maptitle and RFDB
            self.RFDB.readSheet(self.map_title)
            #change to tab_setting, do KPI statistics now
            rf_item["EARFCN"] = self.RFDB.DBStatistics("EARFCN")
            rf_item["Type"] = self.RFDB.DBStatistics("Type")
            rf_item["PCI"] = ['0', '1', '2']

            kml_item = dict()
            kml_item["Shape"] = ["Circle", "Sector"]
            kml_item["Cell Radius"] = ["0.00025", "0.001265"]
            kml_item["Height"] = list()
            #kml_item["Shape"] = [["Circle","0.00025","20"],["Sector","0.001265","10"]]
            kml_item["Line Color"] = list()
            kml_item["Fill Color"] = list()

            self.rules = dict()
            self.rules[("EARFCN", "Fill Color")] = dict().fromkeys(rf_item["EARFCN"], "red")
            self.rules[("PCI", "Line Color")] = dict().fromkeys(rf_item["PCI"], "red")
            self.rules[("Type", "Shape")] = dict.fromkeys(rf_item["Type"],"Circle")
            self.rules[("Type", "Cell Radius")] = dict(zip(rf_item["Type"],kml_item["Cell Radius"]))
            self.rules[("EARFCN", "Height")] = dict().fromkeys(rf_item["EARFCN"], "0")
            
            for rule_type in self.rf_log.rules:
                self.rules[rule_type].update(self.rf_log.rules[rule_type])

            self.updateRuleTable(setTable)
            
            self.int_status = 32
        elif self.tabWidget.currentIndex() == 0:
            self.int_status = 12

    def updateRuleTable(self, setTable):
        setTable.clear()
        setTable.setRowCount(len(self.rules))
        for i, rule_type in enumerate(self.rules):
            kml_Item = QtGui.QTableWidgetItem(rule_type[1])
            rf_Item = QtGui.QTableWidgetItem(rule_type[0])
            pub_options = QtGui.QPushButton('...')
            pub_options.clicked.connect(partial(self._set_options, rule_type, self.rules[rule_type]))
            pub_remove = QtGui.QPushButton('Remove')
            pub_remove.clicked.connect(partial(self.tab_set_remove,i))
            setTable.setItem(i,0,rf_Item)
            setTable.setItem(i,1,kml_Item)
            setTable.setCellWidget(i,2,pub_options)
            setTable.setCellWidget(i,3,pub_remove)
        setTable.resizeColumnToContents(1)

 
    def _set_options(self, rule_type, rule_value):
        kml_type = rule_type[1]

        if kml_type.find('Color') != -1:
            color_dialog = QtGui.QDialog()
            #color_dialog.setModal(True)
            coloreditorfactory.Window(color_dialog,rule_value)
            color_dialog.resize(275,270)
            color_dialog.setWindowTitle(kml_type)
            #color_factory.show()  
            color_dialog.exec_()

        else:
            shape_dialog = QtGui.QDialog()
            #shape_dialog.setModal(True)
            shape_dialog.resize(410,270)
            shape_dialog.setWindowTitle(kml_type)
            
            shape_table = QtGui.QTableWidget(len(rule_value),2,shape_dialog)
            shape_table.horizontalHeader().setVisible(True)
            shape_table.setHorizontalHeaderLabels(("RF_Value", kml_type))
            shape_table.verticalHeader().setVisible(False)
            shape_table.resize(410, 270)
            for i, rf_value in enumerate(rule_value):
                nameItem = QtGui.QTableWidgetItem(rf_value)
                nameItem.setFlags(QtCore.Qt.ItemIsEditable)
                shape_table.setItem(i, 0, nameItem)
                #shapeItem = QtGui.QTableWidgetItem()
                if kml_type.find('Shape') != -1:
                    opt_item = QtGui.QComboBox()
                    opt_item.addItems(["Sector", "Circle"])
                    j = ["Sector", "Circle"].index(rule_value[rf_value])
                    opt_item.setCurrentIndex(j)
                    shape_table.setCellWidget(i, 1, opt_item)
                    opt_item.currentIndexChanged.connect(partial(self._updateTableValue, rf_value, opt_item, rule_value))
                    
                else:
                    #opt_item = QtGui.QLineEdit()
                    opt_item = QtGui.QTableWidgetItem(rule_value[rf_value])
                    #opt_item.setText(rule_value[rf_value])
                    #shape_table.setCellWidget(i, 1, opt_item)
                    shape_table.setItem(i, 1, opt_item)
                    #opt_item.textChanged.connect(partial(self._updateTableValue,rf_value, opt_item, rule_value))
            shape_table.itemChanged.connect(partial(self._catchCurrentItem, shape_table, rule_value))            

            shape_dialog.setWindowTitle(kml_type)
            shape_dialog.exec_()
    def _catchCurrentItem(self, shape_table, rule_value):
        self._updateTableValue(str(shape_table.item(shape_table.currentRow(),0).text()), shape_table.currentItem(), rule_value)
    def _updateTableValue(self, str_key, cellItem, des_dict):
        #cellItem can be anytype of cellwidgetItem, des_dict is dict for store value
        #default widgetTable is only 2 column, 1st for key of dict, 2nd for value of dict
        try:            
            str_value = str(cellItem.text())
            if not str_value.replace(".","",1).isdigit():
                cellItem.setText(des_dict[str_key])
                str_value = des_dict[str_key]
                msgBox = QtGui.QMessageBox()
                msgBox.setWindowTitle("Warning")
                msgBox.setText("Please Input a Valid Number!")
                msgBox.exec_()
        except:
            str_value = str(cellItem.currentText())
                    
        des_dict.update({str_key:str_value})
        
        #logging.debug("Successfully update dict to %s" %str(des_dict))

        
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
        if self.rf_log.writeRuleToCfg(self.rules):
            logging.debug("Setting Records Successfully")
        self.rf_kml.createRules(self.rules)
        
        self.rf_kml.createCells(self.RFDB.DB, self.map_title)
            #tittle.append(self.tbl_mTitle.item(i,1).currentText())
            #tittle[tableData[i]] = self.strTitle.index(self.tbl_mTitle.cellWidget(i,1).currentText())
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

