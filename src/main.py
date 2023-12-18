import os
import sys
import winsound
from pathlib import Path
from time import sleep

from numpy import nan
from pandas import DataFrame, ExcelFile, ExcelWriter, read_excel
from PyQt5 import QtCore, QtGui
from PyQt5.QtWidgets import (QAbstractItemView, QApplication, QDialog,
                             QFileDialog, QInputDialog, QMainWindow,
                             QMessageBox, QShortcut, QSplashScreen,
                             QTableWidgetItem)
from PyQt5.uic import loadUi

from func import convert2StrIntFloat, decodeQRCode, isQRCode, scan_qr_code


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

'______________________________________________________________________________'

class InputSheetsDialog(QDialog):
    def __init__(self):
        super(InputSheetsDialog, self).__init__()
        loadUi(resource_path("./qtui/1.InputDialog.ui"), self)
        self.setWindowTitle("Select Excel sheets")

        self.datasheet_dropdown_1.addItems(['None']+sheets)
        self.decodesheet_dropdown_1.addItems(['None']+sheets)
        self.datasheet_dropdown_1.setCurrentText(sheets[0])
        if len(sheets)==1:
            self.decodesheet_dropdown_1.setCurrentText('None')
        else:
            self.decodesheet_dropdown_1.setCurrentText(sheets[-1])

        self.buttonBox.accepted.connect(self.accept)
        self.buttonBox.rejected.connect(self.reject)
    
    def getSheetNames(self):
        return self.datasheet_dropdown_1.currentText(), self.decodesheet_dropdown_1.currentText()


'______________________________________________________________________________'

class welcome(QMainWindow):
    def __init__(self):
        super(welcome, self).__init__()
        loadUi(resource_path("./qtui/0.Welcome_MW.ui"), self)
        self.setWindowTitle("TabulateQR")
        self.setWindowState(QtCore.Qt.WindowMaximized)

        self.verticalHeader = self.tableWidget.verticalHeader()
        self.verticalHeader.setMaximumSectionSize(35)
        self.tableWidget.setFocusPolicy(QtCore.Qt.NoFocus)

        self.qrCode_data = []
        self.dataMemoryList = {0:(None,None)}
        self.currentState = 0
        self.undoredoState = self.currentState
        self.recordChanges = True

        self.load_btn_wlcm.clicked.connect(self.loadTable)
        self.clear_btn_wlcm.clicked.connect(self.clearTable)
        self.scan_btn_wlcm.clicked.connect(self.getQRCode)
        self.delrow_btn_wlcm.clicked.connect(self.deleteSelRows)
        self.delcol_btn_wlcm.clicked.connect(self.deleteSelCols)
        self.addrow_btn_wlcm.clicked.connect(self.addRow)
        self.addcol_btn_wlcm.clicked.connect(self.addCol)
        self.export_btn_wlcm.clicked.connect(self.export2excel)
        self.undo_btn_wlcm.clicked.connect(self.undo)
        self.redo_btn_wlcm.clicked.connect(self.redo)
        self.tableWidget.cellChanged.connect(self.changeLogged)
        self.useBarcodeScanner_checkBox.toggled.connect(self.toggleScan)

        self.clear_btn_wlcm.setEnabled(False)
        self.scan_btn_wlcm.setEnabled(False)
        self.delrow_btn_wlcm.setEnabled(False)
        self.delcol_btn_wlcm.setEnabled(False)
        self.addrow_btn_wlcm.setEnabled(False)
        self.addcol_btn_wlcm.setEnabled(False)
        self.export_btn_wlcm.setEnabled(False)
        self.undo_btn_wlcm.setEnabled(False)
        self.redo_btn_wlcm.setEnabled(False)
        
        self.ctrl_O = QShortcut(QtGui.QKeySequence("Ctrl+O"), self)
        self.esc    = QShortcut(QtGui.QKeySequence("Esc"), self)
        self.ctrl_Q = QShortcut(QtGui.QKeySequence("Ctrl+Q"), self)
        self.ctrl_S = QShortcut(QtGui.QKeySequence("Ctrl+S"), self)
        self.ctrl_Z = QShortcut(QtGui.QKeySequence("Ctrl+Z"), self)
        self.ctrl_Y = QShortcut(QtGui.QKeySequence("Ctrl+Y"), self)
        
        self.ctrl_O.activated.connect(self.loadTable)
        self.esc.activated.connect(self.clearTable)
        self.ctrl_Q.activated.connect(self.getQRCode)
        self.ctrl_S.activated.connect(self.export2excel)
        self.ctrl_Z.activated.connect(self.undo)
        self.ctrl_Y.activated.connect(self.redo)
        
        self.ctrl_O.setEnabled(True)
        self.esc.setEnabled(False)
        self.ctrl_Q.setEnabled(False)
        self.ctrl_S.setEnabled(False)
        self.ctrl_Z.setEnabled(False)
        self.ctrl_Y.setEnabled(False)

        self.label_tbl.setText('Press `Load Excel` button to load the Excel.')
        

    def loadTable(self):
        self.clearTable()
        self.recordChanges = False
        global excelFile
        excelFile, _ = QFileDialog.getOpenFileName(
                self, 'Single File', '.', '*.xls*')

        if excelFile != '':
            global colsNames, df_decode
            xls = ExcelFile(excelFile)
            global sheets
            sheets = xls.sheet_names

            inputsheetsPopup = InputSheetsDialog()
            global dataSheet, decodeSheet
            dataSheet, decodeSheet = ['None', 'None']
            if inputsheetsPopup.exec_() == QDialog.Accepted:
                dataSheet, decodeSheet = inputsheetsPopup.getSheetNames()
            if dataSheet != 'None':
                df = read_excel(xls, sheet_name=dataSheet)
                try:
                    df_decode = read_excel(
                        xls, sheet_name=decodeSheet, index_col=0)
                except: df_decode = None
                colsNames = df.columns.to_list()

                self.tableWidget.setColumnCount(len(colsNames))
                self.tableWidget.setHorizontalHeaderLabels(colsNames)
                
                global row
                row=0
                for row in df.index:
                    rowPosition = self.tableWidget.rowCount()
                    if 'reload' not in locals():
                        self.tableWidget.insertRow(rowPosition)
                    for col, data in enumerate(df.iloc[row]):
                        output_Item = QTableWidgetItem(str(data))
                        if col != 0:
                            output_Item.setTextAlignment(QtCore.Qt.AlignCenter)
                        self.tableWidget.setItem(row, col, output_Item)
                    row = row + 1
                self.tableWidget.resizeColumnsToContents()
                self.load_btn_wlcm.setEnabled(False)
                self.clear_btn_wlcm.setEnabled(True)
                self.scan_btn_wlcm.setEnabled(True)
                self.delrow_btn_wlcm.setEnabled(True)
                self.delcol_btn_wlcm.setEnabled(True)
                self.addrow_btn_wlcm.setEnabled(True)
                self.addcol_btn_wlcm.setEnabled(True)
                self.export_btn_wlcm.setEnabled(True)

                self.tableWidget.setCurrentItem(None)
                
                self.trackChanges()

                self.label_tbl.setText('Table: Loaded '+'✔')
            else:
                self.label_tbl.setText('Table: Data sheet not selected or Selection Aborted [Please try again] '+'❌')
        else:
            self.label_tbl.setText('Table: Loading Failed [Please try again] '+'❌')
        
        self.enableShortcuts()
        self.recordChanges = True
    
    def clearTable(self):
        self.recordChanges = False
        while self.tableWidget.rowCount() > 0:
            for i in range(self.tableWidget.rowCount()):
                self.tableWidget.removeRow(i)
        while self.tableWidget.columnCount() > 0:
            for i in range(self.tableWidget.columnCount()):
                self.tableWidget.removeColumn(i)
        self.tableWidget.clear()
        self.load_btn_wlcm.setEnabled(True)
        self.clear_btn_wlcm.setEnabled(False)
        self.scan_btn_wlcm.setEnabled(False)
        self.delrow_btn_wlcm.setEnabled(False)
        self.delcol_btn_wlcm.setEnabled(False)
        self.addrow_btn_wlcm.setEnabled(False)
        self.addcol_btn_wlcm.setEnabled(False)
        self.export_btn_wlcm.setEnabled(False)
        
        self.trackChanges()

        self.label_tbl.setText('Press `Load Excel` button to reload the Excel.')
        self.enableShortcuts()
        self.recordChanges = True

    def getQRCode(self):
        qrcode = None
        try:
            qrcode = scan_qr_code()
        except:
            qrcode = None
            warnMsg = QMessageBox()
            warnMsg.setIcon(QMessageBox.Critical)
            warnMsg.setText("No camera found.\nPlease use a barcode scanner for scanning and don't forget to check `Use Barcode Scanner`.")
            warnMsg.setWindowTitle("Warning")
            warnMsg.setStandardButtons(QMessageBox.Ok)
            warnMsg.exec_()
        
        self.writeQRCode(qrcode)
    
    def keyPressEvent(self, event):
        if self.useBarcodeScanner_checkBox.isChecked():
            print(event.text())
            if event.text() != '\r':
                self.qrCode_data.append(event.text())
            else:
                qrcode = ''.join([b for b in self.qrCode_data])
                if isQRCode(qrcode):
                    self.qrCode_data = []
                    self.writeQRCode(qrcode)


    def writeQRCode(self, qrcode):
        self.recordChanges = False
        global row
        
        if qrcode is not None:
            self.tableWidget.clearSelection()
            self.tableWidget.setCurrentItem(None)
            matching_items = self.tableWidget.findItems(
                str(qrcode), QtCore.Qt.MatchContains)
            if matching_items:
                winsound.Beep(1000, 250)
                item = matching_items[0]  # select first item
                # self.tableWidget.setCurrentItem(item)
                self.tableWidget.selectRow(item.row())
                self.tableWidget.scrollToItem(
                    item,QAbstractItemView.ScrollHint.EnsureVisible)
                self.label_tbl.setText(f'QR Code: {qrcode} already exists '+'🔁')
            else:
                winsound.Beep(500, 250)
                output_Item = QTableWidgetItem(str(qrcode))
                rowPosition = self.tableWidget.rowCount()
                self.tableWidget.insertRow(rowPosition)
                self.tableWidget.setItem(row, 0, output_Item)
                self.tableWidget.scrollToItem(
                    output_Item,QAbstractItemView.ScrollHint.EnsureVisible)
                row = row + 1
                self.label_tbl.setText(f'QR Code: {qrcode} scanned '+'✔')
                self.trackChanges()
        else:
            self.label_tbl.setText(f"QR Code: Not detected "+'❌')
        self.enableShortcuts()
        self.recordChanges = True    

    def deleteSelRows(self):
        self.recordChanges = False
        global row
        selectedRows = self.tableWidget.selectionModel().selectedRows()
        if selectedRows:
            row = row - len(selectedRows)
            while self.tableWidget.selectionModel().selectedRows():
                selectedRows = self.tableWidget.selectionModel().selectedRows()
                for selrow in selectedRows:
                    self.tableWidget.removeRow(selrow.row())
            self.tableWidget.resizeColumnsToContents()
            self.label_tbl.setText(f"Table: Selected rows deleted "+'✔')
            
            self.trackChanges()
            
        else:
            self.label_tbl.setText(f"Table: No row selected to delete "+'❌')
        self.enableShortcuts()
        self.recordChanges = True
    
    def deleteSelCols(self):
        self.recordChanges = False
        selectedCols = self.tableWidget.selectionModel().selectedColumns()
        if selectedCols:
            while self.tableWidget.selectionModel().selectedColumns():
                selectedCols = self.tableWidget.selectionModel().selectedColumns()
                for selcol in selectedCols:
                    self.tableWidget.removeColumn(selcol.column())
            self.tableWidget.resizeColumnsToContents()
            self.label_tbl.setText(f"Table: Selected columns deleted "+'✔')
            
            self.trackChanges()

        else:
            self.label_tbl.setText(f"Table: No column selected to delete "+'❌')
        self.enableShortcuts()
        self.recordChanges = True
    
    def addRow(self):
        self.recordChanges = False
        global row
        output_Item = QTableWidgetItem(str(nan))
        rowPosition = self.tableWidget.rowCount()
        self.tableWidget.insertRow(rowPosition)
        self.tableWidget.setItem(row, 0, output_Item)
        self.tableWidget.scrollToItem(
            output_Item,QAbstractItemView.ScrollHint.EnsureVisible)
        row = row + 1
        self.tableWidget.resizeColumnsToContents()
        
        self.trackChanges()
        
        self.label_tbl.setText(f"Table: New row added "+'✔')
        self.enableShortcuts()
        self.recordChanges = True

    def addCol(self):
        self.recordChanges = False
        newColName, ok = QInputDialog.getText(self, 'Add new column', 'Column Name:')
        if ok:
            colPosition = self.tableWidget.columnCount()
            self.tableWidget.setColumnCount(colPosition + 1)
            self.tableWidget.setHorizontalHeaderItem(colPosition, QTableWidgetItem(newColName))
            for row in range(self.tableWidget.rowCount()):
                item = QTableWidgetItem(str(nan))
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                self.tableWidget.setItem(row, colPosition, item)
            colsNames.append(newColName)
            self.tableWidget.resizeColumnsToContents()
            
            self.trackChanges()
            
            self.label_tbl.setText(f"Table: New columns `{newColName}` added "+'✔')
        self.enableShortcuts()
        self.recordChanges = True
    
    def getCurrentTableData(self):
        try:
            rowCount = self.tableWidget.rowCount()
            columnCount = self.tableWidget.columnCount()
            data = []
            for r in range(rowCount):
                rowData = []
                for c in range(columnCount):
                    widgetItem = self.tableWidget.item(r, c)
                    if widgetItem and widgetItem.text:
                        rowData.append(convert2StrIntFloat(widgetItem.text()))
                    else:
                        rowData.append(nan)
                data.append(rowData)
            df = DataFrame(data)
            df.columns = [self.tableWidget.horizontalHeaderItem(col).text() for col in range(columnCount)]
            return df, df_decode
        except: return None, None

    def export2excel(self):
        df, df_decode = self.getCurrentTableData()
        if df is not None:
            options = QFileDialog.Options()
            fileName, _ = QFileDialog.getSaveFileName(self, 
                "Save File As", "", "All Files(*);;Excel Files(*.xlsx)", options = options)
            if fileName:
                try:
                    df_decoded, ord_decodedQRCode = decodeQRCode(qrCodes = df['QR Code'].to_list(), decode_df=df_decode)
                    df = df_decoded.set_index('QR Code').join(df.set_index('QR Code')).reset_index()
                    df = df.sort_values(by = ord_decodedQRCode[1:])
                except: pass
                with ExcelWriter(fileName) as writer:
                    df.to_excel(writer, index=False, sheet_name=dataSheet)
                    try: df_decode.reset_index().to_excel(
                            writer, index=False, sheet_name=decodeSheet)
                    except: pass
                self.label_tbl.setText(f"Export: Success {Path(fileName).name} "+'✔')
            else:
                self.label_tbl.setText(f"Export: Failed [Please provide a file name] "+'❌')
        else:
            self.label_tbl.setText(f"Export: Failed [Data not found] "+'❌')
        self.enableShortcuts()
        
    def undo(self):
        self.recordChanges = False
        self.undoredoState = self.undoredoState - 1
        if len(self.dataMemoryList) == 3:
            if self.undoredoState<=0:
                self.undo_btn_wlcm.setEnabled(False)
                self.redo_btn_wlcm.setEnabled(True)
            elif self.undoredoState==1:
                self.undo_btn_wlcm.setEnabled(True)
                self.redo_btn_wlcm.setEnabled(True)
            elif self.undoredoState>=2:
                self.undo_btn_wlcm.setEnabled(True)
                self.redo_btn_wlcm.setEnabled(False)
        else:
            if self.undoredoState<=0:
                self.undo_btn_wlcm.setEnabled(False)
                self.redo_btn_wlcm.setEnabled(True)
            elif self.undoredoState>=1:
                self.undo_btn_wlcm.setEnabled(True)
                self.redo_btn_wlcm.setEnabled(False)
        if self.undoredoState >= 0:
            print('\nUndo to')
            self.loadCurrentData()
        print(self.undoredoState, self.currentState)
        self.tableWidget.setCurrentItem(None)
        self.enableShortcuts()
        self.recordChanges = True
    
    def redo(self):
        self.recordChanges = False
        self.undoredoState = self.undoredoState + 1
        if len(self.dataMemoryList) == 3:
            if self.undoredoState<=0:
                self.undo_btn_wlcm.setEnabled(False)
                self.redo_btn_wlcm.setEnabled(True)
            elif self.undoredoState==1:
                self.undo_btn_wlcm.setEnabled(True)
                self.redo_btn_wlcm.setEnabled(True)
            elif self.undoredoState>=2:
                self.undo_btn_wlcm.setEnabled(True)
                self.redo_btn_wlcm.setEnabled(False)
        else:
            if self.undoredoState<=0:
                self.undo_btn_wlcm.setEnabled(False)
                self.redo_btn_wlcm.setEnabled(True)
            elif self.undoredoState>=1:
                self.undo_btn_wlcm.setEnabled(True)
                self.redo_btn_wlcm.setEnabled(False)
        if self.undoredoState <= 2:
            print('\nRedo to')
            self.loadCurrentData()
        print(self.undoredoState, self.currentState)
        self.tableWidget.setCurrentItem(None)
        self.enableShortcuts()
        self.recordChanges = True
    
    def loadCurrentData(self):
        self.recordChanges = False
        while self.tableWidget.rowCount() > 0:
            for i in range(self.tableWidget.rowCount()):
                self.tableWidget.removeRow(i)
        while self.tableWidget.columnCount() > 0:
            for i in range(self.tableWidget.columnCount()):
                self.tableWidget.removeColumn(i)
        self.tableWidget.clear()
        global df_decode
        df, df_decode = self.dataMemoryList[self.undoredoState]
        print(df)
        if df is not None:
            colsNames = df.columns.to_list()
            self.tableWidget.setColumnCount(len(colsNames))
            self.tableWidget.setHorizontalHeaderLabels(colsNames)
            
            global row
            row=0
            for row in df.index:
                rowPosition = self.tableWidget.rowCount()
                if 'reload' not in locals():
                    self.tableWidget.insertRow(rowPosition)
                for col, data in enumerate(df.iloc[row]):
                    output_Item = QTableWidgetItem(str(data))
                    if col != 0:
                        output_Item.setTextAlignment(QtCore.Qt.AlignCenter)
                    self.tableWidget.setItem(row, col, output_Item)
                row = row + 1
            self.tableWidget.resizeColumnsToContents()
            self.load_btn_wlcm.setEnabled(False)
            self.clear_btn_wlcm.setEnabled(True)
            self.scan_btn_wlcm.setEnabled(True)
            self.delrow_btn_wlcm.setEnabled(True)
            self.delcol_btn_wlcm.setEnabled(True)
            self.addrow_btn_wlcm.setEnabled(True)
            self.addcol_btn_wlcm.setEnabled(True)
            self.export_btn_wlcm.setEnabled(True)
        else:
            self.load_btn_wlcm.setEnabled(True)
            self.clear_btn_wlcm.setEnabled(False)
            self.scan_btn_wlcm.setEnabled(False)
            self.delrow_btn_wlcm.setEnabled(False)
            self.delcol_btn_wlcm.setEnabled(False)
            self.addrow_btn_wlcm.setEnabled(False)
            self.addcol_btn_wlcm.setEnabled(False)
            self.export_btn_wlcm.setEnabled(False)
        self.enableShortcuts()
        self.recordChanges = True
    
    def changeLogged(self, r, c):
        # print("Cell {} at row {} and column {} was changed.".format(
        #     self.tableWidget.item(r, c).text(), r, c))
        if self.recordChanges:
            self.trackChanges()
    
    def trackChanges(self):
        self.undoredoState = self.currentState
        if self.currentState > 2:
            self.currentState = 2
            self.undoredoState = 2
            self.dataMemoryList[0] = self.dataMemoryList[1]
            self.dataMemoryList[1] = self.dataMemoryList[2]
            self.dataMemoryList[self.currentState] = self.getCurrentTableData()
        else:
            self.dataMemoryList[self.currentState] = self.getCurrentTableData()
            self.currentState = self.currentState + 1
        
        print('\nNext State')
        for k, v in self.dataMemoryList.items():
            print(k, v[0])

        if len(self.dataMemoryList) == 3:
            if self.undoredoState<=0:
                self.undo_btn_wlcm.setEnabled(False)
                self.redo_btn_wlcm.setEnabled(True)
            elif self.undoredoState==1:
                self.undo_btn_wlcm.setEnabled(True)
                self.redo_btn_wlcm.setEnabled(True)
            elif self.undoredoState>=2:
                self.undo_btn_wlcm.setEnabled(True)
                self.redo_btn_wlcm.setEnabled(False)
        else:
            if self.undoredoState<=0:
                self.undo_btn_wlcm.setEnabled(False)
                self.redo_btn_wlcm.setEnabled(True)
            elif self.undoredoState>=1:
                self.undo_btn_wlcm.setEnabled(True)
                self.redo_btn_wlcm.setEnabled(False)
    
    def toggleScan(self):
        if self.useBarcodeScanner_checkBox.isChecked():
            self.scan_btn_wlcm.setEnabled(False)
            self.ctrl_Q.setEnabled(False)
        else:
            self.scan_btn_wlcm.setEnabled(True)
            self.ctrl_Q.setEnabled(True)

    def enableShortcuts(self):
        if self.load_btn_wlcm.isEnabled():
            self.ctrl_O.setEnabled(True)
            self.esc.setEnabled(False)
            self.ctrl_S.setEnabled(False)
            self.ctrl_Q.setEnabled(False)
        else:
            self.ctrl_O.setEnabled(False)
            self.esc.setEnabled(True)
            self.ctrl_Q.setEnabled(True)
            self.ctrl_S.setEnabled(True)
            if self.useBarcodeScanner_checkBox.isChecked():
                self.scan_btn_wlcm.setEnabled(False)
                self.ctrl_Q.setEnabled(False)
            else:
                self.scan_btn_wlcm.setEnabled(True)
                self.ctrl_Q.setEnabled(True)
        if self.undo_btn_wlcm.isEnabled():
            self.ctrl_Z.setEnabled(True)
        else:
            self.ctrl_Z.setEnabled(False)
        if self.redo_btn_wlcm.isEnabled():
            self.ctrl_Y.setEnabled(True)
        else:
            self.ctrl_Y.setEnabled(False)

'################################ QApplication ################################'

# pyinstaller --onedir --noconsole --noconfirm --clean --hidden-import="sklearn.metrics" --add-data "C:\Users\r.gupta\OneDrive - University of Florida\1_Projects\2023 DSSAT-SPACT\qtui;qtui" --icon="C:\Users\r.gupta\OneDrive - University of Florida\1_Projects\2023 TabulateQR\\qtui\\SoilPro.ico" main.py
# pyinstaller --onefile --noconfirm --clean --hidden-import="pyzbar.pyzbar","PIL" --upx-dir="C:\Users\r.gupta\OneDrive - University of Florida\1_Projects\2023 TabulateQR\upx-4.2.1" --add-data "C:\Users\r.gupta\OneDrive - University of Florida\1_Projects\2023 TabulateQR\qtui;qtui" main.py
if __name__ == '__main__':
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    app=QApplication(sys.argv)
    app_icon = QtGui.QIcon()
    app_icon.addFile(resource_path('./qtui/icon.ico'))
    app.setWindowIcon(app_icon)
    
    splash = QSplashScreen(
        QtGui.QPixmap(resource_path("./qtui/SplashScreen_640px.png"))
        # .scaled(
        #     640,576,QtCore.Qt.KeepAspectRatio,
        #     QtCore.Qt.SmoothTransformation),
        # QtCore.Qt.WindowStaysOnTopHint
        )
    mainwindow=welcome()
    splash.show()
    app.processEvents()
    sleep(2)
    mainwindow.resize(720, 480)
    mainwindow.show()
    splash.finish(mainwindow)
    app.exec_()