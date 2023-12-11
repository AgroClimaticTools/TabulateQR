import os
import sys
import winsound
from pathlib import Path

from numpy import nan
from pandas import DataFrame, ExcelWriter, read_excel
from PyQt5 import QtCore, QtGui
from PyQt5.QtWidgets import (QAbstractItemView, QApplication, QDialog,
                             QFileDialog, QInputDialog, QMainWindow,
                             QMessageBox, QSplashScreen, QStackedWidget,
                             QTableWidgetItem)
from PyQt5.uic import loadUi

from func import convert2StrIntFloat, decodeQRCode, scan_qr_code


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

'______________________________________________________________________________'

# class welcome(QDialog):
class welcome(QMainWindow):
    def __init__(self):
        super(welcome, self).__init__()
        loadUi(resource_path("./qtui/0.Welcome_MW.ui"), self)
        self.setWindowTitle("TabulateQR")
        self.setWindowState(QtCore.Qt.WindowMaximized)

        self.dataMemoryList = {0:(None,None)}
        self.currentState = 1
        self.undoredoState = self.currentState

        # self.tableWidget.cellChanged.connect(self.changeLogged)

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

        self.clear_btn_wlcm.setEnabled(False)
        self.scan_btn_wlcm.setEnabled(False)
        self.delrow_btn_wlcm.setEnabled(False)
        self.delcol_btn_wlcm.setEnabled(False)
        self.addrow_btn_wlcm.setEnabled(False)
        self.addcol_btn_wlcm.setEnabled(False)
        self.export_btn_wlcm.setEnabled(False)
        self.undo_btn_wlcm.setEnabled(False)
        self.redo_btn_wlcm.setEnabled(False)

        self.label_tbl.setText('Press `Load Excel` button to load the Excel.')

    def loadTable(self):

        global excelFile
        excelFile, _ = QFileDialog.getOpenFileName(
                self, 'Single File', '.', '*.xls*')

        if excelFile != '':
            global colsNames, df_decode
            df = read_excel(excelFile, sheet_name=0)
            try: 
                df_decode = read_excel(
                    excelFile, sheet_name=1, index_col=0)
            except: pass
            colsNames = df.columns.to_list()

            self.tableWidget.setColumnCount(len(colsNames))
            self.tableWidget.setHorizontalHeaderLabels(colsNames)

            global qrcode_list 
            qrcode_list = list(set(df['QR Code'].astype(str).to_list()))
            
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

            self.undoredoState = self.currentState
            if self.currentState > 2:
                self.currentState = 2
                self.dataMemoryList[0] = self.dataMemoryList[1]
                self.dataMemoryList[1] = self.dataMemoryList[2]
                self.dataMemoryList[self.currentState] = self.getCurrentTableData()
            else:
                self.dataMemoryList[self.currentState] = self.getCurrentTableData()
                self.currentState = self.currentState + 1
            
            print(self.dataMemoryList,'\n')
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

            self.label_tbl.setText('Table: Loaded '+'✔')
        else:
            self.label_tbl.setText('Table: Loading Failed [Press `Load Excel` button to reload the Excel] '+'❌')
    
    def clearTable(self):
        while self.tableWidget.rowCount() > 0:
            for i in range(self.tableWidget.rowCount()):
                self.tableWidget.removeRow(i)
        self.tableWidget.clear()
        self.load_btn_wlcm.setEnabled(True)
        self.clear_btn_wlcm.setEnabled(False)
        self.scan_btn_wlcm.setEnabled(False)
        self.delrow_btn_wlcm.setEnabled(False)
        self.delcol_btn_wlcm.setEnabled(False)
        self.addrow_btn_wlcm.setEnabled(False)
        self.addcol_btn_wlcm.setEnabled(False)
        self.export_btn_wlcm.setEnabled(False)
        
        self.undoredoState = self.currentState
        if self.currentState > 2:
            self.currentState = 2
            self.dataMemoryList[0] = self.dataMemoryList[1]
            self.dataMemoryList[1] = self.dataMemoryList[2]
            self.dataMemoryList[self.currentState] = self.getCurrentTableData()
        else:
            self.dataMemoryList[self.currentState] = self.getCurrentTableData()
            self.currentState = self.currentState + 1
        
        print(self.dataMemoryList,'\n')

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

        self.label_tbl.setText('Press `Load Excel` button to reload the Excel.')

    def getQRCode(self):
        global row, qrcode_list
        qrcode = scan_qr_code()
        if qrcode is not None:        
            if qrcode in qrcode_list:
                self.tableWidget.setCurrentItem(None)
                matching_items = self.tableWidget.findItems(
                    str(qrcode), QtCore.Qt.MatchContains)
                if matching_items:
                    item = matching_items[0]  # Take the first.
                    self.tableWidget.setCurrentItem(item)
                    self.tableWidget.scrollToItem(
                        item,QAbstractItemView.ScrollHint.EnsureVisible)
                self.label_tbl.setText(f'QR Code: {qrcode} already exists '+'🔁')
            else:
                qrcode_list.append(qrcode)
                winsound.Beep(1000, 500)
                output_Item = QTableWidgetItem(str(qrcode))
                rowPosition = self.tableWidget.rowCount()
                self.tableWidget.insertRow(rowPosition)
                self.tableWidget.setItem(row, 0, output_Item)
                self.tableWidget.scrollToItem(
                    output_Item,QAbstractItemView.ScrollHint.EnsureVisible)
                row = row + 1
                self.label_tbl.setText(f'QR Code: {qrcode} scanned '+'✔')
                
                self.undoredoState = self.currentState
                if self.currentState > 2:
                    self.currentState = 2
                    self.dataMemoryList[0] = self.dataMemoryList[1]
                    self.dataMemoryList[1] = self.dataMemoryList[2]
                    self.dataMemoryList[self.currentState] = self.getCurrentTableData()
                else:
                    self.dataMemoryList[self.currentState] = self.getCurrentTableData()
                    self.currentState = self.currentState + 1
                
                print(self.dataMemoryList,'\n')

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

        else:
            self.label_tbl.setText(f"QR Code: Not detected "+'❌')

    # def changeLogged(self, r, c):
    #     print("Cell {} at row {} and column {} was changed.".format(
    #         self.tableWidget.item(r, c).text(), r, c))

    def deleteSelRows(self):
        global row, qrcode_list
        selectedRows = self.tableWidget.selectionModel().selectedRows()
        if selectedRows:
            row = row - len(selectedRows)
            while self.tableWidget.selectionModel().selectedRows():
                selectedRows = self.tableWidget.selectionModel().selectedRows()
                # self.tableWidget.removeRow(self.tableWidget.currentRow())
                for selrow in selectedRows:
                    self.tableWidget.removeRow(selrow.row())
            self.tableWidget.resizeColumnsToContents()
            self.label_tbl.setText(f"Table: Selected rows deleted "+'✔')
            
            self.undoredoState = self.currentState
            if self.currentState > 2:
                self.currentState = 2
                self.dataMemoryList[0] = self.dataMemoryList[1]
                self.dataMemoryList[1] = self.dataMemoryList[2]
                self.dataMemoryList[self.currentState] = self.getCurrentTableData()
            else:
                self.dataMemoryList[self.currentState] = self.getCurrentTableData()
                self.currentState = self.currentState + 1
            
            print(self.dataMemoryList,'\n')

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
            
        else:
            self.label_tbl.setText(f"Table: No row selected to delete "+'❌')
    
    def deleteSelCols(self):
        selectedCols = self.tableWidget.selectionModel().selectedColumns()
        if selectedCols:
            while self.tableWidget.selectionModel().selectedColumns():
                selectedCols = self.tableWidget.selectionModel().selectedColumns()
                # self.tableWidget.removeRow(self.tableWidget.currentRow())
                for selcol in selectedCols:
                    # selectedQRCode = self.tableWidget.item(selrow.row(), 0).text()
                    # if selectedQRCode in qrcode_list:
                    #     qrcode_list.remove(selectedQRCode)
                    self.tableWidget.removeColumn(selcol.column())
            self.tableWidget.resizeColumnsToContents()
            self.label_tbl.setText(f"Table: Selected columns deleted "+'✔')
            
            self.undoredoState = self.currentState
            if self.currentState > 2:
                self.currentState = 2
                self.dataMemoryList[0] = self.dataMemoryList[1]
                self.dataMemoryList[1] = self.dataMemoryList[2]
                self.dataMemoryList[self.currentState] = self.getCurrentTableData()
            else:
                self.dataMemoryList[self.currentState] = self.getCurrentTableData()
                self.currentState = self.currentState + 1
            
            print(self.dataMemoryList,'\n')

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

        else:
            self.label_tbl.setText(f"Table: No column selected to delete "+'❌')
    
    def addRow(self):
        global row
        output_Item = QTableWidgetItem(str(nan))
        rowPosition = self.tableWidget.rowCount()
        self.tableWidget.insertRow(rowPosition)
        self.tableWidget.setItem(row, 0, output_Item)
        self.tableWidget.scrollToItem(
            output_Item,QAbstractItemView.ScrollHint.EnsureVisible)
        row = row + 1
        self.tableWidget.resizeColumnsToContents()
        
        self.undoredoState = self.currentState
        if self.currentState > 2:
            self.currentState = 2
            self.dataMemoryList[0] = self.dataMemoryList[1]
            self.dataMemoryList[1] = self.dataMemoryList[2]
            self.dataMemoryList[self.currentState] = self.getCurrentTableData()
        else:
            self.dataMemoryList[self.currentState] = self.getCurrentTableData()
            self.currentState = self.currentState + 1
        
        print(self.dataMemoryList,'\n')

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
        
        self.label_tbl.setText(f"Table: New row added "+'✔')

    def addCol(self):
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
            
            self.undoredoState = self.currentState
            if self.currentState > 2:
                self.currentState = 2
                self.dataMemoryList[0] = self.dataMemoryList[1]
                self.dataMemoryList[1] = self.dataMemoryList[2]
                self.dataMemoryList[self.currentState] = self.getCurrentTableData()
            else:
                self.dataMemoryList[self.currentState] = self.getCurrentTableData()
                self.currentState = self.currentState + 1
            
            print(self.dataMemoryList,'\n')

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
            
            self.label_tbl.setText(f"Table: New columns `{newColName}` added "+'✔')
    
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
                    df.to_excel(writer, index=False, sheet_name='Records')
                    try: df_decode.reset_index().to_excel(
                            writer, index=False, sheet_name='DecodeQR')
                    except: pass
                self.label_tbl.setText(f"Export: Success {Path(fileName).name} "+'✔')
            else:
                self.label_tbl.setText(f"Export: Failed [Please provide a file name] "+'❌')
        else:
            self.label_tbl.setText(f"Export: Failed [Data not found] "+'❌')
        
    def undo(self):
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
            print('Undo to')
            self.loadCurrentData()
        print(self.undoredoState, self.currentState)
    
    def redo(self):
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
            print('Redo to')
            self.loadCurrentData()
        print(self.undoredoState, self.currentState)
    
    def loadCurrentData(self):
        while self.tableWidget.rowCount() > 0:
            for i in range(self.tableWidget.rowCount()):
                self.tableWidget.removeRow(i)
        self.tableWidget.clear()
        global df_decode
        df, df_decode = self.dataMemoryList[self.undoredoState]
        print(df, df_decode)
        if df is not None:
            colsNames = df.columns.to_list()
            self.tableWidget.setColumnCount(len(colsNames))
            self.tableWidget.setHorizontalHeaderLabels(colsNames)

            global qrcode_list
            qrcode_list = list(set(df['QR Code'].astype(str).to_list()))
            
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

'################################ QApplication ################################'

# pyinstaller --onedir --noconsole --noconfirm --clean --hidden-import="sklearn.metrics" --add-data "C:\Users\r.gupta\OneDrive - University of Florida\1_Projects\2023 DSSAT-SPACT\qtui;qtui" --icon="C:\Users\r.gupta\OneDrive - University of Florida\1_Projects\2023 TabulateQR\\qtui\\SoilPro.ico" main.py
# pyinstaller --onefile --noconfirm --clean --hidden-import="pyzbar.pyzbar","PIL" --upx-dir="C:\Users\r.gupta\OneDrive - University of Florida\1_Projects\2023 TabulateQR\upx-4.2.1" --add-data "C:\Users\r.gupta\OneDrive - University of Florida\1_Projects\2023 TabulateQR\qtui;qtui" main.py
if __name__ == '__main__':
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    app=QApplication(sys.argv)
    app_icon = QtGui.QIcon()
    app_icon.addFile(resource_path('./qtui/icon.ico'))
    app.setWindowIcon(app_icon)
    
    mainwindow=welcome()
    mainwindow.resize(720, 480)
    mainwindow.show()
    app.exec_()