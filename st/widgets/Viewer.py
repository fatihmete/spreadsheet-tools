from PyQt5.QtWidgets import ( QWidget, QLabel, QLineEdit, QPushButton, 
QTableWidget, QTableWidgetItem, QVBoxLayout, QMessageBox, QHBoxLayout, 
QFileDialog, QInputDialog, QDialogButtonBox, QTableWidgetItem, QDialog,
QPlainTextEdit)                      
from PyQt5.QtCore import QThread, pyqtSignal, pyqtSlot
from PyQt5.QtGui import QFont, QIntValidator
import pandas as pd
import json
import os
import re
import math

class Viewer(QWidget):
    def __init__(self,  parent):
        super().__init__()
        self.parent = parent
        self.saveFlag = False
        self.saveFileName = ""
        self.saveSep = ""
        self.show_rows = 100

        self.layout = QVBoxLayout()
        self.viewer_container = QWidget()
        self.viewer_layout = QVBoxLayout()
        self.viewer_container.setLayout(self.viewer_layout)
        # Select input File
        self.layout_select_input_file = QHBoxLayout()
        self.layout.addWidget(QLabel("Open file (*.xlsx, *.csv, *.txt)"))
        self.layout.addLayout(self.layout_select_input_file)
        self.lineedit_input_file = QLineEdit()
        self.lineedit_input_file.setReadOnly(True)
        #self.lineedit_input_file.textChanged.connect(self.action_open_file)
        self.button_select_input_file = QPushButton("...")
        self.layout_select_input_file.addWidget(self.lineedit_input_file)
        self.layout_select_input_file.addWidget(self.button_select_input_file)
        self.button_select_input_file.clicked.connect(self.action_select_input_file) 

        self.button_delete_column = QPushButton("Drop Cols",self)
        self.button_delete_column.clicked.connect(self.action_delete_column)
        self.button_delete_row = QPushButton("Drop Rows",self)
        self.button_delete_row.clicked.connect(self.action_delete_row)
        self.button_python_code = QPushButton("Python Code")
        self.button_python_code.clicked.connect(self.action_run_python_code)

        #Search Bar
        self.lineedit_search = QLineEdit(self)
        self.lineedit_search.setStyleSheet("""font-size:16px;""")
        self.lineedit_search.setPlaceholderText('Query...')
        self.lineedit_search.returnPressed.connect(self.load_data_frame)
        self.top_layout = QHBoxLayout()
        self.top_layout.addWidget(self.lineedit_search)

        self.top_layout.addWidget(self.button_delete_column)
        self.top_layout.addWidget(self.button_delete_row)
        self.top_layout.addWidget(self.button_python_code)
        self.viewer_layout.addLayout(self.top_layout)

        #Table widget
        self.table_widget = QTableWidget()
        self.table_widget.horizontalHeader().sectionDoubleClicked.connect(self.action_change_sort_criteria)
        self.viewer_layout.addWidget(self.table_widget)  

        #bottom layout
        self.bottom_layout = QHBoxLayout()
        # Navigate Buttons
        self.label_show_rows = QLabel("Show Rows:")
        self.lineedit_show_rows = QLineEdit(str(self.show_rows))
        self.lineedit_show_rows.setFixedWidth(100)
        self.lineedit_show_rows.setValidator(QIntValidator(1,1e6)) #max 1e6, highly possible freeze screen
        self.lineedit_show_rows.returnPressed.connect(self.action_change_show_rows)

        self.button_prev = QPushButton("< Prev",self)
        self.button_next = QPushButton("Next >",self)
        self.button_last = QPushButton("Last >>",self)
        self.button_first = QPushButton("<< First",self)
        self.button_save = QPushButton("Save Data",self)

        self.button_next.clicked.connect(self.action_next_page)
        self.button_prev.clicked.connect(self.action_prev_page)
        self.button_last.clicked.connect(self.action_last_page)
        self.button_first.clicked.connect(self.action_first_page)
        self.button_save.clicked.connect(self.action_save_data_frame)

        self.bottom_layout.addWidget(self.label_show_rows)
        self.bottom_layout.addWidget(self.lineedit_show_rows)
        self.bottom_layout.addWidget(self.button_prev)
        self.bottom_layout.addWidget(self.button_next)
        self.bottom_layout.addWidget(self.button_first)
        self.bottom_layout.addWidget(self.button_last)
        self.bottom_layout.addWidget(self.button_save)

        self.viewer_layout.addLayout(self.bottom_layout)
        
        self.layout.addWidget(self.viewer_container)
        self.viewer_container.setDisabled(True)
        self.setLayout(self.layout)

    def action_run_python_code(self):
        dialog_run_python = RunPythonDialog(parent=self)
        result = dialog_run_python.exec_()
        if result==1:
            code = dialog_run_python.editor_python_code.toPlainText()       
            df = self.data
            try:
                exec(f"""{code}""")
                self.data = df
                self.load_data_frame()
            except Exception as exc:
                self.update_status_bar(f"Error: {exc}")

    def action_change_show_rows(self):
        self.show_rows = int(self.lineedit_show_rows.text())
        self.load_data_frame()

    def action_save_data_frame(self):
        self.saveFlag=True
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_name, _ = QFileDialog.getSaveFileName(self,"Save DataFrame","","CSV (*.csv);;TXT (*.txt);;Excel File (*.xlsx)", options=options)
        if file_name:
            if file_name.split(".")[-1] not in ["csv", "xlsx"]:
                file_name = file_name + "." + re.findall("\*\.(.*)\)", _)[0]
            if file_name.split(".")[-1] == "csv":
                sep, okPressed = QInputDialog.getText(self, "CSV Seperator","Seperator:", QLineEdit.Normal, ",")
                self.saveSep=sep
            self.saveFileName = file_name
        self.load_data_frame()

    def action_delete_column(self):
        selected_cols=[]
        for r in self.table_widget.selectedRanges():
            selected_cols.extend(range(r.leftColumn(),r.rightColumn()+1))
        
        if len(selected_cols)>0:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setText(f"Are you sure to delete selected columns?")
            msg.setWindowTitle("Are you sure?")
            msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            answer = msg.exec_()
            if answer == QMessageBox.Yes:
                self.data.drop([self.data.columns[i] for i in selected_cols], axis=1, inplace=True)
                self.lineedit_search.setText("")
                self.load_data_frame()
       
    def action_delete_row(self):
        selected_rows=[]
        for r in self.table_widget.selectedRanges():
            selected_rows.extend(range(r.topRow(),r.bottomRow()+1))
        selected_rows = [i + (self.active_page * self.show_rows) for i in selected_rows]
        if len(selected_rows)>0:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setText(f"Are you sure to delete selected rows?")
            msg.setWindowTitle("Are you sure?")
            msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            answer = msg.exec_()
            if answer == QMessageBox.Yes:
                if len(self.lineedit_search.text())>0:
                    self.data = self.data[~self.data.index.isin(self.data.query(f"""{self.lineedit_search.text()}""", engine="python").iloc[selected_rows].index)]
                else:
                    self.data = self.data[~self.data.index.isin(self.data.iloc[selected_rows].index)]
                self.load_data_frame()

    def action_select_input_file(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_name, _ = QFileDialog.getOpenFileName(self,"Select input file","","Spreadsheet File (*.csv *.txt *.xlsx)", options=options)
        self.sep=""
        if file_name.split(".")[-1] == "csv":
            sep, okPressed = QInputDialog.getText(self, "CSV Seperator","Seperator:", QLineEdit.Normal, ",")
            self.sep=sep
        self.lineedit_input_file.setText(file_name)
        self.action_open_file()
    
    
    def action_open_file(self):
        self.file_name = self.lineedit_input_file.text()
        self.viewerThread = ViewerThread(file=self.file_name, sep=self.sep, method="open_data_frame")
        self.viewerThread.sgn_message.connect(self.update_status_bar)
        self.viewerThread.sgn_status.connect(self.update_elements)
        self.viewerThread.sgn_dataframe.connect(self.open_df)
        self.viewerThread.start()
    
    def action_next_page(self):
        if self.active_page != self.max_page_number:
            self.active_page +=1
        self.load_data_frame()
        
    def action_prev_page(self):
        if self.active_page != 0:
            self.active_page -=1
        self.load_data_frame()
        
    def action_last_page(self):
        self.active_page = self.max_page_number
        self.load_data_frame()
        
    def action_first_page(self):
        self.active_page = 0
        self.load_data_frame()
        
    def action_change_sort_criteria(self,i):
        # Double click change ascending value
        if i == self.sort_index:
            if self.sort_ascending == True:
                self.sort_ascending = False
            else:
                self.sort_ascending = True
                
        self.sort_index = i
        try:
            self.data = self.data.sort_values(self.data.columns[self.sort_index], ascending=self.sort_ascending)
            self.load_data_frame()
        except Exception as exc:
            self.update_status_bar(f'Error: \n {exc}')

    def load_data_frame(self):
        try:
            df = self.data
            # If user drop all columns
            if df.shape[1]==0:
                self.table_widget.setRowCount(0)
                self.table_widget.setColumnCount(0)
                return None

            #Query search
            if len(self.lineedit_search.text())>=1:
                df = df.query(f"""{self.lineedit_search.text()}""", engine="python")

            # If save flag is 1 pass df to Thread
            if self.saveFlag:
                self.file_name = self.saveFileName
                self.dfsavethread = ViewerThread(parent=self, file=self.saveFileName, sep=self.saveSep, df=df, method="save_data_frame")
                self.dfsavethread.sgn_message.connect(self.update_status_bar)
                self.dfsavethread.sgn_status.connect(self.update_elements)
                self.dfsavethread.start()
                self.saveFlag=False
                return None

            #Calculate page number
            self.rows_count = df.shape[0]
            self.cols_count = df.shape[1]
            self.max_page_number = int(math.ceil(self.rows_count / self.show_rows)) - 1    

            #To fix wrong page bug in search mode 
            if self.active_page > self.max_page_number:
                        self.active_page = 0
            
            #Update status bar
            self.update_status_bar("Page {}/{}, Rows: {} - {}, Total Results : {}"\
                                       .format(self.active_page,\
                                           self.max_page_number,\
                                           (self.active_page * self.show_rows),\
                                           (self.active_page * self.show_rows + self.show_rows),\
                                           self.rows_count))
            
            
            df = df.iloc[(self.active_page * self.show_rows):(self.active_page * self.show_rows + self.show_rows)]
            # Clear rows and cols of QTableWidget
            self.table_widget.setRowCount(0)
            self.table_widget.setColumnCount(0)
            
            # Set QTableWidget
            self.table_widget.setColumnCount(df.shape[1])
            self.header_items = []
            self.table_widget.setHorizontalHeaderLabels([str(c) for c in df.columns])
            self.table_widget.setRowCount(df.shape[0])
            # Fill QTableWidget
            i=0
            for _,row in df.iterrows():
                for col in range(df.shape[1]):
                    w_item = QTableWidgetItem(str(row[col]))
                    self.table_widget.setItem(i, col, w_item)
                i+=1

        except Exception as exc:
            self.update_status_bar(f'Error: \n {exc}')

    @pyqtSlot(object)
    def open_df(self, df):
        self.active_page = 0
        self.sort_index = 0
        self.sort_ascending = False
        self.data = df
        self.lineedit_search.setText("")
        self.load_data_frame()

    @pyqtSlot(int)
    def update_elements(self, status):
        if status==1:
            self.parent.toolbarBox.setDisabled(True)
            self.viewer_container.setDisabled(True)
            self.button_select_input_file.setDisabled(True)
        else:
            self.parent.toolbarBox.setEnabled(True)
            self.viewer_container.setEnabled(True)
            self.button_select_input_file.setEnabled(True)

    @pyqtSlot(str)
    def update_status_bar(self, message):
        self.parent.statusBar().showMessage(message)
        
    def stop_thread(self):
        self.viewerThread.terminate()
        self.update_elements(0)
        self.update_status_bar("Task terminated")

class ViewerThread(QThread):
    sgn_message = pyqtSignal(str)
    sgn_status = pyqtSignal(int)
    sgn_dataframe = pyqtSignal(object)
    
    def __init__(self, parent=None, file=None, sep=None,  df=None, method=None):
        QThread.__init__(self, parent)
        self.file = file 
        self.sep = sep
        self.df = df
        self.method = method

    def run(self):
        getattr(self, self.method)()

    def open_data_frame(self):
        self.sgn_status.emit(1)
        self.sgn_message.emit(f"File opening..")
        try:
            if self.file.split(".")[-1] == "xlsx":
                df = pd.read_excel(self.file, engine="openpyxl")
            else:
                df = pd.read_csv(self.file, sep=self.sep)

            self.sgn_dataframe.emit(df)
            self.sgn_status.emit(0)
            
        except Exception as exc:
            self.sgn_message.emit(f'Please check required areas. Error: \n {exc}')
            self.sgn_status.emit(2)
    
    def save_data_frame(self):
        self.sgn_status.emit(1)
        self.sgn_message.emit(f"File saving...")
        try:
            if self.file.split(".")[-1] == "xlsx":
                self.df.to_excel(self.file, index=False)
            else:
                self.df.to_csv(self.file, sep=self.sep)

            self.sgn_message.emit(f"File has been save at {self.file}")
            self.sgn_status.emit(0)
            
        except Exception as exc:
            self.sgn_message.emit(f'Please check required areas. Error: \n {exc}')
            self.sgn_status.emit(2)

class RunPythonDialog(QDialog):
    def __init__(self, parent):
        super(RunPythonDialog, self).__init__()
        
        self.setWindowTitle("Run python code!")
        self.parent = parent
        self.editor_python_code =  QPlainTextEdit()

        self.editor_python_code.setStyleSheet(""" background-color: #162F3E; font-size:16px; color:white;""")

        self.editor_font = QFont()
        self.editor_font.setFamily("Courier")
        self.editor_font.setStyleHint(QFont().Monospace)
        self.editor_font.setPointSize(14)
        self.editor_font.setFixedPitch(True)

        self.editor_python_code.setFont(self.editor_font)

        self.label_text = QLabel("""You can run python code on loaded data. 
        \nLoaded data data frame variable is "df" and pandas variable is "pd".
        \nYou can use all pandas functions as well as all python functions.
        \nFor example you can create new column with "df["new_col"] = df["df_existing_1"] + df["df_existing_2"]" command
         """)
        
        QBtn = QDialogButtonBox.Ok | QDialogButtonBox.Cancel
        
        self.buttonBox = QDialogButtonBox(QBtn)
        self.buttonBox.accepted.connect(self.accept)
        self.buttonBox.rejected.connect(self.reject)

        self.layout = QVBoxLayout()
        self.layout.addWidget(self.label_text)
        self.layout.addWidget(self.editor_python_code)
        self.layout.addWidget(self.buttonBox)
        self.setLayout(self.layout)