from PyQt5.QtWidgets import (QWidget, QLabel, QLineEdit, QPushButton,
QVBoxLayout, QHBoxLayout, QComboBox, QFileDialog)              
from PyQt5.QtCore import QDir, QThread, pyqtSignal, pyqtSlot
from PyQt5.QtGui import QFont
from glob import glob
import pandas as pd
import os

class Splitter(QWidget):

    def __init__(self,  parent):
        super().__init__()

        self.parent = parent
        self.layout = QVBoxLayout()
        
        # Select İnput File
        self.layout_select_input_file = QHBoxLayout()
        self.layout.addWidget(QLabel("Input File (*.xlsx, *.csv - file you want to split):"))
        self.layout.addLayout(self.layout_select_input_file)
        self.lineedit_input_file = QLineEdit()
        self.lineedit_input_file.setReadOnly(True)
        self.button_select_input_file = QPushButton("...")
        self.layout_select_input_file.addWidget(self.lineedit_input_file)
        self.layout_select_input_file.addWidget(self.button_select_input_file)
        self.button_select_input_file.clicked.connect(self.action_select_input_file)

        # Select Output Files Location
        self.layout_select_output_files_path = QHBoxLayout()
        self.layout.addWidget(QLabel("Output Files Path (The location of the new files):"))
        self.layout.addLayout(self.layout_select_output_files_path)
        self.lineedit_output_files_path = QLineEdit()
        self.lineedit_output_files_path.setReadOnly(True)
        self.button_select_output_files_path = QPushButton("...")
        self.layout_select_output_files_path.addWidget(self.lineedit_output_files_path)
        self.layout_select_output_files_path.addWidget(self.button_select_output_files_path)
        self.button_select_output_files_path.clicked.connect(self.action_select_output_files_path)

        # File Types
        self.layout_select_file_format = QVBoxLayout()
        self.layout.addWidget(QLabel("Select format of new files: "))
        self.layout.addLayout(self.layout_select_file_format)
        self.combo_format = QComboBox()
        self.combo_format.addItems([ "*.csv", "*.xlsx"])
        self.layout_select_file_format.addWidget(self.combo_format)

        # CSV Seperator
        self.layout.addWidget(QLabel("Row counts of new files:"))
        self.lineedit_row_count = QLineEdit("1000")
        self.layout.addWidget(self.lineedit_row_count)

        # CSV Seperator
        self.layout.addWidget(QLabel("CSV/Text Seperator for input file:"))
        self.lineedit_csv_sep_input = QLineEdit(",")
        self.layout.addWidget(self.lineedit_csv_sep_input)

        self.layout.addWidget(QLabel("CSV/Text Seperator for output file:"))
        self.lineedit_csv_sep_output = QLineEdit(",")
        self.layout.addWidget(self.lineedit_csv_sep_output)

        #Merge Button
        self.button_split_files = QPushButton("Split Files")
        self.layout.addWidget(self.button_split_files)
        self.button_split_files.clicked.connect(self.split_files)

        #Stop Button
        self.button_stop_split = QPushButton("Stop Split")
        self.button_stop_split.clicked.connect(self.stop_thread)
        self.button_stop_split.setVisible(False)
        self.layout.addWidget(self.button_stop_split)         

        self.layout.addStretch()
        self.setLayout(self.layout)
        
    def action_select_output_files_path(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        options |= QFileDialog.DontUseCustomDirectoryIcons
        dialog = QFileDialog()
        dialog.setOptions(options)
        dialog.setFilter(dialog.filter() | QDir.Hidden)
        path  = dialog.getExistingDirectory(self, 'Select directory', options=options)
        if path:
            self.lineedit_output_files_path.setText(path)
    
    def action_select_input_file(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_name, _ = QFileDialog.getOpenFileName(self,"Select input file","","CSV (*.csv);;TXT (*.txt);;Excel File (*.xlsx)", options=options)
        self.lineedit_input_file.setText(file_name)
    
    def split_files(self):
        rules = {}
        rules["output_files_path"] = self.lineedit_output_files_path.text()
        rules["input_file"] = self.lineedit_input_file.text()
        rules["seperator_input"] = self.lineedit_csv_sep_input.text()
        rules["seperator_output"] = self.lineedit_csv_sep_output.text()
        rules["file_format"] = self.combo_format.currentIndex()
        rules["row_count"] = int(self.lineedit_row_count.text())

        self.splitterThread = SplitterThread(rules=rules)
        self.splitterThread.sgn_message.connect(self.update_status_bar)
        self.splitterThread.sgn_status.connect(self.update_elements)
        self.splitterThread.start()

    def update_elements(self, status):
        if status==1:
            self.parent.toolbarBox.setDisabled(True)
            self.button_split_files.setDisabled(True)
            self.button_stop_split.setVisible(True)
        else:
            self.button_stop_split.setVisible(False)
            self.parent.toolbarBox.setEnabled(True)
            self.button_split_files.setEnabled(True)

    def update_status_bar(self, message):
        self.parent.statusBar().showMessage(message)
        
    def stop_thread(self):
        self.splitterThread.terminate()
        self.update_elements(0)
        self.update_status_bar("Task terminated")

class SplitterThread(QThread):
    sgn_message = pyqtSignal(str)
    sgn_status = pyqtSignal(int)
    
    def __init__(self, parent=None, rules=None):
        QThread.__init__(self, parent)
        self.rules = rules

    @pyqtSlot()
    def run(self):
        self.sgn_status.emit(1)
        self.sgn_message.emit(f"Starting")
        try:
            if self.rules["input_file"].split(".")[-1] == "xlsx":
                df = pd.read_excel(self.rules["input_file"], engine="openpyxl")
            else:
                df = pd.read_csv(self.rules["input_file"], sep=self.rules["seperator_input"])

            row_count = self.rules["row_count"]
            file_count = int(df.shape[0] / row_count)
            residual = df.shape[0] % row_count
            if residual!=0:
                file_count+=1
            s = 0
            for i in range(1,file_count+1):
                if i==file_count:
                    sl = slice(s,(df.shape[0]-1))
                else:
                    sl = slice(s,(i*row_count))
                    s += row_count
                
                if self.rules["file_format"] == 1:
                    df.iloc[sl].to_excel(os.path.join(self.rules["output_files_path"], f"{i}.xlsx"), index=False)
                else:
                    df.iloc[sl].to_csv(os.path.join(self.rules["output_files_path"],f"{i}.csv"), index=False, sep=self.rules["seperator_output"])
                
                self.sgn_message.emit(f'{i} th file saved')

            self.sgn_message.emit(f'All files have been saved to {self.rules["output_files_path"]}')
            self.sgn_status.emit(0)

        except Exception as exc:
            self.sgn_message.emit(f'Please check required areas. Error: \n {exc}')
            self.sgn_status.emit(2)