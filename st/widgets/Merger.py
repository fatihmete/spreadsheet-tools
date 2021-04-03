from PyQt5.QtWidgets import (QWidget, QLabel, QLineEdit, QPushButton, 
QVBoxLayout, QHBoxLayout, QComboBox, QFileDialog)                          
from PyQt5.QtCore import QDir, QThread, pyqtSignal, pyqtSlot
from PyQt5.QtGui import QFont
import pandas as pd
from glob import glob
import os
import re

class Merger(QWidget):

    def __init__(self,  parent):
        super().__init__()
        self.parent = parent

        # Select Input Files Location
        self.layout = QVBoxLayout()
        self.layout_select_input_files_path = QHBoxLayout()
        self.layout.addWidget(QLabel("Input Files Path (The location of the files you want to merge):"))
        self.layout.addLayout(self.layout_select_input_files_path)
        self.lineedit_input_files_path = QLineEdit()
        self.lineedit_input_files_path.setReadOnly(True)
        self.button_select_input_files_path = QPushButton("...")
        self.layout_select_input_files_path.addWidget(self.lineedit_input_files_path)
        self.layout_select_input_files_path.addWidget(self.button_select_input_files_path)
        self.button_select_input_files_path.clicked.connect(self.action_select_input_files_path)

        # Select Output File
        self.layout_select_output_file = QHBoxLayout()
        self.layout.addWidget(QLabel("Output File (*.xlsx, *.csv):"))
        self.layout.addLayout(self.layout_select_output_file)
        self.lineedit_output_file = QLineEdit()
        self.lineedit_output_file.setReadOnly(True)
        self.button_select_output_file = QPushButton("...")
        self.layout_select_output_file.addWidget(self.lineedit_output_file)
        self.layout_select_output_file.addWidget(self.button_select_output_file)
        self.button_select_output_file.clicked.connect(self.action_select_output_file)

        # File Types
        self.layout_select_file_format = QVBoxLayout()
        self.layout.addWidget(QLabel("Select format of files in input files path (Only selected format files will be merge ): "))
        self.layout.addLayout(self.layout_select_file_format)
        self.combo_format = QComboBox()
        self.combo_format.addItems(["Only *.xlsx files", "Only *.csv", "Only *.txt files", "Mix type (all files)"])
        self.layout_select_file_format.addWidget(self.combo_format)
        self.layout.addWidget(QLabel("If you select mix type, files other than *.xlsx will be assumed to be *.csv/*.txt ."))

        # CSV Seperator
        self.layout.addWidget(QLabel("CSV/Text Seperator (If you want to merge csv/txt files):"))
        self.lineedit_csv_sep = QLineEdit(",")
        self.layout.addWidget(self.lineedit_csv_sep)

        #Merge Button
        self.button_merge_files = QPushButton("Merge Files")
        self.layout.addWidget(self.button_merge_files)
        self.button_merge_files.clicked.connect(self.merge_files)

        #Stop Button
        self.button_stop_merge = QPushButton("Stop Merge")
        self.button_stop_merge.clicked.connect(self.stop_thread)
        self.button_stop_merge.setVisible(False)
        self.layout.addWidget(self.button_stop_merge)         

        self.layout.addStretch()
        self.setLayout(self.layout)
        
    def action_select_input_files_path(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        options |= QFileDialog.DontUseCustomDirectoryIcons
        dialog = QFileDialog()
        dialog.setOptions(options)
        dialog.setFilter(dialog.filter() | QDir.Hidden)
        path  = dialog.getExistingDirectory(self, 'Select directory', options=options)
        if path:
            self.lineedit_input_files_path.setText(path)
    
    def action_select_output_file(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_name, _ = QFileDialog.getSaveFileName(self,"QFileDialog.getSaveFileName()","","CSV (*.csv);;TXT (*.txt);;Excel File (*.xlsx)", options=options)
        if file_name:
            if file_name.split(".")[-1] not in ["csv", "xlsx"]:
                file_name = file_name + "." + re.findall("\*\.(.*)\)", _)[0] 
            self.lineedit_output_file.setText(file_name)
    
    def merge_files(self):
        formats = {
            0 : "*.xlsx",
            1 : "*.csv",
            2 : "*.txt",
            3 : "*."
        }
        rules = {}
        rules["input_files_path"] = self.lineedit_input_files_path.text()
        rules["output_file"] = self.lineedit_output_file.text()
        rules["seperator"] = self.lineedit_csv_sep.text()
        rules["file_format"] = formats[self.combo_format.currentIndex()]

        self.mergerThread = MergerThread(rules=rules)
        self.mergerThread.sgn_message.connect(self.update_status_bar)
        self.mergerThread.sgn_status.connect(self.update_elements)
        self.mergerThread.start()

    def update_elements(self, status):
        if status==1:
            self.parent.toolbarBox.setDisabled(True)
            self.button_merge_files.setDisabled(True)
            self.button_stop_merge.setVisible(True)
        else:
            self.button_stop_merge.setVisible(False)
            self.parent.toolbarBox.setEnabled(True)
            self.button_merge_files.setEnabled(True)

    def update_status_bar(self, message):
        self.parent.statusBar().showMessage(message)
        
    def stop_thread(self):
        self.mergerThread.terminate()
        self.update_elements(0)
        self.update_status_bar("Task terminated")

class MergerThread(QThread):
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
            dfs = []
            files = glob(os.path.join(self.rules["input_files_path"], self.rules["file_format"]))
            
            if len(files)==0:
                raise Exception(f"""There is no {self.rules["file_format"]} file in {self.rules["input_files_path"]}""")
            
            for file in files:
                self.sgn_message.emit(f"Reading {file}")
                full_file_name = os.path.join(self.rules["input_files_path"], file)
                if file.split(".")[-1] == "xlsx":
                    dfs.append(pd.read_excel(full_file_name, engine="openpyxl"))
                else:
                    dfs.append(pd.read_csv(full_file_name, sep=self.rules["seperator"]))

            df = pd.concat(dfs)
            self.sgn_message.emit(f"Saving output file")

            if self.rules["output_file"].split(".")[-1] == "xlsx":
                df.to_excel(self.rules["output_file"], index=False)
            else:
                df.to_csv(self.rules["output_file"], index=False, sep=self.rules["seperator"])

            self.sgn_message.emit(f'All files have been merged to {self.rules["output_file"]}')
            self.sgn_status.emit(0)

        except Exception as exc:
            self.sgn_message.emit(f'Please check required areas. Error: \n {exc}')
            self.sgn_status.emit(2)