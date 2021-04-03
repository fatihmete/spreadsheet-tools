from PyQt5.QtWidgets import (QWidget, QLabel, QLineEdit, QPushButton, 
QTableWidget,QTableWidgetItem, QVBoxLayout, QHeaderView, QMessageBox, 
QHBoxLayout, QComboBox, QHeaderView, QFileDialog, QInputDialog)
from PyQt5.QtCore import Qt, QDir, QThread, pyqtSignal, pyqtSlot
import json
import os
import re
import openpyxl
import pandas as pd

class ExcelReader(QWidget):
    def __init__(self,  parent):
        super().__init__()
        # Init definations
        self.sheets = set()
        self.rules = []
        self.combos_sheet = {}
        self.input_files_path=""
        self.ouput_file = ""

        self.parent = parent
        self.layout = QHBoxLayout()
        self.left_layout = QVBoxLayout()
        self.right_layout = QVBoxLayout()

        self.layout.addLayout(self.left_layout)
        self.layout.addLayout(self.right_layout)
        #Left Layout
        self.layout_rules_buttons = QHBoxLayout()
        self.left_layout.addLayout(self.layout_rules_buttons)
        button_add_new_rule = QPushButton("Add New Rule")
        button_add_new_rule.clicked.connect(self.action_add_new_rule)
        button_delete_rule = QPushButton("Delete Selected Rule")
        button_delete_rule.clicked.connect(self.action_delete_rule)
        
        self.layout_rules_buttons.addWidget(button_add_new_rule)
        self.layout_rules_buttons.addWidget(button_delete_rule)

        labels_table_rules = QLabel("Reading Rules:")
        self.table_rules = QTableWidget()
        self.table_rules.setColumnCount(3)
        self.table_rules.setHorizontalHeaderLabels(["Sheet Name","Cell","Column Name"])
        header = self.table_rules.horizontalHeader()       
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        self.left_layout.addWidget(labels_table_rules)
        self.left_layout.addWidget(self.table_rules)

        #Right Layout
        button_save_reading_rules = QPushButton("Save Reading Rules")
        button_save_reading_rules.clicked.connect(self.action_save_rules)
        self.right_layout.addWidget(button_save_reading_rules)

        button_load_reading_rules = QPushButton("Load Reading Rules")
        button_load_reading_rules.clicked.connect(self.action_load_rules)
        self.right_layout.addWidget(button_load_reading_rules)

        label_sheets = QLabel("Sheets:")

        self.table_sheets = QTableWidget()
        self.table_sheets.setColumnCount(1)
        self.table_sheets.setHorizontalHeaderLabels(["Sheet Name"])
        self.table_sheets.setEditTriggers(QTableWidget.NoEditTriggers)

        header = self.table_sheets.horizontalHeader()       
        header.setSectionResizeMode(0, QHeaderView.Stretch)

        self.right_layout.addWidget(label_sheets)
        self.right_layout.addWidget(self.table_sheets)

        # Layout Edit Buttons
        self.layout_add_sheet = QHBoxLayout()
        self.right_layout.addLayout(self.layout_add_sheet)

        self.lineedit_new_sheet_name = QLineEdit()
        button_add_new_sheet = QPushButton("Add Sheet")
        button_add_new_sheet.clicked.connect(self.action_add_sheet)

        button_delete_sheet = QPushButton("Delete Selected Sheet")
        button_delete_sheet.clicked.connect(self.action_delete_sheet)

        self.layout_add_sheet.addWidget(self.lineedit_new_sheet_name)
        self.layout_add_sheet.addWidget(button_add_new_sheet)
        self.layout_add_sheet.addWidget(button_delete_sheet)


        # Select Files Location
        self.layout_select_folder = QHBoxLayout()
        self.right_layout.addWidget(QLabel("Input Files Path (only read *.xlsx files):"))
        self.right_layout.addLayout(self.layout_select_folder)
        self.lineedit_folder_path = QLineEdit()
        self.lineedit_folder_path.setReadOnly(True)
        self.button_select_input_files_path = QPushButton("...")
        self.layout_select_folder.addWidget(self.lineedit_folder_path)
        self.layout_select_folder.addWidget(self.button_select_input_files_path)
        self.button_select_input_files_path.clicked.connect(self.action_select_input_files_path)

        # Select Output File
        self.layout_select_output_file = QHBoxLayout()
        self.right_layout.addWidget(QLabel("Output File (*.xlsx, *.csv):"))
        self.right_layout.addLayout(self.layout_select_output_file)
        self.lineedit_output_file = QLineEdit()
        self.lineedit_output_file.setReadOnly(True)
        self.button_select_output_file = QPushButton("...")
        self.layout_select_output_file.addWidget(self.lineedit_output_file)
        self.layout_select_output_file.addWidget(self.button_select_output_file)
        self.button_select_output_file.clicked.connect(self.action_select_output_file)

        self.button_run_rules = QPushButton("Run Rules")
        self.button_run_rules.clicked.connect(self.run_rules)
        self.right_layout.addWidget(self.button_run_rules)

        self.button_stop = QPushButton("Stop Reading")
        self.button_stop.setVisible(False)
        self.button_stop.clicked.connect(self.stop_thread)
        self.right_layout.addWidget(self.button_stop)       

        self.right_layout.addStretch()

        self.setLayout(self.layout)

    def action_add_new_rule(self):
        row_position = self.table_rules.rowCount()
        self.table_rules.insertRow(row_position)
        #Combo
        self.combos_sheet[row_position] = QComboBox()
        for sheet in self.sheets:
            self.combos_sheet[row_position].addItem(sheet)
        
        self.table_rules.setCellWidget(row_position, 0, self.combos_sheet[row_position])
        self.table_rules.setItem(row_position, 1, QTableWidgetItem(""))
        self.table_rules.setItem(row_position, 2, QTableWidgetItem(""))

    def restore_and_load_rules(self):
        #Remove all existing rules in table
        self.table_rules.setRowCount(0)
        self.combos_sheet = {}
        for rule in self.rules:
            row_position = self.table_rules.rowCount()
            self.table_rules.insertRow(row_position)
            self.combos_sheet[row_position] =  QComboBox()
            for sheet in self.sheets:
                self.combos_sheet[row_position].addItem(sheet)
            index = self.combos_sheet[row_position].findText(rule[0],Qt.MatchFixedString)
            self.combos_sheet[row_position].setCurrentIndex(index)
            self.table_rules.setCellWidget(row_position, 0, self.combos_sheet[row_position])
            self.table_rules.setItem(row_position, 1, QTableWidgetItem(rule[1])) 
            self.table_rules.setItem(row_position, 2, QTableWidgetItem(rule[2]))

    def action_add_sheet(self):
        sheet_name = self.lineedit_new_sheet_name.text()
        self.sheets.add(sheet_name)
        self.action_load_sheets()
        self.action_update_sheets_combos()
    
    def action_delete_sheet(self):
        current_row = self.table_sheets.currentRow()
        if current_row != -1:
            sheet_name = self.table_sheets.item(current_row,0).text()
            self.sheets.remove(sheet_name)
            self.action_load_sheets()
            self.action_update_sheets_combos()
        
    def action_load_sheets(self):
        self.table_sheets.setRowCount(0)
        for i, sheet in enumerate(self.sheets):
            self.table_sheets.insertRow(i)
            self.table_sheets.setItem(i, 0 ,QTableWidgetItem(sheet))

    def action_update_sheets_combos(self):
        for i in range(len(self.combos_sheet)):
            current = self.combos_sheet[i].currentText()
            self.combos_sheet[i].clear()
            for sheet in self.sheets:
                self.combos_sheet[i].addItem(sheet)
            if current in self.sheets:
                index = self.combos_sheet[i].findText(current,Qt.MatchFixedString)
                self.combos_sheet[i].setCurrentIndex(index)
            else:
                self.combos_sheet[i].setCurrentIndex(0)

    def action_delete_rule(self):
        current_row = self.table_rules.currentRow()
        if current_row != -1:
            self.table_rules.removeRow(current_row)

    def action_save_rules(self):
        rules = self.generate_rules()
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_name, _ = QFileDialog.getSaveFileName(self,"QFileDialog.getSaveFileName()", "", "Multiple Excel Read Rules Files (*.json)", options=options)
        if file_name:
            if file_name.split(".")[-1] !="json":
                file_name = file_name + ".json"
            with open(file_name, 'w', encoding='utf-8') as f:
                json.dump(rules, f, ensure_ascii=False, indent=4)
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText("Success")
        msg.setInformativeText(f'All rules have been saved at {file_name}')
        msg.setWindowTitle("Success")
        msg.exec_()
    
    def action_load_rules(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_name, _ = QFileDialog.getOpenFileName(self, "QFileDialog.getOpenFileName()", "", "Multiple Excel Read Rules Files (*.json)", options=options)
        if file_name:
            
            with open(file_name, 'r', encoding="utf-8") as f:
                read_file = json.load(f)
            #control
            if all([key in read_file.keys() for key in ["type","sheets","rules","ouput_file","input_files_path"] ]) and read_file["type"] == "multiple_read":
                self.sheets = set(read_file["sheets"])
                self.rules = read_file["rules"]
                self.lineedit_folder_path.setText(read_file["input_files_path"])
                self.lineedit_output_file.setText(read_file["ouput_file"])
                self.action_load_sheets()
                self.restore_and_load_rules()

                msg = QMessageBox()
                msg.setIcon(QMessageBox.Information)
                msg.setText("Success")
                msg.setInformativeText(f'All rules have been loaded from {file_name}')
                msg.setWindowTitle("Success")
                msg.exec_()

            else:
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Critical)
                msg.setText("Error")
                msg.setInformativeText('Not compatible reading file, please select compitable file')
                msg.setWindowTitle("Error")
                msg.exec_()

    def generate_rules(self):
        row_position = self.table_rules.rowCount()
        self.rules = []
        for row in range(row_position):
            sheet_name = self.table_rules.item(row,0)
            cell = self.table_rules.item(row,1)
            column = self.table_rules.item(row,2)
            self.rules.append([self.combos_sheet[row].currentText(), cell.text(), column.text()])
        self.input_files_path = self.lineedit_folder_path.text()
        self.ouput_file = self.lineedit_output_file.text()
        rules = {"type" : "multiple_read", 
                 "sheets" : [s for s in self.sheets], 
                 "rules" : self.rules,
                 "ouput_file" : self.ouput_file,
                 "input_files_path" : self.input_files_path}
        return rules

    def action_select_input_files_path(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        options |= QFileDialog.DontUseCustomDirectoryIcons
        dialog = QFileDialog()
        dialog.setOptions(options)
        dialog.setFilter(dialog.filter() | QDir.Hidden)
        path  = dialog.getExistingDirectory(self, 'Select directory', options=options)
        if path:
            self.lineedit_folder_path.setText(path)

    def action_select_output_file(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_name, _ = QFileDialog.getSaveFileName(self,"QFileDialog.getSaveFileName()","","Excel File (*.xlsx);;CSV (*.csv)", options=options)
        print(_)
        if file_name:
            if file_name.split(".")[-1] not in ["csv", "xlsx"]:
                file_name = file_name + "." + re.findall("\*\.(.*)\)", _)[0] 
            self.lineedit_output_file.setText(file_name)
    def run_rules(self):
        self.temp_rules = self.generate_rules()
        sep=""
        if self.temp_rules["ouput_file"].split(".")[-1] == "csv":
                sep, okPressed = QInputDialog.getText(self, "CSV Seperator","Seperator:", QLineEdit.Normal, ",")

        self.readerThread = Reader(parent=self, rules=self.temp_rules, seperator=sep)
        self.readerThread.sgn_message.connect(self.update_status_bar)
        self.readerThread.sgn_status.connect(self.update_elements)
        self.readerThread.start()

    def update_elements(self, status):
        if status==1:
            self.parent.toolbarBox.setDisabled(True)
            self.button_run_rules.setDisabled(True)
            self.button_stop.setVisible(True)
        else:
            self.button_stop.setVisible(False)
            self.parent.toolbarBox.setEnabled(True)
            self.button_run_rules.setEnabled(True)

    def update_status_bar(self, message):
        self.parent.statusBar().showMessage(message)
        
    def stop_thread(self):
        self.readerThread.terminate()
        self.update_elements(0)
        self.update_status_bar("Task terminated")

class Reader(QThread):
    sgn_message = pyqtSignal(str)
    sgn_status = pyqtSignal(int)
    
    def __init__(self, parent=None, rules=None, seperator=None):
        QThread.__init__(self, parent)
        self.rules = rules
        self.seperator = seperator
    
    @pyqtSlot()
    def run(self):
        self.sgn_status.emit(1)
        self.sgn_message.emit(f"Starting")
        try:
            rules = self.rules
            columns = [rule[2] for rule in rules["rules"]]
            path = rules["input_files_path"]
            data = []
            for excel_file in os.listdir(path):
                full_file_name = os.path.join(path,excel_file)
                if excel_file.split(".")[-1] in ["xlsm", "xlsx", "xltx", "xltm"]:
                    self.sgn_message.emit(f"Reading {excel_file}")
                    work_book = openpyxl.load_workbook(full_file_name, data_only=True)
                    file_data = []
                    for rule in rules["rules"]:
                        work_sheet = work_book[rule[0]]
                        file_data.append(work_sheet[rule[1]].value)
                    data.append(file_data)  
            df = pd.DataFrame(data, columns=columns)

            if rules["ouput_file"].split(".")[-1] == "csv":
                df.to_csv(rules["ouput_file"], index=False, sep=self.seperator)
            else:
                df.to_excel(rules["ouput_file"], index=False)

            self.sgn_message.emit(f'All data have been saved to {rules["ouput_file"]}')
            self.sgn_status.emit(0)

        except Exception as exc:
            self.sgn_message.emit(f'Please check required areas. Error: \n {exc}')
            self.sgn_status.emit(2)