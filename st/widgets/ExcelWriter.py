from PyQt5.QtWidgets import (QWidget, QLabel, QLineEdit, QPushButton, 
QTableWidget, QTableWidgetItem, QVBoxLayout, QHeaderView, QMessageBox, 
QHBoxLayout, QComboBox, QHeaderView, QFileDialog, QInputDialog)
from PyQt5.QtCore import Qt, QDir, QThread, pyqtSignal, pyqtSlot
import json
import os
import re
import openpyxl
import pandas as pd

class ExcelWriter(QWidget):
    def __init__(self,  parent):
        super().__init__()
        # Init definations
        self.sheets = set()
        self.rules = []
        self.combos_sheet = {}
        self.input_files_path=""
        self.input_file = ""
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

        labels_table_rules = QLabel("Writing Rules:")
        self.table_rules = QTableWidget()
        self.table_rules.setColumnCount(3)
        self.table_rules.setHorizontalHeaderLabels(["Column Name","Sheet Name","Cell"])
        header = self.table_rules.horizontalHeader()       
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.left_layout.addWidget(labels_table_rules)
        self.left_layout.addWidget(self.table_rules)

        #Right Layout
        button_save_reading_rules = QPushButton("Save Writing Rules")
        button_save_reading_rules.clicked.connect(self.action_save_rules)
        self.right_layout.addWidget(button_save_reading_rules)

        button_load_reading_rules = QPushButton("Load Writing Rules")
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
        # / Layout Edit Buttons

        # Select Files Location
        self.layout_select_output_files_path = QHBoxLayout()
        self.right_layout.addWidget(QLabel("Output Files Path:"))
        self.right_layout.addLayout(self.layout_select_output_files_path)
        self.lineedit_output_files_path = QLineEdit()
        self.lineedit_output_files_path.setReadOnly(True)
        self.button_select_output_files_path = QPushButton("...")
        self.layout_select_output_files_path.addWidget(self.lineedit_output_files_path)
        self.layout_select_output_files_path.addWidget(self.button_select_output_files_path)
        self.button_select_output_files_path.clicked.connect(self.action_select_output_files_path)

        # Select Input File
        self.layout_select_input_file = QHBoxLayout()
        self.right_layout.addWidget(QLabel("Input (data) File (*.xlsx, *.csv):"))
        self.right_layout.addLayout(self.layout_select_input_file)
        self.lineedit_input_file = QLineEdit()
        self.lineedit_input_file.setReadOnly(True)
        self.button_select_input_file = QPushButton("...")
        self.layout_select_input_file.addWidget(self.lineedit_input_file)
        self.layout_select_input_file.addWidget(self.button_select_input_file)
        self.button_select_input_file.clicked.connect(self.action_select_input_file)

        # Select Theme File
        self.layout_select_template_file = QHBoxLayout()
        self.right_layout.addWidget(QLabel("Template File (*.xlsx):"))
        self.right_layout.addLayout(self.layout_select_template_file)
        self.lineedit_template_file = QLineEdit()
        self.lineedit_template_file.setReadOnly(True)
        self.button_select_template_file = QPushButton("...")
        self.layout_select_template_file.addWidget(self.lineedit_template_file)
        self.layout_select_template_file.addWidget(self.button_select_template_file)
        self.button_select_template_file.clicked.connect(self.action_select_template_file)

        self.button_run_rules = QPushButton("Run Rules")
        self.button_run_rules.clicked.connect(self.run_rules)
        self.right_layout.addWidget(self.button_run_rules)

        self.button_stop = QPushButton("Stop Writing")
        self.button_stop.setVisible(False)
        self.button_stop.clicked.connect(self.stop_thread)
        self.right_layout.addWidget(self.button_stop)
        
        self.right_layout.addStretch()
        self.setLayout(self.layout)      

    def action_add_new_rule(self):
        row_position = self.table_rules.rowCount()
        self.table_rules.insertRow(row_position)
        #Combo
        self.combos_sheet[row_position] =  QComboBox()
        for sheet in self.sheets:
            self.combos_sheet[row_position].addItem(sheet)
        
        self.table_rules.setItem(row_position, 0, QTableWidgetItem(""))
        self.table_rules.setCellWidget(row_position, 1, self.combos_sheet[row_position])
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
            index = self.combos_sheet[row_position].findText(rule[1],Qt.MatchFixedString)
            self.combos_sheet[row_position].setCurrentIndex(index)
            self.table_rules.setCellWidget(row_position, 1, self.combos_sheet[row_position])
            self.table_rules.setItem(row_position, 2, QTableWidgetItem(rule[2])) 
            self.table_rules.setItem(row_position, 0, QTableWidgetItem(rule[0]))

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
        file_name, _ = QFileDialog.getSaveFileName(self,"QFileDialog.getSaveFileName()","","Multiple Excel Write Rules Files (*.json)", options=options)
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
        file_name, _ = QFileDialog.getOpenFileName(self,"QFileDialog.getOpenFileName()", "","Multiple Excel Read Rules Files (*.json)", options=options)
        if file_name:
            with open(file_name, 'r', encoding="utf-8") as f:
                read_file = json.load(f)
            #control
            if all([key in read_file.keys() for key in ["type","sheets","rules","input_file","output_files_path","template_file"] ]) and read_file["type"] == "multiple_write":
                self.sheets = set(read_file["sheets"])
                self.rules = read_file["rules"]
                self.lineedit_output_files_path.setText(read_file["output_files_path"])
                self.lineedit_input_file.setText(read_file["input_file"])
                self.lineedit_template_file.setText(read_file["template_file"])
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
            column = self.table_rules.item(row,0)
            sheet_name = self.table_rules.item(row,1)
            cell = self.table_rules.item(row,2)
            self.rules.append([ column.text(), self.combos_sheet[row].currentText(), cell.text()])
        self.output_files_path = self.lineedit_output_files_path.text()
        self.input_file = self.lineedit_input_file.text()
        self.template_file = self.lineedit_template_file.text()
        rules = {"type" : "multiple_write", 
                 "sheets" : [s for s in self.sheets], 
                 "rules" : self.rules,
                 "input_file" : self.input_file,
                 "output_files_path" : self.output_files_path,
                 "template_file" : self.template_file,}
        return rules

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
        file_name, _ = QFileDialog.getOpenFileName(self,"Select input file","","Excel File (*.xlsx);;CSV (*.csv)", options=options)
        if file_name:
            self.lineedit_input_file.setText(file_name)

    def action_select_template_file(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_name, _ = QFileDialog.getOpenFileName(self,"Select template file","","Excel File (*.xlsx)", options=options)
        if file_name:
            self.lineedit_template_file.setText(file_name)   
        
    def run_rules(self):
        self.temp_rules = self.generate_rules()
        sep=""
        if self.temp_rules["input_file"].split(".")[-1] == "csv":
                sep, okPressed = QInputDialog.getText(self, "CSV Seperator","Seperator:", QLineEdit.Normal, ",")

        self.writerThread = Writer(parent=self, rules=self.temp_rules, seperator=sep)
        self.writerThread.sgn_message.connect(self.update_status_bar)
        self.writerThread.sgn_status.connect(self.update_elements)
        self.writerThread.start()
        
    def update_elements(self, status):
        if status==1:
            #Disable all elements
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
        self.writerThread.terminate()
        self.update_elements(0)
        self.update_status_bar("Task terminated")

class Writer(QThread):
    """
    docstring
    """
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
            if rules["input_file"].split(".")[-1] == "csv":
                df = pd.read_csv(rules["input_file"], sep=self.seperator)
            else:
                df = pd.read_excel(rules["input_file"], engine="openpyxl")
            i=1
            for index,row in df.iterrows():
                self.sgn_message.emit(f"Writing {i}.xlsx")
                wb = openpyxl.load_workbook(rules["template_file"], read_only=False, data_only=False)
                for rule in rules["rules"]:
                    work_sheet = wb[rule[1]]
                    work_sheet[rule[2]] = row[rule[0]]
                wb.save(os.path.join(rules["output_files_path"],f"{i}.xlsx"))
                i += 1
            
            self.sgn_message.emit(f'All files have been saved at {rules["output_files_path"]}')
            self.sgn_status.emit(0)
        except Exception as exc:
            self.sgn_message.emit(f'Please check required areas. Error: \n {exc}')
            self.sgn_status.emit(2)