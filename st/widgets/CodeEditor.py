from PyQt5.QtGui import (QFont, QFontMetrics)
from PyQt5.QtWidgets import (QWidget, 
QPushButton, QVBoxLayout, QHBoxLayout, QPlainTextEdit)             
from PyQt5.QtCore import (QThread, pyqtSignal, pyqtSlot)
import sys
import time
from datetime import datetime
from io import StringIO

class CodeEditor(QWidget):

    def __init__(self,  parent):
        super().__init__()
        self.parent = parent        
        # Layouts
        self.layout = QVBoxLayout()
        self.top_layout = QHBoxLayout()
    
        # Code Editor
        self.text_code_editor = QPlainTextEdit()
        self.text_code_editor.setStyleSheet(""" background-color: #162F3E; font-size:16px; color:white;""")
        self.editor_font = QFont()
        self.editor_font.setFamily("Courier")
        self.editor_font.setStyleHint(QFont().Monospace)
        self.editor_font.setPointSize(16)
        self.editor_font.setFixedPitch(True)

        # Set tab metric of self.text_code_editor
        metrics = QFontMetrics(self.editor_font)
        self.text_code_editor.setTabStopWidth(4 * metrics.width(' '))
        self.text_code_editor.setFont(self.editor_font)
        
        # Results
        self.plaintext_results = QPlainTextEdit()
        self.plaintext_results.setFont(self.editor_font)
        self.plaintext_results.setMaximumHeight(300)
        self.plaintext_results.setReadOnly(True)
        self.plaintext_results.setStyleSheet(""" background-color: #162F3E; font-size:13px; color:white;""")  
        
        # Rub button
        self.button_run = QPushButton("Run Code (CTRL+Enter)")
        self.button_run.clicked.connect(self.run_code)
        
        # Kill kernel button
        self.button_kill = QPushButton("Kill Kernel")
        self.button_kill.clicked.connect(self.stop_thread)
        
        # Start kernel button
        self.button_start = QPushButton("Start Kernel")
        self.button_start.clicked.connect(self.start_kernel)
        
        # adding layout and widgets
        self.top_layout.addWidget(self.button_run)
        self.top_layout.addWidget(self.button_start)
        self.top_layout.addWidget(self.button_kill)
        self.layout.addLayout(self.top_layout)
        self.layout.addWidget(self.text_code_editor)
        self.layout.addWidget(self.plaintext_results)

        self.setLayout(self.layout)
        self.start_kernel()

    def keyPressEvent(self, key):
        if self.text_code_editor.hasFocus():
            # Ctrl+Enter Pressed
            if key.key()==16777220:
                self.run_code()

    def start_kernel(self):
        self.button_start.setDisabled(True)
        self.button_kill.setEnabled(True)
        self.button_run.setEnabled(True)
        self.parent.toolbarBox.setDisabled(True)
        self.plaintext_results.clear()
        self.text_code_editor.clear()
        # start thread
        self.codeThread = CodeThread(parent=self)
        self.codeThread.sgn_message.connect(self.update_results)
        self.codeThread.sgn_status.connect(self.update_elements)
        self.codeThread.start()

    def run_code(self):
        cursor = self.text_code_editor.textCursor()
        selected = cursor.selectedText().replace(u"\u2029","\n")
        if len(selected) > 0:
            self.codeThread.code = selected
        else:
            self.codeThread.code = self.text_code_editor.toPlainText()  

    def update_elements(self, status):
        if status==1:
            self.parent.toolbarBox.setDisabled(True)
            self.text_code_editor.setDisabled(True)
        else:
            self.text_code_editor.setEnabled(True)

    def update_results(self, message):
        all_messages = self.plaintext_results.toPlainText() + '\n' + message
        self.plaintext_results.setPlainText(all_messages)
        self.plaintext_results.verticalScrollBar().setValue(self.plaintext_results.verticalScrollBar().maximum())
        
    def stop_thread(self):
        self.codeThread.terminate()
        self.update_elements(0)
        self.update_results("Kernel terminated")
        self.button_start.setEnabled(True)
        self.button_run.setDisabled(True)
        self.button_kill.setDisabled(True)
        self.parent.toolbarBox.setEnabled(True)

class CodeThread(QThread):
    """
    Python kernel thread. This thread starts infinite loop which control self.code variable
    every 0.2 seconds. It emit stdout message with selg.sgn_message.

    """
    sgn_message = pyqtSignal(str)
    sgn_status = pyqtSignal(int)
    
    def __init__(self, parent=None):
        QThread.__init__(self, parent)
        self.code=""

    @pyqtSlot()
    def run(self):
        self.sgn_message.emit(f"Kernel Ready")
        while True:
            time.sleep(0.2)
            try:
                if self.code != "":
                    sys.stdout = buffer = StringIO()
                    self.sgn_status.emit(1)
                    exec(f"{self.code}")
                    output = buffer.getvalue()
                    if len(output) > 0:
                        self.sgn_message.emit(f'{datetime.now()} \n {output}')
                    else:
                        self.sgn_message.emit(f'{datetime.now()} \n OK!')
                    self.sgn_status.emit(0)
                    self.code = ""
            except Exception as exc:
                self.sgn_message.emit(f'Error: \n {exc}')
                self.sgn_status.emit(2)
                self.code = ""