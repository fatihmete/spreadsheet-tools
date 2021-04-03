
from PyQt5.QtWidgets import QWidget,  QLabel, QVBoxLayout 
                                                       
class Home(QWidget):
    def __init__(self,  parent):
        super().__init__()
        self.parent = parent
        self.layout = QVBoxLayout()
        self.label_about = QLabel(f"""Spreadsheet in Hand version {self.parent.version}
        \nPlease press the buttons top to access the feature you want. """)
        self.layout.addWidget(self.label_about)
        self.layout.addStretch()
        self.setLayout(self.layout)
