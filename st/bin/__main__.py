import sys
from PyQt5.QtWidgets import QApplication
from st.widgets.Main import MainWindow

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    app.exec_()
