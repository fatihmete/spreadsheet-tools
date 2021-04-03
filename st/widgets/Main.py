from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import ( QMainWindow, QAction, QToolBar, QDesktopWidget)
from PyQt5.QtCore import Qt, QSize
from sys import platform as _platform
from .Home import Home
import os

try:
    # to fix icon on task bar (windows and macosx)
    from PyQt5.QtWinExtras import QtWin
    myappid = 'spreadsheettools.0.0.1'
    QtWin.setCurrentProcessExplicitAppUserModelID(myappid)
except ImportError:
    pass

class MainWindow(QMainWindow):
    def __init__(self, *args, **kwargs):
        super(MainWindow, self).__init__(*args, **kwargs)
        
        #Main Widget
        self.version = "0.0.1"
        self.name = "Spreadsheet Tools"
        self.setWindowIcon(QIcon(os.path.join(os.path.dirname(__file__),"..","icons","sih.png")))
        self.statusBar().showMessage('Ready')
        self.setWindowTitle(f"{self.name} {self.version}")
        self.toolbarBox = QToolBar(self)
        self.toolbarBox.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        self.toolbarBox.setMovable(False)
        
        self.toolbarBox.setIconSize(QSize(80, 40))
        
        self.addToolBar(Qt.TopToolBarArea, self.toolbarBox)
        viewer_toolbar_button = QAction(QIcon(os.path.join(os.path.dirname(__file__),"..", 'icons', "editor.png")),"Excel/CSV Viewer",self)
        viewer_toolbar_button.triggered.connect(lambda : self.set_central_widget("Viewer"))
        self.toolbarBox.addAction(viewer_toolbar_button)

        merge_toolbar_button = QAction(QIcon(os.path.join(os.path.dirname(__file__),"..", 'icons', "merge.png")),"Merge Excel/CSV",self)
        merge_toolbar_button.triggered.connect(lambda : self.set_central_widget("Merger"))
        self.toolbarBox.addAction(merge_toolbar_button)

        split_toolbar_button = QAction(QIcon(os.path.join(os.path.dirname(__file__),"..", 'icons', "split.png")),"Split Excel/CSV",self)
        split_toolbar_button.triggered.connect(lambda : self.set_central_widget("Splitter"))
        self.toolbarBox.addAction(split_toolbar_button)

        excel_reader_toolbar_button = QAction(QIcon(os.path.join(os.path.dirname(__file__),"..", 'icons', "read.png")),"Multiple Excel Reader",self)
        excel_reader_toolbar_button.triggered.connect(lambda : self.set_central_widget("ExcelReader"))
        self.toolbarBox.addAction(excel_reader_toolbar_button)

        excel_filler_toolbar_button = QAction(QIcon(os.path.join(os.path.dirname(__file__),"..", 'icons', "write.png")),"Multiple Excel Writer",self)
        excel_filler_toolbar_button.triggered.connect(lambda : self.set_central_widget("ExcelWriter"))
        self.toolbarBox.addAction(excel_filler_toolbar_button)
    
        vode_editor_button = QAction(QIcon(os.path.join(os.path.dirname(__file__),"..", 'icons', "notebook.png")),"Python Shell",self)
        vode_editor_button.triggered.connect(lambda : self.set_central_widget("CodeEditor"))
        self.toolbarBox.addAction(vode_editor_button)

        about_button = QAction(QIcon(os.path.join(os.path.dirname(__file__),"..", 'icons', "about.png")),"About/Help",self)
        about_button.triggered.connect(self.action_about_button_clicked)
        self.toolbarBox.addAction(about_button)

        exit_button = QAction(QIcon(os.path.join(os.path.dirname(__file__),"..", 'icons', "exit.png")),"Exit",self)
        exit_button.triggered.connect(self.close)
        self.toolbarBox.addAction(exit_button)

        self.setGeometry(800, 400, 1000, 600)
        center_point = QDesktopWidget().availableGeometry().center()
        qtRectangle = self.frameGeometry()
        qtRectangle.moveCenter(center_point)
        self.setCentralWidget(Home(self))

    def set_central_widget(self, widget):
        #Dynamically import widgets
        module =  __import__(f"st.widgets.{widget}", fromlist=[widget])
        instance = getattr(module, widget)(self)
        self.setCentralWidget(instance)
    
    def action_about_button_clicked(self):
        url = "https://github.com/fatihmete/spreadsheet-tools"
        print(_platform)
        if _platform == "linux" or _platform == "linux2":
            os.system(f"xdg-open {url}")
        elif _platform == "darwin":
            os.system(f"open \"\" {url}")
        else: 
            os.system(f"start \"\" {url}")
