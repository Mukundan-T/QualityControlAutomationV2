"""
Authored by James Gaskell

08/14/2024

Edited by:

"""

#import File_name_gen
import Singleton

import sys, os, qtstylish
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QPixmap, QPainter, QPen, QPainterPath, QFont
from PyQt5.QtCore import Qt


class App(QWidget):
    def __init__(self):
        super().__init__()
        self.title = 'Quality Control Automation'
        self.left = 20
        self.top = 80
        self.width = 600
        self.height = 360
        self.setStyleSheet(qtstylish.light())
        self.program = Singleton.Program()
        self.initUI()
  
    def initUI(self):
        self.setWindowTitle(self.title)
        self.move(self.left, self.top)
        self.setFixedSize(self.width, self.height)
        self.setGeometry(self.left, self.top, self.width, self.height)

        self.draw_assets()

        self.show()
    
    def draw_assets(self):

        self.font = QFont()

        self.excel = QLabel(self)
        self.excelPixmap = QPixmap(os.getcwd() + "\\Python Scripts\\excel.png")
        self.smaller_pixmap = self.excelPixmap.scaled(24, 24, Qt.KeepAspectRatio)
        self.excel.setPixmap(self.smaller_pixmap)   
        self.excel.move(20,15)

        self.File_text = QLineEdit("C://", self)
        self.File_text.move(60,18)
        self.font.setPointSize(9)
        self.File_text.setFont(self.font)
        self.File_text.resize(440, 25)
        self.File_text.setReadOnly(True) 

        self.Find_file = QPushButton('Search', self)
        self.font.setPointSize(8)
        self.Find_file.setFont(self.font)
        self.Find_file.setGeometry(510, 20, 63,19)
        self.Find_file.clicked.connect(self.updateTextbox)

        self.font.setPointSize(10)
        self.Prelim_Spreadsheet = QPushButton('Spreadsheet Checks', self)
        self.Prelim_Spreadsheet.setFont(self.font)
        self.Prelim_Spreadsheet.setGeometry(25, 90, 250, 50)
        self.Prelim_Spreadsheet.clicked.connect(self.program.run_spreadsheet_checks) #Needs to run name checks and date formatting checks (Spreadsheet_checks.py)

        self.Prelim_Quality = QPushButton('Preliminary Quality Control', self)
        self.Prelim_Quality.setFont(self.font)
        self.Prelim_Quality.setGeometry(25, 160,  250, 50)
        self.Prelim_Quality.clicked.connect(self.runPrelimQC) #Needs to run name checks and date formatting checks (Spreadsheet_checks.py)

        self.Full_Quality = QPushButton('Full Quality Control', self)
        self.Full_Quality.setFont(self.font)
        self.Full_Quality.setGeometry(25, 230,  250, 50)
        self.Full_Quality.clicked.connect(Singleton.under_construction) #Needs to run name checks and date formatting checks (Spreadsheet_checks.py)

        
        self.show()

    def updateTextbox(self):
        self.program.get_excel_file()
        if self.program.Spreadsheet.filepath != None:
            text = self.program.Spreadsheet.filepath
        self.File_text.setText(text)

        self.show()

    def runPrelimQC(self):
        self.program.get_parent_directory()




def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    ex = App()
    sys.exit(app.exec_())

main()
        