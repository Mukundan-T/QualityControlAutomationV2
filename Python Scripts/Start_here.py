"""
Authored by James Gaskell

08/14/2024

Edited by:

"""

#import File_name_gen
import Spreadsheet_checks
import Singleton

from File_finder import get_file

import sys, os, qtstylish
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QPixmap, QPainter, QPen, QPainterPath
from PyQt5.QtCore import Qt


class App(QWidget):
    def __init__(self):
        super().__init__()
        self.title = 'Quality Control Automation'
        self.left = 20
        self.top = 80
        self.width = 600
        self.height = 400
        self.setStyleSheet(qtstylish.light())
        self.program = Singleton.Program()
        self.initUI()
  
    def initUI(self):
        self.setWindowTitle(self.title)
        self.move(self.left, self.top)
        self.setFixedSize(self.width, self.height)
        self.setGeometry(self.left, self.top, self.width, self.height)

        self.excel = QLabel(self)
        self.excelPixmap = QPixmap(os.getcwd() + "\\Python Scripts\\excel.png")
        self.smaller_pixmap = self.excelPixmap.scaled(24, 24, Qt.KeepAspectRatio)
        self.excel.setPixmap(self.smaller_pixmap)   
        self.excel.move(15,15)

        Find_file = QPushButton('Search', self)
        Find_file.setGeometry(480, 18, 100,25)
        Find_file.clicked.connect(self.program.get_excel_file())



        Prelim = QPushButton('Preliminary Spreadsheet Checks', self)
        Prelim.setGeometry(self.left, self.top - 20, int(self.width/2 - 20), 30)
        Prelim.clicked.connect(Spreadsheet_checks.run_checks) #Needs to run name checks and date formatting checks (Spreadsheet_checks.py)

        self.showTextbox("")

        self.show()


    def showTextbox(self, text):
        File_text = QTextEdit(text, self)
        File_text.move(65,18)
        File_text.resize(400, 25)
        File_text.setReadOnly(True)

            
        #Prelim = QPushButton('Preliminary Spreadsheet Checks', self)
        #Prelim.setGeometry(self.left, self.top - 60, self.width - 40, 30)
        #Prelim.clicked.connect(Spreadsheet_checks.run_checks) #Needs to run name checks and date formatting checks (Spreadsheet_checks.py)

        
        
        self.update()

class shape():
    def __init__(self,length,position,color):
        self.length=length
        self.position=position
        self.color = color

class rectangle(shape):
    def paintEvent(self, event):

        painter = QPainter(self)

        painter.setPen(QPen(Qt.yellow,  7, Qt.DotLine))

        painter.drawRect(10, 20, 100, 200)

class square(shape):
    def paintEvent(self, event):

        painter = QPainter()

        path = QPainterPath()

        painter.begin(self)

        painter.setRenderHint(QPainter.Antialiasing)

        painter.setPen(Qt.red)

        path.lineTo(20, 12)

        path.lineTo(20, 28)

        path.lineTo(36, 28)
        path.lineTo(36,12)
        path.lineTo(20,12)
        painter.drawPath(path)
        

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    ex = App()
    ex.showTextbox("Hello")
    sys.exit(app.exec_())


main()