"""
Authored by James Gaskell

08/14/2024

Edited by:

"""

#import File_name_gen
import Spreadsheet_checks
import Quality_control

import sys
import qtstylish
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSlot

class App(QWidget):
    def __init__(self):
        super().__init__()
        self.title = 'Menu'
        self.left = 20
        self.top = 80
        self.width = 320
        self.height = 200
        self.setStyleSheet(qtstylish.light())
        self.initUI()
  
    def initUI(self):
        self.setWindowTitle(self.title)
        self.move(self.left, self.top)
        self.setFixedSize(self.width, self.height)
        self.setGeometry(self.left, self.top, self.width, self.height)

        Prelim = QPushButton('Preliminary Spreadsheet Checks', self)
        Prelim.setGeometry(self.left, self.top - 60, self.width - 40, 30)
        Prelim.clicked.connect(Spreadsheet_checks.run_checks) #Needs to run name checks and date formatting checks (Spreadsheet_checks.py)

        QualityControl = QPushButton('Quality Control', self)
        QualityControl.setGeometry(self.left, self.top - 20, self.width - 40, 30)
        #QualityControl.clicked.connect() #Needs to run from the Quality_control file

        FinalChecks = QPushButton('Final Checks', self)
        FinalChecks.setGeometry(self.left, self.top + 20, self.width -40, 30)
        #FinalChecks.clicked.connect() #Don't know where this will connect yet

        self.show()

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    ex = App()
    sys.exit(app.exec_())



main()