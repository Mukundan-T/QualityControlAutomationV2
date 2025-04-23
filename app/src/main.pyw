"""
Authored by James Gaskell

2/12/2025

Edited by:

"""
import sys
import os
import threading
import time
import easygui
from tkinter import messagebox

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QDesktopServices
from PyQt5.QtCore import QUrl

import files, spreadsheetChecks, preliminaryQC, fileHandler
import matplotlib.colors as mcolors


class SettingsDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Settings")
        self.setGeometry(100, 100, 300, 180)
        self.setWindowFlags(self.windowFlags() ^ QtCore.Qt.WindowContextHelpButtonHint)
        if parent:
            pg = parent.frameGeometry()
            pr = parent.mapToGlobal(pg.topRight())
            self.move(pr.x() - self.width() + 20, pr.y() + 20)

        button_style = """
            QPushButton {
                background-color: rgb(225, 225, 225);
                border-style: outset;
                border-radius: 8px;
                border-width: 1px;
                border-color: rgb(0, 0, 0);
                padding: 1px;
            }
            QPushButton:hover {
                background-color: rgb(205, 205, 205);
            }
        """

        # Factory Reset
        self.reset_factory = QtWidgets.QPushButton("Factory Reset", self)
        self.reset_factory.setGeometry(QtCore.QRect(10, 10, 280, 40))
        self.reset_factory.setFixedHeight(40)
        self.reset_factory.setStyleSheet(button_style)
        self.reset_factory.setToolTip(
            "Restore all error colors to their original defaults\n"
            "(overwrites any custom settings)."
        )
        self.reset_factory.clicked.connect(self.resetFactory)

        # Clear Cached Colors
        self.clear_cached_colors = QtWidgets.QPushButton("Clear Cached Colors", self)
        self.clear_cached_colors.setGeometry(QtCore.QRect(10, 60, 280, 40))
        self.clear_cached_colors.setFixedHeight(40)
        self.clear_cached_colors.setStyleSheet(button_style)
        self.clear_cached_colors.setToolTip(
            "Discard any previously used colors in memory\n"
            "so the program can clear them from the Excel spreadsheet"
        )
        self.clear_cached_colors.clicked.connect(self.clearCachedColors)

        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(self.reset_factory)
        layout.addWidget(self.clear_cached_colors)
        self.setLayout(layout)

    def resetFactory(self):
        # result 0, then close
        self.done(1)

    def clearCachedColors(self):
        # result , then close
        self.done(2)


class GenerateSpreadsheetDialog(QtWidgets.QDialog):
    def __init__(self, parent=None, file_obj=None):
        super().__init__(parent)
        self.file_obj = file_obj
        self.setWindowTitle("Generate Spreadsheet")
        self.setGeometry(100, 100, 350, 300)
        self.setWindowFlags(self.windowFlags() ^ QtCore.Qt.WindowContextHelpButtonHint)
        if parent:
            pg = parent.frameGeometry()
            pr = parent.mapToGlobal(pg.topRight())
            self.move(pr.x() - self.width() + 20, pr.y() + 20)

        # reuse your main‐window button style if available
        self.button_style = getattr(parent, 'button_style', """
            QPushButton {
                background-color: rgb(225, 225, 225);
                border-style: outset;
                border-width: 1px;
                border-radius: 10px;
                border-color: rgb(0, 0, 0);
                padding: 4px;
            }
            QPushButton:hover {
                background-color: rgb(205, 205, 205);
            }
        """)

        # --- Filename input ---
        lbl_file = QtWidgets.QLabel("Filename:", self)
        lbl_file.setGeometry(10, 10, 200, 20)
        self.filenameEdit = QtWidgets.QLineEdit(self)
        self.filenameEdit.setGeometry(10, 35, 330, 20)
        self.filenameEdit.setPlaceholderText("e.g. output.xlsx")

class Ui_MainWindow(object):
    def __init__(self):
        super().__init__()
        self.file = files.ExcelFile(None)
        self.button_style = """
            QPushButton {
                background-color: rgb(225, 225, 225);
                border-style: outset;
                border-width: 1px;
                border-radius: 10px;
                border-color: rgb(0, 0, 0);
                padding: 4px;
            }
            QPushButton:hover {
                background-color: rgb(205, 205, 205);
            }
        """

    def setupUi(self, MainWindow):
        MainWindow.setWindowTitle("Quality Control Automation")
        MainWindow.setFixedSize(682, 401)
        self.centralwidget = QtWidgets.QWidget(MainWindow)

        # Menu frame
        self.menuFrame = QtWidgets.QFrame(self.centralwidget)
        self.menuFrame.setGeometry(20, 80, 311, 291)
        self.menuFrame.setStyleSheet("border: 1px solid black;\nborder-radius: 6px;")

        # Buttons font
        font = QtGui.QFont("Arial", 11)

        # Generate Spreadsheet
        self.GenerateSpreadsheet = QtWidgets.QPushButton("Generate Spreadsheet", self.menuFrame)
        self.GenerateSpreadsheet.setGeometry(20, 50, 271, 61)
        self.GenerateSpreadsheet.setFont(font)
        self.GenerateSpreadsheet.setStyleSheet(self.button_style)
        self.GenerateSpreadsheet.clicked.connect(self.openSpreadsheetMaker)

        # Spreadsheet Checks
        self.SpreadsheetChecks = QtWidgets.QPushButton("Spreadsheet Checks", self.menuFrame)
        self.SpreadsheetChecks.setGeometry(20, 130, 271, 61)
        self.SpreadsheetChecks.setFont(font)
        self.SpreadsheetChecks.setStyleSheet(self.button_style)
        self.SpreadsheetChecks.clicked.connect(self.spreadsheetChecks)

        # Preliminary QC
        self.PrelimQC = QtWidgets.QPushButton("Preliminary QC Checks", self.menuFrame)
        self.PrelimQC.setGeometry(20, 210, 271, 61)
        self.PrelimQC.setFont(font)
        self.PrelimQC.setStyleSheet(self.button_style)
        self.PrelimQC.clicked.connect(self.prelimQC)

        # Menu Header
        self.menuHeader = QtWidgets.QLabel("Menu", self.menuFrame)
        self.menuHeader.setGeometry(10, 10, 291, 20)
        hdr_font = QtGui.QFont("Arial", 12)
        self.menuHeader.setFont(hdr_font)
        self.menuHeader.setAlignment(QtCore.Qt.AlignCenter)
        self.menuHeader.setStyleSheet("border: 0px solid black;")

        # Output frame
        self.outputFrame = QtWidgets.QFrame(self.centralwidget)
        self.outputFrame.setGeometry(340, 180, 321, 191)
        self.outputFrame.setStyleSheet("border: 1px solid black;\nborder-radius: 6px;")

        self.outputHeader = QtWidgets.QLabel("Output", self.outputFrame)
        self.outputHeader.setGeometry(10, 10, 301, 20)
        self.outputHeader.setFont(hdr_font)
        self.outputHeader.setAlignment(QtCore.Qt.AlignCenter)
        self.outputHeader.setStyleSheet("border: 0px solid black;")

        self.outputBox = QtWidgets.QTextBrowser(self.outputFrame)
        self.outputBox.setGeometry(10, 40, 301, 131)
        txt_font = QtGui.QFont("Arial", 9)
        self.outputBox.setFont(txt_font)
        self.outputBox.setStyleSheet(
            "background-color: rgb(255, 255, 255);\n"
            "border-style: outset;\n"
            "border-radius: 10px;\n"
            "border-color: rgb(0, 0, 0);\n"
            "padding: 4px;"
        )

        # Color selector frame
        self.colorFrame = QtWidgets.QFrame(self.centralwidget)
        self.colorFrame.setGeometry(340, 80, 321, 91)
        self.colorFrame.setStyleSheet("border: 1px solid black;\nborder-radius: 6px;")

        self.colorSelectorHeader = QtWidgets.QLabel("Color Selector", self.colorFrame)
        self.colorSelectorHeader.setGeometry(10, 10, 301, 20)
        self.colorSelectorHeader.setFont(hdr_font)
        self.colorSelectorHeader.setAlignment(QtCore.Qt.AlignCenter)
        self.colorSelectorHeader.setStyleSheet("border: 0px solid black;")

        self.colorDisplay = QtWidgets.QPushButton(self.colorFrame)
        self.colorDisplay.setGeometry(20, 45, 21, 21)
        self.colorDisplay.setStyleSheet("background-color: rgb(255, 255, 255)")
        self.colorDisplay.clicked.connect(self.openColorDialog)

        self.errorSelector = QtWidgets.QComboBox(self.colorFrame)
        self.errorSelector.setGeometry(80, 45, 221, 21)
        self.errorSelector.addItems(list(self.file.errorColors.keys()) + list(self.file.failColors.keys()))
        self.errorSelector.currentIndexChanged.connect(self.updateColorSelector)
        self.updateColorSelector()

        self.selectionLine = QtWidgets.QFrame(self.colorFrame)
        self.selectionLine.setGeometry(50, 54, 21, 2)
        self.selectionLine.setFrameShape(QtWidgets.QFrame.HLine)
        self.selectionLine.setFrameShadow(QtWidgets.QFrame.Sunken)

        # File input + icons
        self.fileInput = QtWidgets.QLineEdit("C://", self.centralwidget)
        self.fileInput.setGeometry(65, 30, 485, 25)
        self.fileInput.setFont(QtGui.QFont("Arial", 8))
        self.fileInput.setStyleSheet(
            """
            background-color: rgb(255, 255, 255);
            border: 1px solid black;
            border-radius: 6px;
            padding: 2px;      /* optional: gives a bit of inner space */
            """
        )

        self.excel = QtWidgets.QLabel(self.centralwidget)
        exc_pix = QtGui.QPixmap(os.path.join(os.path.dirname(__file__), "assets\\excel.png"))
        self.excel.setPixmap(exc_pix.scaled(24, 24, QtCore.Qt.KeepAspectRatio))
        self.excel.move(25, 30)

        self.settings = QtWidgets.QPushButton(self.centralwidget)
        set_pix = QtGui.QPixmap(os.path.join(os.path.dirname(__file__), "assets\\settings.webp"))
        self.settings.setIcon(QtGui.QIcon(set_pix))
        self.settings.setIconSize(QtCore.QSize(22, 22))
        self.settings.move(636, 32)
        self.settings.setStyleSheet("QPushButton { border: none; background: transparent; padding: 0; }")
        self.settings.clicked.connect(self.openSettings)

        # Search button
        self.Search = QtWidgets.QPushButton("Search", self.centralwidget)
        self.Search.setGeometry(565, 33, 60, 20)
        self.Search.setFont(QtGui.QFont("Arial", 7))
        self.Search.setStyleSheet("""
            QPushButton {
                background-color: rgb(225, 225, 225);
                border-style: outset;
                border-radius: 8px;
                border-width: 1px;
                border-color: rgb(0, 0, 0);
                padding: 1px;
            }
            QPushButton:hover {
                background-color: rgb(205, 205, 205);
            }
        """)
        self.Search.clicked.connect(self.updateSearchbar)

        MainWindow.setCentralWidget(self.centralwidget)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)


    def updateSearchbar(self):
        file_selected = self.file.setFilePath()
        if file_selected:
            self.fileInput.setText(self.file.filePath)


    def openSettings(self):
        dialog = SettingsDialog(self.centralwidget)
        code = dialog.exec_()
        match code:
            case 0: # Window closed (neither)
                pass
            case 1: # Factory Reset
                self.file = files.ExcelFile(None)
                self.file.resetErrorColors()
                self.file.retrieveErrorColors()
                self.file.clearColorCache()
                self.updateColorSelector()
                self.setupUi(MainWindow)
            case 2: # Clear Cache
                self.file.retrieveErrorColors()
                self.file.clearColorCache()
                self.updateColorSelector()


    def updateColorSelector(self):
        err = self.errorSelector.currentText()
        if err in self.file.errorColors:
            rgb_f = mcolors.hex2color("#" + self.file.errorColors[err][2:])
        else:
            rgb_f = mcolors.hex2color("#" + self.file.failColors[err][2:])
        bg = tuple(int(c * 255) for c in rgb_f)
        self.colorDisplay.setStyleSheet(f"background-color: rgb{bg}")


    def openColorDialog(self):
        color = QtWidgets.QColorDialog.getColor().name()
        err = self.errorSelector.currentText()
        if color != "#000000":
            if err in self.file.errorColors:
                self.file.setErrorColor(err, "FF" + color.upper()[1:])
            else:
                self.file.setFailColor(err, "FF" + color.upper()[1:])
            old = self.colorDisplay.palette().color(self.colorDisplay.backgroundRole()).name().upper()[1:]
            self.file.extendColorCache("FF" + old)
        self.file.writeErrorColors()
        self.updateColorSelector()

    def openSpreadsheetMaker(self):
        dialog = GenerateSpreadsheetDialog(self.centralwidget)
        code = dialog.exec_() # This needs to return 0 if user cancels or 1 if user completes
        match code:
            case 0:
                # Do something here
                pass
            case 1:
                # Do something here
                pass

    def proc_updates(self):
        QtWidgets.QApplication.processEvents()


    def spreadsheetChecks(self):
        if not fileHandler.extract_ext(self.file.filePath):
            messagebox.showerror("Error", "The selected file is not a spreadsheet. The file type must be .xlsx, .csv or .xls")
            return

        self.file.createFileStructure()

        if not self.file.filePath:
            messagebox.showerror("Error", "You must select an excel file before proceeding")
            return
        
        self.outputBox.setText("** Working... **")
        self.proc_updates()
        for i, sheet in enumerate(self.file.sheetList):
            out = ""
            if i < 0:
                out += "\n"
            spreadsheetChecks.check_date_format(sheet)
            spreadsheetChecks.check_duplicate_filenames(sheet)
            spreadsheetChecks.check_location_filename(sheet)
            out += f"{sheet.sheetName} error rate: {round(sheet.getErrorRate(), 2)}%"
            self.outputBox.append(out)
            self.proc_updates()

        success = fileHandler.highlight_errors(self.file)
        if not success:
            messagebox.showerror("Error", "The excel file is open in editor so changes could not be saved")
            return
        
        self.outputBox.append("** Success! **")
        self.proc_updates()
        time.sleep(1)
        self.outputBox.append("** Opening File **")
        self.proc_updates()
        time.sleep(2)
        threading.Thread(target=self._open_excel_file, daemon=True).start()


    def prelimQC(self):
        if not fileHandler.extract_ext(self.file.filePath):
            messagebox.showerror("Error", "The selected file is not a spreadsheet. The file type must be .xlsx, .csv or .xls")
            return
        
        self.file.createFileStructure()

        if not self.file.filePath:
            messagebox.showerror("Error", "You must select an excel file before proceeding")
            return

        parent_dir = easygui.diropenbox()
        self.outputBox.setText("** Working... **")
        self.proc_updates()
        for i, sheet in enumerate(self.file.sheetList):
            out = ""
            if i < 0:
                out += "\n"
            preliminaryQC.check_files(sheet, parent_dir)
            out += f"{sheet.sheetName} failure rate: {round(sheet.getFailureRate(), 2)}%"
            self.outputBox.append(out)
            self.proc_updates()

        try:
            self.file.updateDataFrames()
            success = fileHandler.write_excelfile(self.file)
            if not success:
                messagebox.showerror("Error", "The excel file is open in editor so changes could not be saved")
                return

            success = fileHandler.highlight_errors(self.file)
            if not success:
                messagebox.showerror("Error", "Could not highlight errors in Excel file")
                return

            self.outputBox.append("** Success! **")
            self.proc_updates()
            time.sleep(1)
            self.outputBox.append("** Opening File **")
            self.proc_updates()
            time.sleep(2)
            threading.Thread(target=self._open_excel_file, daemon=True).start()

        except KeyError:
            messagebox.showerror("Error", "The column headers are not in the expected format!")


    def _open_excel_file(self):
        path = self.file.filePath
        if not path or not os.path.exists(path):
            return
        try:
            os.startfile(path)
        except AttributeError:
            QDesktopServices.openUrl(QUrl.fromLocalFile(path))


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    app.exec_()
