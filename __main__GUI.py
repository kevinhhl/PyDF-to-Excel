# -*- coding: utf-8 -*-
"""
PyDF to Excel    : A program that reads PDF files as inputs, and then outputs them
                   in Excel format.
                   
Created on       : Sat Sep 12 12:00:01 2020

@author: kevinhhl
Source code is publicly available on https://github.com/kevinhhl
"""
from PyQt5 import QtWidgets, QtCore
import IO_Wrapper
import threading
import os
from UI_Dialog import Ui_dialog as pyqtGeneratedCode

VERSION_NUMBER = "Version 1.1 \nSource code is publicly available on https://github.com/kevinhhl"

# [next versions]    - TODO alert windows for invalid inputs
#                    - TODO browse buttons


class MainUI(pyqtGeneratedCode):

    def modifyUIs(self,dialog):
        _translate = QtCore.QCoreApplication.translate
        self.label_5.setText(_translate("dialog", VERSION_NUMBER))
        self.label_5.setGeometry(QtCore.QRect(200, 70, 351, 30))
        self.label_9.hide()
        self.progressBar.hide() # starts hidden, shows when progress is running
        self.progressBar.setEnabled(False)
        self.pushButton.clicked.connect(self.onClickButtonConvert)
        IO_Wrapper.define_output_dir()
        self.lineEdit_2.setText(IO_Wrapper.define_output_dir())#To make this dynamic for user to input
        QtCore.QMetaObject.connectSlotsByName(dialog)

        
    def onClickButtonConvert(self):
        src = self.lineEdit.text().strip('\n').strip().strip("\"").strip()
        dest = self.lineEdit_2.text().strip('\n').strip().strip("\"").strip()
        
        if os.path.exists(src) or src.lower().startswith("http"):
            self.pushButton.setEnabled(False)        
            pgInstructions = self.lineEdit_3.text().strip('\n').strip("\"")
            # IO exceptions should be caught at thread level
            threading.Thread(target=IO_Wrapper.process_PDF, args=(self, pgInstructions, src, dest)).start()            
        else:
            alertMsg = "Input file does not exists. If you are trying to enter a web URL, make sure it beings with http or https."
            print(alertMsg)
            #TODO alert window        
            
        
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    dialog = QtWidgets.QDialog()
    ui = MainUI()
    ui.setupUi(dialog)
    ui.modifyUIs(dialog)
    dialog.show()
    sys.exit(app.exec_())
