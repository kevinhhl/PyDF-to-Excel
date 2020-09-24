# -*- coding: utf-8 -*-
"""
PyDF to Excel    : A program that reads PDF files as inputs, and then outputs them
                   in Excel format.
                   
Created on       : Sat Sep 12 12:00:01 2020

@author: kevinhhl
Source code is publicly available on https://github.com/kevinhhl
"""
# Form  generated from reading ui file 'qtGUI_PyDF.ui'
# GUI implementation by: PyQt5 UI code generator 5.9.2
from PyQt5 import QtCore, QtGui, QtWidgets
import IO_Wrapper
import threading
import os

VERSION_NUMBER = "Version 1.1 Rev.2020-09-21a\nSource code is publicly available on https://github.com/kevinhhl"

# [next versions]               - TODO alert windows for invalid inputs
#                               - TODO browse buttons
#                               - 

# Version 1.1 Rev.2020-09-21a : - modify parse_page_instructions to output list of strings, representing range of pages; for camelot to take in as argument
#                               - fork camelot, pass in an instance of Ui_dialog to allow PDF_Handler within to call Ui_dialog.progress.update() as it parses through individual PDF pages
#                               - Fix: allow http/https input to pass through; currently bounded by os.exist()
#

# Version 1.1 Rev.2020-09-20  : - switch to camelot as backend


class Ui_dialog(object):
    
    def setupUi(self, dialog):
        dialog.setObjectName("dialog")
        dialog.setEnabled(True)
        dialog.resize(571, 434)
        self.horizontalLayoutWidget = QtWidgets.QWidget(dialog)
        self.horizontalLayoutWidget.setGeometry(QtCore.QRect(50, 160, 501, 51))
        self.horizontalLayoutWidget.setObjectName("horizontalLayoutWidget")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget)
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label = QtWidgets.QLabel(self.horizontalLayoutWidget)
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.horizontalLayout.addWidget(self.label)
        self.lineEdit = QtWidgets.QLineEdit(self.horizontalLayoutWidget)
        self.lineEdit.setInputMask("")
        self.lineEdit.setText("")
        self.lineEdit.setFrame(True)
        self.lineEdit.setObjectName("lineEdit")
        self.horizontalLayout.addWidget(self.lineEdit)
        spacerItem = QtWidgets.QSpacerItem(80, 20, QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout.addItem(spacerItem)
        self.label_2 = QtWidgets.QLabel(dialog)
        self.label_2.setGeometry(QtCore.QRect(10, 0, 251, 51))
        font = QtGui.QFont()
        font.setPointSize(20)
        font.setBold(True)
        font.setWeight(75)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(dialog)
        self.label_3.setGeometry(QtCore.QRect(10, 30, 191, 31))
        font = QtGui.QFont()
        font.setPointSize(8)
        self.label_3.setFont(font)
        self.label_3.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_3.setObjectName("label_3")
        self.horizontalLayoutWidget_2 = QtWidgets.QWidget(dialog)
        self.horizontalLayoutWidget_2.setGeometry(QtCore.QRect(30, 260, 521, 51))
        self.horizontalLayoutWidget_2.setObjectName("horizontalLayoutWidget_2")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget_2)
        self.horizontalLayout_2.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.label_4 = QtWidgets.QLabel(self.horizontalLayoutWidget_2)
        self.label_4.setObjectName("label_4")
        self.horizontalLayout_2.addWidget(self.label_4)
        self.lineEdit_2 = QtWidgets.QLineEdit(self.horizontalLayoutWidget_2)
        self.lineEdit_2.setEnabled(False)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.horizontalLayout_2.addWidget(self.lineEdit_2)
        spacerItem1 = QtWidgets.QSpacerItem(80, 20, QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_2.addItem(spacerItem1)
        self.label_5 = QtWidgets.QLabel(dialog)
        self.label_5.setEnabled(False)
        self.label_5.setGeometry(QtCore.QRect(200, 70, 351, 40))
        self.label_5.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_5.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_5.setObjectName("label_5")
        self.label_6 = QtWidgets.QLabel(dialog)
        self.label_6.setEnabled(False)
        self.label_6.setGeometry(QtCore.QRect(280, 180, 191, 49))
        self.label_6.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_6.setObjectName("label_6")
        self.horizontalLayoutWidget_3 = QtWidgets.QWidget(dialog)
        self.horizontalLayoutWidget_3.setGeometry(QtCore.QRect(50, 210, 501, 51))
        self.horizontalLayoutWidget_3.setObjectName("horizontalLayoutWidget_3")
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget_3)
        self.horizontalLayout_4.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        spacerItem2 = QtWidgets.QSpacerItem(19, 20, QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_4.addItem(spacerItem2)
        self.label_7 = QtWidgets.QLabel(self.horizontalLayoutWidget_3)
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label_7.setFont(font)
        self.label_7.setObjectName("label_7")
        self.horizontalLayout_4.addWidget(self.label_7)
        self.lineEdit_3 = QtWidgets.QLineEdit(self.horizontalLayoutWidget_3)
        self.lineEdit_3.setInputMask("")
        self.lineEdit_3.setText("")
        self.lineEdit_3.setFrame(True)
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.horizontalLayout_4.addWidget(self.lineEdit_3)
        self.pushButton = QtWidgets.QPushButton(self.horizontalLayoutWidget_3)
        self.pushButton.setObjectName("pushButton")
        self.horizontalLayout_4.addWidget(self.pushButton)
        self.label_8 = QtWidgets.QLabel(dialog)
        self.label_8.setEnabled(False)
        self.label_8.setGeometry(QtCore.QRect(280, 230, 191, 49))
        self.label_8.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_8.setObjectName("label_8")
        self.horizontalLayoutWidget_4 = QtWidgets.QWidget(dialog)
        self.horizontalLayoutWidget_4.setGeometry(QtCore.QRect(0, 360, 561, 61))
        self.horizontalLayoutWidget_4.setObjectName("horizontalLayoutWidget_4")
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget_4)
        self.horizontalLayout_3.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        spacerItem3 = QtWidgets.QSpacerItem(49, 20, QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_3.addItem(spacerItem3)
        self.progressBar = QtWidgets.QProgressBar(self.horizontalLayoutWidget_4)
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        self.horizontalLayout_3.addWidget(self.progressBar)
        spacerItem4 = QtWidgets.QSpacerItem(9, 20, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_3.addItem(spacerItem4)
        self.label_9 = QtWidgets.QLabel(dialog)
        self.label_9.setGeometry(QtCore.QRect(30, 345, 511, 30)) #modifed yPos 360 -> 345; height 20->30 
        self.label_9.setObjectName("label_9")

        self.retranslateUi(dialog)
        QtCore.QMetaObject.connectSlotsByName(dialog)

    def retranslateUi(self, dialog):
        _translate = QtCore.QCoreApplication.translate
        dialog.setWindowTitle(_translate("dialog", "PyDF to Excel - by kevinhh.li"))
        self.label.setText(_translate("dialog", "Input File:"))
        self.label_2.setText(_translate("dialog", "PyDF to Excel"))
        self.label_3.setText(_translate("dialog", "by kevinhhl"))
        self.label_4.setText(_translate("dialog", "Output location:"))
        self.lineEdit_2.setText(_translate("dialog", "[desktop path]"))
        self.label_6.setText(_translate("dialog", "(absolute path to .pdf file)"))
        self.label_7.setText(_translate("dialog", "Pages:"))
        self.pushButton.setText(_translate("dialog", "Convert"))
        self.label_8.setText(_translate("dialog", "(i.e. 1-5,7,10,11-20)"))
        self.label_9.setText(_translate("dialog", "[current inputfile.pdf]"))
        font = QtGui.QFont()
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.label_9.setFont(font)
        self.label_9.hide()
        
# Non-generated blocks {
        self.label_5.setText(_translate("dialog", VERSION_NUMBER))
        self.progressBar.hide() # starts hidden, shows when progress is running
        self.progressBar.setEnabled(False)
        self.pushButton.clicked.connect(self.onClickButtonConvert)#.connect(self.onClickButtonConvert)
        IO_Wrapper.define_output_dir()
        self.lineEdit_2.setText(IO_Wrapper.define_output_dir())#To make this dynamic for user to input

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
# } Non-generated blocks

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    dialog = QtWidgets.QDialog()
    ui = Ui_dialog()
    ui.setupUi(dialog)
    dialog.show()
    sys.exit(app.exec_())
