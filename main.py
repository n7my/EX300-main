# This Python file uses the following encoding: utf-8
import sys
from PySide2.QtWidgets import QApplication

# 添加的所需库文件
# from PyQt5 import QtCore, QtGui, QtWidgets
# from PyQt5.QtCore import *
# from PyQt5.QtGui import *
# from PyQt5.QtWidgets import *
# mainwindow.ui Python文件
from main_ui import Ui_Form
from main_logic import Ui_Control
# from AO_calibrate_test import AOCalibrate
# from AI_calibrate_test import AICalibrate
# from DIDO_calibrate_test import DIDOCalibrate
# from excelGenerate import ExcelGenerate

class uWindow(Ui_Control):
    pass

if __name__ == "__main__":
    app = QApplication(sys.argv)

    # 此处调用GUI的程序
    # widgets = QtWidgets.QMainWindow()
    mainWin = uWindow()
    mainWin.show()
    # ui = uWindow()

    # 结束

    sys.exit(app.exec_())



