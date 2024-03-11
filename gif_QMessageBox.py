from PyQt5.QtWidgets import QApplication, QMessageBox, QLabel, QVBoxLayout, QWidget
from PyQt5.QtGui import QMovie
from PyQt5 import QtCore
from PyQt5.QtCore import QSize
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
class MyMessageBox(QMessageBox):
    def __init__(self, parent=None):
        super().__init__(parent)

        # 创建一个QLabel控件，用于显示gif图片
        self.label = QLabel(self)

        # 创建一个QMovie对象，并设置gif图片
        self.movie = QMovie('D:/EX300检测软件/gitlab/新建文件夹/beast/ET1600_even.gif')
        # 修改gif图片的尺寸
        size = QSize(400, 300)
        self.movie.setScaledSize(size)
        self.label.setMovie(self.movie)

        self.label.setAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter)
        # self.label.setAlignment(QtCore.Qt.AlignVCenter)
        self.layout().addWidget(self.label)

        # 创建一个QLabel控件，用于显示文字
        self.label_text = QLabel(self)
        self.label_text.setText('ET1600通道指示灯是否如图所示闪烁？')
        self.label_text.setAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
        self.layout().addWidget(self.label_text)

        # 设置QMessageBox的大小和标题
        self.resize(400, 400)
        self.setWindowTitle('提示')

    def showEvent(self, event):
        # 在QMessageBox显示时启动QMovie动画
        self.movie.start()

    def hideEvent(self, event):
        # 在QMessageBox隐藏时停止QMovie动画
        self.movie.stop()