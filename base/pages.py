#!/usr/bin/python3
# Encoding: utf-8 -*-
# @Time : 2023/12/16 13:15
# @Author : qifan
# @Email : qifan.wang@westwell-lab.com
# @File : pages.py
# @Project : PyQtDemo
import sys
from PyQt5.QtWidgets import QWidget, QMessageBox, QApplication, QDesktopWidget


class MessageWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setGeometry(600, 200, 1600, 1000)
        self.center()

        self.setWindowTitle('计算窗口')
        self.show()

    def center(self):
        qr = self.frameGeometry()  # 获得窗口
        cp = QDesktopWidget().availableGeometry().center()  # 获得屏幕中心点
        qr.moveCenter(cp)  # 显示到屏幕中心
        self.move(qr.topLeft())

    def closeEvent(self, event):
        reply = QMessageBox.question(self, '确认',
                                     "你确定要退出吗？", QMessageBox.Yes |
                                     QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()


app = QApplication(sys.argv)
ex = MessageWidget()
app.exec_()
