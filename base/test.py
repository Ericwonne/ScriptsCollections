#!/usr/bin/python3
# Encoding: utf-8 -*-
# @Time : 2023/12/16 13:44
# @Author : qifan
# @Email : qifan.wang@westwell-lab.com
# @File : test.py
# @Project : PyQtDemo

# import sys
# from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QLineEdit, QPushButton, QVBoxLayout, QWidget
# import pandas as pd
#
# class MyWindow(QMainWindow):
#     def __init__(self):
#         super().__init__()
#
#         self.initUI()
#
#     def initUI(self):
#         # 创建组件
#         self.label = QLabel('请输入内容:')
#         self.line_edit = QLineEdit(self)
#         self.export_button = QPushButton('导出到Excel', self)
#         self.export_button.clicked.connect(self.export_to_excel)
#
#         # 创建布局
#         layout = QVBoxLayout()
#         layout.addWidget(self.label)
#         layout.addWidget(self.line_edit)
#         layout.addWidget(self.export_button)
#
#         # 创建主窗口部件
#         container = QWidget(self)
#         container.setLayout(layout)
#         self.setCentralWidget(container)
#
#         self.setGeometry(300, 300, 400, 200)
#         self.setWindowTitle('PyQt Keyboard Input to Excel')
#         self.show()
#
#     def export_to_excel(self):
#         # 获取用户输入
#         user_input = self.line_edit.text()
#
#         # 将输入数据存储到 DataFrame
#         data = {'用户输入': [user_input]}
#         df = pd.DataFrame(data)
#
#         # 导出到 Excel 文件
#         df.to_excel('user_input.xlsx', index=False)
#         print('数据已导出到 user_input.xlsx 文件')
#
# if __name__ == '__main__':
#     app = QApplication(sys.argv)
#     window = MyWindow()
#     sys.exit(app.exec_())

# import pandas as pd
#
# json_data = {
#     "product": {
#         "product_name": "computer",
#         "price": 1200
#     },
#     "store": {
#         "store_number": 77,
#         "store_city": "London"
#     },
#     "time": {
#         "opening_hours": "8-18",
#         "opening_days": "Mon-Fri"
#     }
# }
#
# # Json to DataFrame
# df = pd.json_normalize(json_data)
#
# # DataFrame to Excel
# excel_filename = 'json_data_to_excel.xlsx'
# df.to_excel(excel_filename, index=False)

def generate_result_list(A, B):
    C = []
    for a_value in A:
        sublist = []
        while a_value != 0:
            if B and B[0] <= a_value:
                sublist.append(B.pop(0))
                a_value -= sublist[-1]
            else:
                sublist.append(a_value)
                B[0] -= a_value
                a_value = 0
        C.append(sublist)
    return C


# Example usage
# A, B = [30, 15, 45], [13, 24, 28, 20, 5]
A, B = [30, 30, 30, 40], [13, 24, 28, 20, 25,10, 10]
result_list = generate_result_list(A, B)

print(result_list)
