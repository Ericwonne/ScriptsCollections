#!/usr/bin/python3
# Encoding: utf-8 -*-
# @Time : 2023/12/16 22:54
# @Author : qifan
# @Email : qifan.wang@westwell-lab.com
# @File : plain_dos_page.py
# @Project : PyQtDemo

import os
import datetime
from loguru import logger
from openpyxl import workbook, styles
from openpyxl.formula import Tokenizer
from dateutil import parser

logger.remove()  # remove log to std
logger.add('record.log')


def get_struct_from_input():
    logger.debug(
        "结算单输入信息收集，默认拉板对回板为一对多\n"
    )
    unit_price = float(input("请输入拉板/回板单价，默认全程使用该价格进行计算！\n"))

    result_structure = list()
    while True:
        from_date = input("请输入【拉板】日期，格式示例：20230130（2023年1月30日），键入字母“e”退出程序！\n")
        if from_date.lower() == 'e':
            break
        from_date = parser.parse(from_date).date()
        from_amount = int(input("请输入【拉板】块数\n"))

        to_index = 1
        to_structure = []
        while True:
            to_date1 = input(f"请输入第{to_index}个【回板】日期，键入字母e结束本次输入\n")
            if to_date1.lower() == 'e':
                break
            to_date1 = parser.parse(to_date1).date()
            to_amount1 = int(input(f"请输入第{to_index}次【回板】块数\n"))
            to_index += 1

            to_structure.append([to_date1, to_amount1])
        result_structure.append([from_date, from_amount, to_structure])

    logger.debug(f"单价：{unit_price}")
    logger.debug(f"结构：{result_structure}")

    return unit_price, result_structure


def form_xlsx_file(price, structure):
    workbook_ = workbook.Workbook()
    sheet_ = workbook_.active

    # region manipulate the cells
    center_alignment = styles.Alignment(horizontal='center', vertical='center')

    # header part
    sheet_.merge_cells('C3:D3')
    sheet_['C3'], sheet_['C4'], sheet_['D4'] = "拉板", "日期", "块数"
    sheet_['C3'].alignment = center_alignment

    sheet_.merge_cells('E3:F3')
    sheet_['E3'], sheet_['E4'], sheet_['F4'] = "回板", "日期", "块数"
    sheet_['E3'].alignment = center_alignment

    sheet_.merge_cells('G3:G4')
    sheet_.merge_cells('H3:H4')
    sheet_.merge_cells('I3:I4')
    sheet_['G3'], sheet_['H3'], sheet_['I3'] = '天数', '单价', '总价'
    sheet_['G3'].alignment, sheet_['H3'].alignment, sheet_['I3'].alignment = \
        center_alignment, center_alignment, center_alignment

    # body part
    start_row_, start_column = 5, 3
    for item in structure:
        depth_ = len(item[-1])
        if depth_ == 1:
            frm_date_cell = sheet_.cell(row=start_row_, column=start_column, value=item[0])  # 拉板日期
            sheet_.cell(row=start_row_, column=start_column + 1, value=item[1])  # 拉板块数
            to_date_cell = sheet_.cell(row=start_row_, column=start_column + 2, value=item[-1][0][0])  # 回板日期
            to_amnt_cell = sheet_.cell(row=start_row_, column=start_column + 3, value=item[-1][0][1])  # 回板块数
            day_cnt_cell = sheet_.cell(row=start_row_, column=start_column + 4,
                                       value=f"={to_date_cell.coordinate}-{frm_date_cell.coordinate}")  # 天数
            uni_prc_cell = sheet_.cell(row=start_row_, column=start_column + 5, value=price)  # 单价
            sheet_.cell(row=start_row_, column=start_column + 6,
                        value=f'={to_amnt_cell.coordinate}*{day_cnt_cell.coordinate}*{uni_prc_cell.coordinate}')  # 总价

        else:
            sheet_.merge_cells(start_row=start_row_, start_column=start_column, end_row=start_row_ + depth_ - 1,
                               end_column=start_column)
            sheet_.merge_cells(start_row=start_row_, start_column=start_column + 1, end_row=start_row_ + depth_ - 1,
                               end_column=start_column + 1)
            frm_date_cell = sheet_.cell(row=start_row_, column=start_column, value=item[0])  # 拉板日期
            frm_date_cell.alignment = center_alignment
            sheet_.cell(row=start_row_, column=start_column + 1, value=item[1]).alignment = center_alignment  # 拉板块数

            for idx, sub_item in enumerate(item[-1]):
                row = start_row_ + idx
                to_date_cell = sheet_.cell(row=row, column=start_column + 2, value=sub_item[0])  # 回板日期
                to_amnt_cell = sheet_.cell(row=row, column=start_column + 3, value=sub_item[1])  # 回板块数
                day_cnt_cell = sheet_.cell(row=row, column=start_column + 4,
                                           value=f"={to_date_cell.coordinate}-{frm_date_cell.coordinate}")  # 天数
                uni_prc_cell = sheet_.cell(row=row, column=start_column + 5, value=price)  # 单价
                sheet_.cell(row=row, column=start_column + 6,
                            value=f"={to_amnt_cell.coordinate}*{day_cnt_cell.coordinate}*{uni_prc_cell.coordinate}")  # 总价

        start_row_ += (depth_ + 1)

    # sum
    sheet_.cell(row=start_row_, column=start_column - 1, value="总计")
    sheet_.cell(row=start_row_, column=start_column + 6, value=f"=SUM(I5:I{start_row_ - 1})")
    # endregion manipulate the cells

    filename_ = os.path.join(os.path.join(os.environ['USERPROFILE'], 'Desktop'), '年度钢板结算单.xlsx')
    logger.info(f"the file is saved to {filename_}")

    workbook_.save(filename=filename_)
    return filename_


def main():
    price, structure = get_struct_from_input()
    # price, structure = 2.5, [[datetime.date(2023, 1, 1), 20, [[datetime.date(2023, 2, 1), 10], [datetime.date(2023, 2, 5), 5], [datetime.date(2023, 3, 3), 5]]], [datetime.date(2023, 2, 1), 15, [[datetime.date(2023, 5, 1), 10], [datetime.date(2023, 5, 10), 5]]], [datetime.date(2023, 3, 3), 5, [[datetime.date(2024, 5, 5), 5]]]]
    filename = form_xlsx_file(price, structure)
    print(f"已将输出信息保存到{filename}中！\n")
    os.system("pause")


if __name__ == '__main__':
    # From https://chat.openai.com/c/34221f1f-8901-42dc-9ff7-0929658ef3fb
    # pyinstaller --onefile -n SillyBee.py
    main()
