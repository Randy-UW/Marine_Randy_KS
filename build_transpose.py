import datetime

import openpyxl
import xlrd


# -*- coding = utf-8 -*-
# @Time : 2023/5/18 下午 11:28
# @Author : Randy
# @File : build_transpose.py
# @Software : PyCharm

def build_transpose_file(anno_file_xls: str, anno_sheet: str, row_number: int):
    wb = xlrd.open_workbook(anno_file_xls)
    ws = wb.sheet_by_name(anno_sheet)
    first_row_value: list = ws.row_values(row_number)
    transpose_xlsx = openpyxl.Workbook()
    transpose_xlsx.epoch = openpyxl.utils.datetime.CALENDAR_MAC_1904
    transpose_xlsx_ws = transpose_xlsx.active
    start_column = 'A'
    start_row = 1
    for value in first_row_value:
        transpose_xlsx_ws[start_column + str(start_row)] = value
        transpose_xlsx_ws[start_column + str(start_row)].number_format = 'YYYY/MM/DD'
        start_row += 1
    transpose_xlsx.save(anno_file_xls[: anno_file_xls.find('.')] + '_transpose.xlsx')
    print('transpose file completed')