from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd


def copy_arrange_data():

    # 读取 excel 文件
    excel_file = '/Users/xinwang/Desktop/auto的副本/jk埋点.xlsx'

    wb = load_workbook(excel_file)

    # 创建一个名为“格式化后埋点用例”的新工作表
    new_sheet = wb.create_sheet(title='格式化后埋点用例')

    # 获取所有工作表的名称
    sheet_names = wb.sheetnames

    # 选择第2个工作表
    sheet2 = wb[sheet_names[1]]

    # 读取数据并整理为指定格式
    arranged_data = []

    for row_index in range(2, sheet2.max_row + 1):
        base_value = sheet2.cell(row=row_index, column=1).value
        for col_index in range(2, sheet2.max_column + 1):
            current_value = sheet2.cell(row=row_index, column=col_index).value
            arranged_data.append([base_value, current_value])

    # 将整理后的数据写入新工作表
    for row in arranged_data:
        new_sheet.append(row)

    # 保存工作簿，包含整理后的数据
    output_file = '/Users/xinwang/Desktop/auto的副本/jk埋点.xlsx'
    wb.save(output_file)


copy_arrange_data()
