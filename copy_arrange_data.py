from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd


def copy_arrange_data():

    # 读取 excel 文件
    excel_file = "I:/Buried-point-testing/jk埋点.xlsx"

    wb = load_workbook(excel_file)

    # 创建一个名为“格式化后埋点用例”的新工作表
    new_sheet = wb.create_sheet(title='格式化后埋点用例')

    # 获取所有工作表的名称
    sheet_names = wb.sheetnames

    # 选择第2个工作表
    sheet2 = wb[sheet_names[1]]

    # 提取第一列数据
    first_column = [cell.value for cell in sheet2['A']]

    # 获取数据的起始行和结束行
    start_row = sheet2.min_row
    end_row = sheet2.max_row

    # 获取数据的起始列和结束列
    start_column = sheet2.min_column
    end_column = sheet2.max_column


    # 获取第二列及其之后的数据
    all_columns = []
    for column in sheet2.iter_cols(min_row = start_row, values_only = True):
        all_columns.append(column)


   # 将数据放到单数列中
    column_indices = [i for i in range(1, end_column, 2)]
    while True:
        for i, value in enumerate(all_columns):
            if i % 2 != 0:
                cell = new_sheet.cell(row = 1,column = column_indices)
                for i, value in enumerate(first_column):
                    cell = new_sheet.cell(row=i+1, column=1)
                    cell.value = value
            
            else:
                if i == end_column*2-2:
                    break
        break
                 
        
          

        #else:
         #   cell = new_sheet.cell(row = 1,column = i)
          #  cell.value = value





    # 保存工作簿，包含整理后的数据
    output_file = "I:/Buried-point-testing/jk埋点.xlsx"
    wb.save(output_file)


copy_arrange_data()
