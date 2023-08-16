from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd


def copy_arrange_data():

    def ATMlog_data():

        # 读取 excel 文件
        excel_file = "I:/Buried-point-testing/ATMlog.xlsx"

        wb = load_workbook(excel_file)

        # 创建一个名为“格式化后埋点用例”的新工作表
        new_sheet = wb.create_sheet(title='测试结果')

        # 获取所有工作表的名称
        sheet_names = wb.sheetnames

        # 选择查询数据表中的第1个工作表
        sheet1 = wb[sheet_names[0]]

        # 删除第三列
        sheet1.delete_cols(3)

        # 提取第五列数据
        first_column = [cell.value for cell in sheet1['E']]

        # 保存工作簿，包含整理后的数据
        output_file = "I:/Buried-point-testing/ATMlog.xlsx"
        wb.save(output_file)


    def traspose_data():

        excel_file2 = 'I:/Buried-point-testing/jk埋点.xlsx'

        wb2 = load_workbook(excel_file2)

         # 选择查询数据表中的第1个工作表
        sheet_names = wb2.sheetnames

        sheet1 = wb2[sheet_names[0]]

        # 提取指定列的数据
        column_indices = [1, 2, 4]

         # 提取第1、2、4列数据
        extracted_data = []
        for row in sheet1.iter_rows():
            row_data = [row[i - 1].value for i in column_indices]
            extracted_data.append(row_data)


    ATMlog_data()
    traspose_data()

    




        


copy_arrange_data()
