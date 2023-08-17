from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd



def ATMlog_data():

    # 读取源 excel 文件
    excel_file = "I:/Buried-point-testing/ATMlog.xlsx"

    source_workbook = load_workbook(excel_file)

    source_sheet = source_workbook.active

    # 加载目标 excel 文件
    excel_file2 = 'I:/Buried-point-testing/jk埋点.xlsx'

    target_workbook = load_workbook(excel_file2)

    # 创建一个名为“埋点数据”的新工作表
 
    new_sheet = target_workbook.create_sheet(title="埋点数据")

    target_sheet = target_workbook.sheetnames

    # 选择埋点数据表
    summary_table_sheet = target_workbook[target_sheet[2]]

    # 从源表中取出数据并放置到目标表中
    for row in source_sheet.iter_rows():
        row_data = [cell.value for cell in row]
        summary_table_sheet.append(row_data)


    # 删除第三列
    summary_table_sheet.delete_cols(3)

    # 提取第五列数据
    first_column = [cell.value for col in summary_table_sheet.iter_cols(min_col=5, max_col=5) for cell in col]



    # 保存工作簿，包含整理后的数据
    output_file = 'I:/Buried-point-testing/jk埋点.xlsx'
    target_workbook.save(output_file)


#ATMlog_data()


def result_table():
    '''处理查询出来的埋点数据，整理成结果表'''
    # 加载目标 excel 文件
    excel_file2 = 'I:/Buried-point-testing/jk埋点.xlsx'

    target_workbook = load_workbook(excel_file2)

    # 创建一个名为“QQ埋点数据”的新工作表
    new_sheet = target_workbook.create_sheet(title="QQ埋点结果表")

    sheets = target_workbook.sheetnames

    # 选择源数据表和目标数据表
    source_sheet = target_workbook.worksheets[0]
    target_sheet = target_workbook.worksheets[3]

    # 指定要选中的列的索引
    source_column_indices = [1, 2, 4]

    target_column_indices = [1, 2, 3]

    # 将124列的内容放到QQ埋点数据表中
    for source_index, target_index in zip(source_column_indices, target_column_indices):
        source_column = [cell[0].value for cell in source_sheet.iter_cols(min_col=source_index, max_col=source_index)]
        for row, value in enumerate(source_column, start=1):
            target_sheet.cell(row=row, column=target_index, value=value)

result_table()
 

    




        



