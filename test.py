import xlrd
import os

DATA_DIR = os.path.join(os.path.dirname(__file__))
df = xlrd.open_workbook(os.path.join(DATA_DIR, "data.xls"))


def get_first_and_second_column_values_by_keyword(keyword):
    # 获取第一张表
    sheet = df.sheet_by_index(0)
    # 初始化结果列表
    result = []
    # 遍历所有行
    for row_idx in range(sheet.nrows):
        # 获取第一行的值
        first_col_value = sheet.cell_value(row_idx, 0)
        # 检查第一列是否包含关键字
        if keyword in str(first_col_value):
            # 如果包含，获取第二列的值
            second_col_value = sheet.cell_value(row_idx, 1)
            # 将第一列和第二列的值作为元组添加到结果列表
            result.append((first_col_value, second_col_value))
    return result


# 示例调用
for item in get_first_and_second_column_values_by_keyword("艾尔登"):
    print(f"文件名: {item[0]} -> 链接: {item[1]}")
