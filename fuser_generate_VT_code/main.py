import sys
import os
import openpyxl

def format_number(num, length):
    return f"{num:>{length}}"

def process_excel(file_path, col_a_length=5, col_b_length=5):
    # 读取 Excel 文件
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    # 初始化结果列表
    result = []

    # 读取每一行的数据
    for row in sheet.iter_rows(min_row=1, values_only=True):
        if row[0] is not None and row[1] is not None:
            a = int(row[0])
            b = int(row[1])
            formatted_a = format_number(a, col_a_length)
            formatted_b = format_number(b, col_b_length)
            result.append(f"\t{{{formatted_a}, {formatted_b}}}")

    # 将结果列表转换为字符串
    formatted_result = ',\n'.join(result)

    # 输出最终格式
    output = f"{{\n{formatted_result}\n}}"

    # 获取文件名和路径
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    c_file_path = os.path.join(os.path.dirname(file_path), f"{base_name}.c")

    # 写入 .c 文件
    with open(c_file_path, 'w') as f:
        f.write(output)

if __name__ == "__main__":
    if len(sys.argv) > 1:
        file_paths = sys.argv[1:]
        col_a_length = 5
        col_b_length = 5
        
        for file_path in file_paths:
            if file_path.endswith('.xlsx'):
                process_excel(file_path, col_a_length, col_b_length)
            else:
                print(f"无效的文件类型：{file_path}")
    else:
        print("请将 .xlsx 文件拖放到此程序上。")


