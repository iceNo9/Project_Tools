import openpyxl

#默认字典
paper_dict = {
        '彩色': {'慢速': (0,0,0,0,0), '常速': (0,0,0,0,0)},
        '黑白': {'慢速': (0,0,0,0,0), '常速': (0,0,0,0,0)}
    }
    

def read_excel(file_path):
    # 加载Excel文件
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    
    # 初始化变量
    data_structure = []
    current_paper_type = None
    paper_types = {}
    
    for row in sheet.iter_rows(min_row=1, values_only=True):
        if row[0] is not None:  # 新纸型开始
            if current_paper_type is not None:
                data_structure.append(paper_types)
                paper_types = {}
            current_paper_type = row[0]
            paper_types[current_paper_type]={
                '彩色': {'慢速': (0,0,0,0,0), '常速': (0,0,0,0,0)},
                '黑白': {'慢速': (0,0,0,0,0), '常速': (0,0,0,0,0)}
                }

            # 纸型数据      
            if row[1] is not None:  # 存在颜色
                color = row[1]
                if row[2] is not None:  # 检查是否有速度值，保存数据
                    if row[2] == 124: #常速
                        speed = '常速'
                    else: 
                        speed = '慢速'
                
                    paper_types[current_paper_type][color][speed] = row[3:8]
        else:
            # 同类纸型数据
            if row[1] is not None:  # 存在颜色
                color = row[1]
                if row[2] is not None:  # 检查是否有速度值，保存数据
                    if row[2] == 124: #常速
                        speed = '常速'
                    else: 
                        speed = '慢速'
                
                    paper_types[current_paper_type][color][speed] = row[3:8]
            
    if current_paper_type is not None and len(paper_types) > 0:
        data_structure.append(paper_types)

    # print(data_structure)
    
    return data_structure

def generate_c_code2(data_structure):
    c_code = "const unsigned int Data[] = {// start\n"
    
    for paper_type, colors in enumerate(data_structure):
        for paper_type_name, color_data in colors.items():
            c_code += "    {// " + paper_type_name.split('-')[0].strip() + "\n"
            
            for color_name, speed_data in color_data.items():
                c_code += "        {// " + color_name + "\n"
                
                for speed, steps in speed_data.items():
                    c_code += "            {" + ', '.join(map(str, steps)) + "}, // " + speed + "\n"
                    
                c_code += "        },\n"  # 结束颜色层级
        
            c_code += "    },\n"  # 结束纸型层级
    
    c_code += "};\n"
    return c_code

def generate_c_code(data_structure):
    c_code = "const unsigned int Data[] = {// start\n"

    # 遍历纸张类型
    for paper_type_index, paper_type in enumerate(data_structure):
        for paper_type_name, color_data in paper_type.items():
            c_code += "    {// " + paper_type_name + "\n"

            # 获取颜色名称及其速度数据
            color_items = list(color_data.items())
            for color_index, (color_name, speed_data) in enumerate(color_items):
                c_code += "        {// " + color_name + "\n"

                # 获取速度数据列表
                speed_items = list(speed_data.items())
                for i, (speed, steps) in enumerate(speed_items):
                    c_code += "            {" + ', '.join(map(str, steps)) + "}"  # 不加逗号
                    
                    # 如果不是最后一个元素，则添加逗号和换行符
                    if i < len(speed_items) - 1:
                        c_code += ",// " + speed + "\n"
                    else:  # 是最后一个元素，则只换行
                        c_code += "// " + speed + "\n"
                        
                c_code += "        }"  # 结束颜色层级
                
                # 如果不是最后一个颜色，则添加逗号和换行符
                if color_index < len(color_items) - 1:
                    c_code += ",\n"
                else:  # 是最后一个颜色，则只换行
                    c_code += "\n"
        
            # 结束纸型层级
            c_code += "    }"  # 先不加逗号
            
            # 如果不是最后一个纸型，则添加逗号和换行符
            if paper_type_index < len(data_structure) - 1:
                c_code += ",\n"
            else:  # 是最后一个纸型，则只换行
                c_code += "\n"
    
    c_code += "};\n"
    return c_code



def write_to_file(c_code, file_path):
    output_file = file_path.replace('.xlsx', '.c')
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(c_code)

if __name__ == '__main__':
    import sys
    if len(sys.argv) != 2:
        print("Usage: python script.py <xlsx_file>")
        sys.exit(1)
        
    xlsx_file = sys.argv[1]
    data_structure = read_excel(xlsx_file)
    c_code = generate_c_code(data_structure)
    write_to_file(c_code, xlsx_file)
    print(f"C code has been written to {xlsx_file.replace('.xlsx', '.c')}")
