import os
import json
import re
import openpyxl
from pathlib import Path

# 配置路径

def is_valid_filename(filename):
    """验证Excel文件名是否符合规范"""
    pattern = r'^[a-zA-Z][a-zA-Z0-9_]*$'
    return re.match(pattern, filename) is not None

EXCEL_DIR = "Excel"
JSON_DIR = "JSON"
CSHARP_DIR = "CSharp"

def get_csharp_type(excel_type):
    """将Excel类型转换为C#类型"""
    excel_type = str(excel_type).lower().strip()
    if excel_type == "float":
        return "float"
    elif excel_type == "int":
        return "int"
    return "string"  # 默认类型

def generate_csharp_class(class_name, headers, types):
    """生成C#类"""
    properties = []
    for header, excel_type in zip(headers, types):
        if header:  # 跳过空列
            prop_type = get_csharp_type(excel_type)
            prop_name = header.replace(" ", "_")
            properties.append(f"    public {prop_type} {prop_name} {{ get; set; }}")
    
    return f"""using System;

public class {class_name}
{{
{'\n'.join(properties)}
}}"""

def convert_excel_to_json():
    # 确保输出目录存在
    os.makedirs(JSON_DIR, exist_ok=True)
    os.makedirs(CSHARP_DIR, exist_ok=True)

    for excel_file in Path(EXCEL_DIR).glob("*.xlsx"):
        if excel_file.name.startswith("~$"):  # 跳过临时文件
            continue
            
        # 验证文件名有效性
        stem = excel_file.stem
        if not is_valid_filename(stem):
            print(f"Skipping invalid filename: {excel_file.name} (must start with letter, only letters/numbers/underscore)")
            continue
            
        print(f"Processing {excel_file.name}...")
        wb = openpyxl.load_workbook(excel_file)
        sheet = wb.active
        
        # 第一行备注(跳过)，第二行变量名，第三行类型
        headers = [cell.value for cell in sheet[2]]  # 第二行
        types = [cell.value for cell in sheet[3]]    # 第三行
        
        # 提取数据(从第四行开始)
        data = []
        for row in sheet.iter_rows(min_row=4, values_only=True):
            if not any(row):  # 跳过空行
                continue
            data.append(dict(zip(headers, row)))
        
        # 写入JSON
        json_path = Path(JSON_DIR) / f"{excel_file.stem}.json"
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        # 生成C#类
        cs_path = Path(CSHARP_DIR) / f"{excel_file.stem}.cs"
        with open(cs_path, 'w', encoding='utf-8') as f:
            f.write(generate_csharp_class(excel_file.stem, headers, types))
            
        print(f"Generated {json_path} and {cs_path}")

if __name__ == "__main__":
    convert_excel_to_json()
    print("Conversion completed!")
