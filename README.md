# Excel转JSON工具 (Python版)

## 功能说明
- 将Excel(.xlsx)转换为JSON格式
- 自动生成对应的C#类文件
- 支持基础数据类型自动转换
- 轻量级Python实现，无需编译

## 使用说明
1. 安装Python 3.x
2. 安装依赖：
   ```bash
   pip install -r requirements.txt
   ```
3. 将Excel文件放入`Excel`目录
4. 运行转换：
   ```bash
   python excel2json.py
   ```
   或双击`convert.bat`

## 输出文件
- JSON文件：`JSON/{表名}.json`
- C#类文件：`CSharp/{表名}.cs`

## 文件结构
```
/DataTable
  ├── convert.bat       # 一键转换脚本
  ├── excel2json.py     # 核心转换程序
  ├── requirements.txt  # Python依赖
  ├── README.md         # 使用说明
  ├── Excel/            # 输入Excel目录
  ├── JSON/             # 输出JSON目录
  └── CSharp/           # C#类输出目录
```

## 注意事项
- 只处理Excel中的第一个工作表
- 第一行作为JSON键名和C#属性名
- 自动跳过空行和临时文件(~$开头的文件)
- 所有字段类型默认为string(可在代码中修改)
