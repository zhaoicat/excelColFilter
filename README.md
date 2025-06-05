# Excel 命令行导出工具

一个简单易用的命令行工具，用于从 Excel 文件中读取数据并按指定列导出新的 Excel 文件。

## 🚀 快速使用 - EXE 版本（推荐）

**无需安装 Python 环境，直接运行！**

1. **下载**: 获取 `dist/ExcelColumnFilter.exe` (约 40MB)
2. **运行**: 双击或命令行运行
3. **使用**: `ExcelColumnFilter.exe -i input.xls -o output.xlsx -c "1,2,5"`

详细说明请查看 [EXE 使用说明.md](EXE使用说明.md)

### EXE 版本快速示例

```bash
# 查看帮助
ExcelColumnFilter.exe --help

# 查看所有列
ExcelColumnFilter.exe -i data.xls --list-columns

# 导出指定列
ExcelColumnFilter.exe -i data.xls -o result.xlsx -c "编号,平台,商品标题"

# 下载图片
ExcelColumnFilter.exe -i data.xls -o result.xlsx -c "编号,商品图片,商品标题" --download-images
```

## 功能特点

- 支持 `.xls` 和 `.xlsx` 格式的 Excel 文件
- 支持按列序号或列名选择要导出的列
- 支持范围选择和批量导出
- 严格按照指定顺序导出列，保持用户设定的列顺序
- 自动处理长数字 ID，避免科学计数法显示
- 支持多线程下载商品原图并在 Excel 中展示（可选功能）
- 命令行操作，适合自动化脚本

## 安装依赖

```bash
pip install -r requirements.txt
```

## 基本用法

### 查看文件中所有可用的列

```bash
python cli_excel_processor.py -i your_file.xls --list-columns
```

### 按列序号导出

```bash
# 选择单个列
python cli_excel_processor.py -i input.xls -o output.xlsx -c "1,2,5,10"

# 选择范围
python cli_excel_processor.py -i input.xls -o output.xlsx -c "1-5,10,25-30"

# 混合选择
python cli_excel_processor.py -i input.xls -o output.xlsx -c "1,3-5,10,15-20"
```

### 按列名导出

```bash
python cli_excel_processor.py -i input.xls -o output.xlsx -c "编号,平台,站点,店铺名称,主订单号"
```

### 导出所有列

```bash
python cli_excel_processor.py -i input.xls -o output.xlsx -c "all"
```

### 下载商品图片并展示

```bash
# 导出包含商品图片的列，并多线程下载原图到本地和Excel中
python cli_excel_processor.py -i input.xls -o output.xlsx -c "编号,商品图片,商品标题" --download-images
```

图片功能特点：

- 多线程并行下载，速度快
- 保存完整原图到本地 `images/` 目录
- Excel 中显示时自动缩放到合适大小
- 支持图片去重和重复使用

## 参数说明

| 参数                | 说明                                    | 必需 |
| ------------------- | --------------------------------------- | ---- |
| `-i, --input`       | 输入 Excel 文件路径                     | ✅   |
| `-o, --output`      | 输出 Excel 文件路径 (默认: output.xlsx) | ❌   |
| `-c, --columns`     | 要导出的列                              | ❌\* |
| `--list-columns`    | 仅显示所有可用列，不进行导出            | ❌   |
| `--download-images` | 下载商品图片并在 Excel 中展示           | ❌   |

\*注：`-c` 和 `--list-columns` 必须至少使用其中一个

## 使用示例

```bash
# 查看帮助
python cli_excel_processor.py --help

# 查看文件列信息
python cli_excel_processor.py -i 82.xls --list-columns

# 导出基本订单信息
python cli_excel_processor.py -i 82.xls -o orders.xlsx -c "编号,平台,站点,店铺名称,主订单号,订单状态"

# 按序号导出财务数据
python cli_excel_processor.py -i 82.xls -o financial.xlsx -c "1,2,11,13,30,31,36,41,42"

# 导出商品信息并下载图片
python cli_excel_processor.py -i 82.xls -o products.xlsx -c "编号,商品图片,商品标题,商品数量" --download-images
```

## 批处理示例

### Linux/Mac 脚本

```bash
#!/bin/bash
# 批量导出不同类型的数据

# 导出基本订单信息
python cli_excel_processor.py -i orders.xls -o basic_orders.xlsx -c "编号,平台,站点,店铺名称,主订单号"

# 导出财务数据
python cli_excel_processor.py -i orders.xls -o financial.xlsx -c "编号,主订单号,买家付款金额（RMB）,最终毛利（RMB）"

# 导出客户信息
python cli_excel_processor.py -i orders.xls -o customers.xlsx -c "主订单号,买家姓名,买家地址,手机号"

echo "批处理完成！"
```

### Windows 批处理

```batch
@echo off
rem 批量导出不同类型的数据

python cli_excel_processor.py -i orders.xls -o basic_orders.xlsx -c "编号,平台,站点,店铺名称,主订单号"
python cli_excel_processor.py -i orders.xls -o financial.xlsx -c "编号,主订单号,买家付款金额（RMB）,最终毛利（RMB）"
python cli_excel_processor.py -i orders.xls -o customers.xlsx -c "主订单号,买家姓名,买家地址,手机号"

echo 批处理完成！
pause
```

## 注意事项

1. 列名必须与 Excel 文件中的完全匹配（包括空格和特殊字符）
2. 输出文件如果已存在会被覆盖
3. 输出文件统一为 `.xlsx` 格式
4. 工具自动处理订单号等长数字，确保不显示为科学计数法
5. 图片下载功能需要网络连接，下载的原图会保存在本地 images 目录中
6. 包含中文的路径在某些系统中可能需要使用引号包围

## 文件说明

### 核心文件

- `cli_excel_processor.py`: 命令行 Excel 处理工具
- `requirements.txt`: 项目依赖包列表
- `命令行使用说明.md`: 详细的使用说明和高级功能
- `README.md`: 本文件，快速开始指南

### EXE 打包相关

- `dist/ExcelColumnFilter.exe`: 打包后的可执行文件 (约 40MB)
- `excel_processor.spec`: PyInstaller 配置文件
- `build_simple.bat`: 简单打包脚本
- `build_exe.bat`: 高级打包脚本
- `test_exe.bat`: EXE 功能测试脚本
- `EXE使用说明.md`: EXE 版本详细使用说明

### 如何重新打包

```bash
# 方法1: 使用简单脚本（推荐）
.\build_simple.bat

# 方法2: 使用spec文件
.\build_exe.bat

# 方法3: 手动打包
pip install pyinstaller
pyinstaller --onefile --console --name "ExcelColumnFilter" cli_excel_processor.py
```
