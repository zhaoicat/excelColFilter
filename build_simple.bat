@echo off
chcp 65001 >nul
echo ========================================
echo Excel列过滤工具 - 简单打包
echo ========================================
echo.

echo 正在清理之前的构建文件...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
echo.

echo 开始打包（一体化exe文件）...
echo 这可能需要几分钟时间，请耐心等待...
echo.

pyinstaller --onefile --console --name "ExcelColumnFilter" ^
    --hidden-import pandas ^
    --hidden-import openpyxl ^
    --hidden-import xlrd ^
    --hidden-import lxml ^
    --hidden-import requests ^
    --hidden-import PIL ^
    --hidden-import PIL.Image ^
    --hidden-import openpyxl.drawing.image ^
    --hidden-import concurrent.futures ^
    --hidden-import urllib.parse ^
    --hidden-import hashlib ^
    --exclude-module tkinter ^
    --exclude-module matplotlib ^
    --exclude-module scipy ^
    --exclude-module IPython ^
    --exclude-module jupyter ^
    --exclude-module PyQt5 ^
    --exclude-module PyQt6 ^
    cli_excel_processor.py

echo.
if exist "dist\ExcelColumnFilter.exe" (
    echo ========================================
    echo ✅ 打包成功！
    echo ========================================
    echo 可执行文件位置: dist\ExcelColumnFilter.exe
    echo 文件大小: 
    for %%A in ("dist\ExcelColumnFilter.exe") do echo     %%~zA 字节 ^(约 %%~zA/1024/1024 MB^)
    echo.
    echo 使用方法:
    echo   ExcelColumnFilter.exe -i input.xls -o output.xlsx -c "1,2,5"
    echo   ExcelColumnFilter.exe --help
    echo.
    echo 正在测试exe文件...
    "dist\ExcelColumnFilter.exe" --help
) else (
    echo ========================================
    echo ❌ 打包失败！
    echo ========================================
    echo 请检查上面的错误信息
)

echo.
echo 按任意键退出...
pause >nul 