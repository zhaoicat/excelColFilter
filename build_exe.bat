@echo off
chcp 65001 >nul
echo ========================================
echo Excel列过滤工具 - 打包为EXE文件
echo ========================================
echo.

echo 正在清理之前的构建文件...
if exist "build" rmdir /s /q "build"
if exist "dist\ExcelColumnFilter.exe" del /q "dist\ExcelColumnFilter.exe"
echo.

echo 开始使用PyInstaller打包...
echo 这可能需要几分钟时间，请耐心等待...
echo.

pyinstaller --clean excel_processor.spec

echo.
if exist "dist\ExcelColumnFilter.exe" (
    echo ========================================
    echo ✅ 打包成功！
    echo ========================================
    echo 可执行文件位置: dist\ExcelColumnFilter.exe
    echo 文件大小: 
    for %%A in ("dist\ExcelColumnFilter.exe") do echo     %%~zA 字节
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