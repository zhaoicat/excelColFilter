@echo off
chcp 65001 >nul
echo ========================================
echo Excel列过滤工具 - EXE功能测试
echo ========================================
echo.

set EXE_PATH=dist\ExcelColumnFilter.exe
set TEST_FILE=82.xls

echo 检查exe文件是否存在...
if not exist "%EXE_PATH%" (
    echo ❌ 错误: %EXE_PATH% 不存在
    echo 请先运行 build_simple.bat 进行打包
    pause
    exit /b 1
)

echo 检查测试文件是否存在...
if not exist "%TEST_FILE%" (
    echo ❌ 错误: %TEST_FILE% 不存在
    echo 请确保测试文件存在
    pause
    exit /b 1
)

echo.
echo ✅ 文件检查通过
echo.

echo ========================================
echo 测试1: 显示帮助信息
echo ========================================
"%EXE_PATH%" --help
echo.

echo ========================================
echo 测试2: 显示列信息
echo ========================================
"%EXE_PATH%" -i "%TEST_FILE%" --list-columns
echo.

echo ========================================
echo 测试3: 导出前5列
echo ========================================
"%EXE_PATH%" -i "%TEST_FILE%" -o "test_output_1.xlsx" -c "1-5"
if exist "test_output_1.xlsx" (
    echo ✅ 测试3通过: test_output_1.xlsx 已生成
) else (
    echo ❌ 测试3失败: 文件未生成
)
echo.

echo ========================================
echo 测试4: 按列名导出
echo ========================================
"%EXE_PATH%" -i "%TEST_FILE%" -o "test_output_2.xlsx" -c "编号,平台,站点,店铺名称,商品标题"
if exist "test_output_2.xlsx" (
    echo ✅ 测试4通过: test_output_2.xlsx 已生成
) else (
    echo ❌ 测试4失败: 文件未生成
)
echo.

echo ========================================
echo 测试5: 混合选择导出
echo ========================================
"%EXE_PATH%" -i "%TEST_FILE%" -o "test_output_3.xlsx" -c "1,2,10-15,20"
if exist "test_output_3.xlsx" (
    echo ✅ 测试5通过: test_output_3.xlsx 已生成
) else (
    echo ❌ 测试5失败: 文件未生成
)
echo.

echo ========================================
echo 测试完成！
echo ========================================
echo 生成的测试文件:
if exist "test_output_1.xlsx" echo   - test_output_1.xlsx
if exist "test_output_2.xlsx" echo   - test_output_2.xlsx  
if exist "test_output_3.xlsx" echo   - test_output_3.xlsx
echo.

echo 是否要清理测试文件？(y/n)
set /p cleanup=
if /i "%cleanup%"=="y" (
    if exist "test_output_1.xlsx" del "test_output_1.xlsx"
    if exist "test_output_2.xlsx" del "test_output_2.xlsx"
    if exist "test_output_3.xlsx" del "test_output_3.xlsx"
    echo 测试文件已清理
)

echo.
echo 按任意键退出...
pause >nul 