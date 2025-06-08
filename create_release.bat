@echo off
chcp 65001 >nul
echo ========================================
echo Excel 列过滤工具 - 创建发布包
echo ========================================
echo.

set RELEASE_DIR=release
set VERSION=v1.0.0
set RELEASE_NAME=ExcelColumnFilter_%VERSION%

echo 正在创建发布包...
echo.

REM 检查exe文件是否存在
if not exist "dist\ExcelColumnFilter.exe" (
    echo ❌ 错误: ExcelColumnFilter.exe 不存在
    echo 请先运行 build_simple.bat 进行打包
    pause
    exit /b 1
)

REM 清理并重新创建发布目录
if exist "%RELEASE_DIR%" rmdir /s /q "%RELEASE_DIR%"
mkdir "%RELEASE_DIR%"
mkdir "%RELEASE_DIR%\示例文件"

echo 复制主程序文件...
copy "dist\ExcelColumnFilter.exe" "%RELEASE_DIR%\"

echo 复制文档文件...
copy "EXE使用说明.md" "%RELEASE_DIR%\"

echo 创建发布版README...
echo # Excel 列过滤工具 v1.0 > "%RELEASE_DIR%\README.md"
echo. >> "%RELEASE_DIR%\README.md"
echo 一个简单易用的 Excel 列导出工具，无需安装 Python 环境即可使用。 >> "%RELEASE_DIR%\README.md"
echo. >> "%RELEASE_DIR%\README.md"
echo ## 🚀 快速开始 >> "%RELEASE_DIR%\README.md"
echo. >> "%RELEASE_DIR%\README.md"
echo 1. 查看帮助：ExcelColumnFilter.exe --help >> "%RELEASE_DIR%\README.md"
echo 2. 查看列信息：ExcelColumnFilter.exe -i 文件.xls --list-columns >> "%RELEASE_DIR%\README.md"
echo 3. 导出列：ExcelColumnFilter.exe -i 文件.xls -o 输出.xlsx -c "1,2,5" >> "%RELEASE_DIR%\README.md"
echo. >> "%RELEASE_DIR%\README.md"
echo 详细说明请查看 EXE使用说明.md 和 快速开始.txt >> "%RELEASE_DIR%\README.md"

echo 创建快速开始指南...
echo 🚀 Excel 列过滤工具 - 快速开始指南 > "%RELEASE_DIR%\快速开始.txt"
echo ===================================== >> "%RELEASE_DIR%\快速开始.txt"
echo. >> "%RELEASE_DIR%\快速开始.txt"
echo 1. 双击 "示例文件/查看列信息.bat" 查看 Excel 文件结构 >> "%RELEASE_DIR%\快速开始.txt"
echo 2. 使用 "示例文件" 中的脚本快速导出常用数据 >> "%RELEASE_DIR%\快速开始.txt"
echo 3. 或使用命令行：ExcelColumnFilter.exe -i 文件.xls -o 输出.xlsx -c "1,2,5" >> "%RELEASE_DIR%\快速开始.txt"
echo. >> "%RELEASE_DIR%\快速开始.txt"
echo 详细说明请查看 README.md 和 EXE使用说明.md >> "%RELEASE_DIR%\快速开始.txt"

echo 创建版本信息...
echo Excel 列过滤工具 - 版本信息 > "%RELEASE_DIR%\版本信息.txt"
echo ============================= >> "%RELEASE_DIR%\版本信息.txt"
echo. >> "%RELEASE_DIR%\版本信息.txt"
echo 版本：%VERSION% >> "%RELEASE_DIR%\版本信息.txt"
echo 发布日期：%date% >> "%RELEASE_DIR%\版本信息.txt"
echo 文件大小：约 40MB >> "%RELEASE_DIR%\版本信息.txt"
echo 兼容系统：Windows 10/11 >> "%RELEASE_DIR%\版本信息.txt"
echo. >> "%RELEASE_DIR%\版本信息.txt"
echo 主要功能：Excel 列选择导出、图片下载、批处理支持 >> "%RELEASE_DIR%\版本信息.txt"

echo 创建示例脚本...

REM 查看列信息脚本
echo @echo off > "%RELEASE_DIR%\示例文件\查看列信息.bat"
echo chcp 65001 ^>nul >> "%RELEASE_DIR%\示例文件\查看列信息.bat"
echo echo 请拖拽 Excel 文件到此窗口： >> "%RELEASE_DIR%\示例文件\查看列信息.bat"
echo set /p input_file= >> "%RELEASE_DIR%\示例文件\查看列信息.bat"
echo ..\ExcelColumnFilter.exe -i %%input_file%% --list-columns >> "%RELEASE_DIR%\示例文件\查看列信息.bat"
echo pause >> "%RELEASE_DIR%\示例文件\查看列信息.bat"

REM 导出基本信息脚本
echo @echo off > "%RELEASE_DIR%\示例文件\导出基本信息.bat"
echo chcp 65001 ^>nul >> "%RELEASE_DIR%\示例文件\导出基本信息.bat"
echo echo 请拖拽 Excel 文件到此窗口： >> "%RELEASE_DIR%\示例文件\导出基本信息.bat"
echo set /p input_file= >> "%RELEASE_DIR%\示例文件\导出基本信息.bat"
echo ..\ExcelColumnFilter.exe -i %%input_file%% -o "基本信息.xlsx" -c "编号,平台,站点,店铺名称,主订单号,订单状态" >> "%RELEASE_DIR%\示例文件\导出基本信息.bat"
echo pause >> "%RELEASE_DIR%\示例文件\导出基本信息.bat"

REM 使用说明
echo 示例文件使用说明 > "%RELEASE_DIR%\示例文件\使用说明.txt"
echo ================== >> "%RELEASE_DIR%\示例文件\使用说明.txt"
echo. >> "%RELEASE_DIR%\示例文件\使用说明.txt"
echo 1. 查看列信息.bat - 查看 Excel 文件的所有列 >> "%RELEASE_DIR%\示例文件\使用说明.txt"
echo 2. 导出基本信息.bat - 导出常用的基本信息列 >> "%RELEASE_DIR%\示例文件\使用说明.txt"
echo. >> "%RELEASE_DIR%\示例文件\使用说明.txt"
echo 使用方法：双击脚本，然后拖拽 Excel 文件到窗口 >> "%RELEASE_DIR%\示例文件\使用说明.txt"

echo.
echo ✅ 发布包创建完成！
echo.
echo 📁 发布目录：%RELEASE_DIR%\
echo 📦 包含文件：
echo   - ExcelColumnFilter.exe (主程序)
echo   - README.md (使用说明)
echo   - EXE使用说明.md (详细文档)
echo   - 快速开始.txt (快速指南)
echo   - 版本信息.txt (版本信息)
echo   - 示例文件\ (示例脚本)
echo.

echo 是否要创建压缩包？(y/n)
set /p create_zip=
if /i "%create_zip%"=="y" (
    if exist "%RELEASE_NAME%.zip" del "%RELEASE_NAME%.zip"
    powershell -command "Compress-Archive -Path '%RELEASE_DIR%\*' -DestinationPath '%RELEASE_NAME%.zip'"
    if exist "%RELEASE_NAME%.zip" (
        echo ✅ 压缩包已创建：%RELEASE_NAME%.zip
    ) else (
        echo ❌ 压缩包创建失败
    )
)

echo.
echo 🎉 发布包准备完成！可以交付给其他人使用了。
echo.
echo 按任意键退出...
pause >nul 