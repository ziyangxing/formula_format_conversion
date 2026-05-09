@echo off
chcp 65001 >nul
echo ========================================
echo   FormulaConverter - 打包构建脚本
echo ========================================
echo.

REM 检查 PyInstaller
pyinstaller --version >nul 2>&1
if errorlevel 1 (
    echo [安装] 正在安装 PyInstaller...
    pip install pyinstaller
    echo.
)

echo [清理] 清理旧的构建文件...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "FormulaConverter.spec" del /q "FormulaConverter.spec"

echo [构建] 正在打包 FormulaConverter.exe...
pyinstaller --noconfirm ^
    --onefile ^
    --windowed ^
    --name "FormulaConverter" ^
    --icon=NONE ^
    --add-data "convert.py;." ^
    --add-data "latex2omml.py;." ^
    --hidden-import "lxml.etree" ^
    --hidden-import "docx" ^
    --hidden-import "latex2omml" ^
    --hidden-import "convert" ^
    --collect-submodules "lxml" ^
    gui.py

if errorlevel 1 (
    echo.
    echo [失败] 构建出错，请检查错误信息
    pause
    exit /b 1
)

echo.
echo ========================================
echo   构建完成！
echo   输出文件: dist\FormulaConverter.exe
echo ========================================
echo.
pause
