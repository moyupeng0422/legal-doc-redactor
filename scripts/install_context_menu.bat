@echo off
setlocal enabledelayedexpansion

echo ========================================
echo   法律文件脱敏工具 - 右键菜单安装
echo ========================================
echo.

:: 获取脚本所在目录和项目根目录
set "SCRIPT_DIR=%~dp0"
set "PROJECT_DIR=%SCRIPT_DIR%.."

:: 检查核心文件是否存在
if not exist "%SCRIPT_DIR%context_menu_server.py" (
    echo [错误] 未找到 context_menu_server.py
    echo 请确保此脚本位于 scripts\ 目录中
    pause
    exit /b 1
)
if not exist "%SCRIPT_DIR%quick_redact.py" (
    echo [错误] 未找到 quick_redact.py
    echo 请确保此脚本位于 scripts\ 目录中
    pause
    exit /b 1
)
if not exist "%PROJECT_DIR%\assets\index.html" (
    echo [错误] 未找到 assets\index.html
    echo 请确保目录结构完整
    pause
    exit /b 1
)
echo [OK] 文件检查通过

:: 检测 Python
echo.
echo [1/4] 检测 Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到 Python
    echo 请先安装 Python 3.x: https://www.python.org/downloads/
    pause
    exit /b 1
)
for /f "tokens=*" %%v in ('python --version 2^>^&1') do echo [OK] %%v

:: 检测 pythonw.exe（无窗口运行），需要完整路径
echo.
echo [2/4] 检测 pythonw.exe...
set "PYTHONW_CMD="
for /f "delims=" %%i in ('where pythonw 2^>nul') do set "PYTHONW_CMD=%%i"
if "%PYTHONW_CMD%"=="" (
    echo [警告] 未找到 pythonw.exe，尝试使用 python
    for /f "delims=" %%i in ('where python 2^>nul') do set "PYTHONW_CMD=%%i"
    if "%PYTHONW_CMD%"=="" (
        echo [错误] 未找到 Python，请先安装 Python 3.x
        pause
        exit /b 1
    )
    echo [OK] 使用 python（运行时会显示命令行窗口）
) else (
    echo [OK] %PYTHONW_CMD%
)

:: 检测/安装 python-docx
echo.
echo [3/4] 检测 python-docx...
python -c "import docx" >nul 2>&1
if errorlevel 1 (
    echo python-docx 未安装，正在自动安装...
    pip install python-docx >nul 2>&1
    if errorlevel 1 (
        echo [错误] python-docx 安装失败，请手动运行: pip install python-docx
        pause
        exit /b 1
    )
    echo [OK] python-docx 安装成功
) else (
    echo [OK] python-docx 已安装
)

:: 写入注册表
echo.
echo [4/4] 写入注册表...
set "REG_BASE=HKCU\SOFTWARE\Classes\SystemFileAssociations\.docx\shell"
set "ICON_PATH=%PROJECT_DIR%\assets\redact.ico"
set "SERVER_CMD=\"%PYTHONW_CMD%\" \"%SCRIPT_DIR%context_menu_server.py\" \"%%1\""
set "REDACT_CMD=\"%PYTHONW_CMD%\" \"%SCRIPT_DIR%quick_redact.py\" \"%%1\""

:: 安装"用脱敏工具打开"
echo   - 用脱敏工具打开...
reg add "%REG_BASE%\RedactOpen" /ve /t REG_SZ /d "用脱敏工具打开" /f >nul 2>&1
if errorlevel 1 (
    echo [错误] 注册表写入失败
    pause
    exit /b 1
)
reg add "%REG_BASE%\RedactOpen" /v Icon /t REG_SZ /d "%ICON_PATH%" /f >nul 2>&1
reg add "%REG_BASE%\RedactOpen\command" /ve /t REG_SZ /d "%SERVER_CMD%" /f >nul 2>&1

:: 安装"一键脱敏"
echo   - 一键脱敏...
reg add "%REG_BASE%\QuickRedact" /ve /t REG_SZ /d "一键脱敏" /f >nul 2>&1
reg add "%REG_BASE%\QuickRedact" /v Icon /t REG_SZ /d "%ICON_PATH%" /f >nul 2>&1
reg add "%REG_BASE%\QuickRedact\command" /ve /t REG_SZ /d "%REDACT_CMD%" /f >nul 2>&1

echo.
echo ========================================
echo   安装完成！
echo ========================================
echo.
echo 已添加以下右键菜单项：
echo   [1] 用脱敏工具打开 - 在浏览器中打开脱敏工具
echo   [2] 一键脱敏       - 直接在后台执行脱敏
echo.
echo 安装位置: "%PROJECT_DIR%"
echo.
echo 卸载方法: "运行 scripts\uninstall_context_menu.bat"
echo.
pause
