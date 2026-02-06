@echo off
chcp 65001 >nul
cd /d "%~dp0.."

echo ========================================
echo    法律文件脱敏处理工具
echo ========================================
echo.

if "%~1"=="" (
    echo 用法：将 DOCX 文件拖放到此批处理文件上
    echo.
    pause
    exit /b 1
)

echo 输入文件: %~1
echo.
echo 请选择规则来源:
    echo   [1] 前端导出的规则 (redaction-rules.json)
    echo   [2] 默认规则库 (data/default-rules.json)
echo.
set /p choice=请输入选择:

if "%choice%"=="" set choice=1

if "%choice%"=="1" (
    set "rules=redaction-rules.json"
    if not exist "redaction-rules.json" (
        echo.
        echo 错误: 未找到 redaction-rules.json
        echo 请先在前端工具中导出脱敏规则
        echo.
        pause
        exit /b 1
    )
) else (
    set "rules=data\default-rules.json"
)

echo.
echo 使用规则: %rules%
echo.
echo 正在处理...
echo.

node scripts\redact.js "%~1" "%rules%" "%~n1_脱敏%~x1"

if %errorlevel% neq 0 (
    echo.
    echo ========================================
    echo    错误: 脚本执行失败 (错误代码: %errorlevel%)
    echo ========================================
    echo.
    echo 请检查上方错误信息
    echo.
    pause
    exit /b %errorlevel%
)

echo.
echo ========================================
echo    处理完成！
echo ========================================
pause
