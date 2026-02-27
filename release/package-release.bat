@echo off
setlocal enabledelayedexpansion

echo ========================================
echo   法律文件脱敏工具 - 自动打包脚本
echo ========================================
echo.

set VERSION=1.3.0
set PROJECT_ROOT=..
set RELEASE_DIR=output\legal-doc-redactor-v%VERSION%
set ZIP_NAME=legal-doc-redactor-v%VERSION%.zip

REM 创建发布目录
if exist "%RELEASE_DIR%" rd /s /q "%RELEASE_DIR%"
mkdir "%RELEASE_DIR%"
mkdir "%RELEASE_DIR%\examples"

echo [1/5] 复制核心文件...
copy "%PROJECT_ROOT%\assets\index.html" "%RELEASE_DIR%\" >nul
copy "%PROJECT_ROOT%\assets\mammoth.browser.min.js" "%RELEASE_DIR%\" >nul
copy "%PROJECT_ROOT%\assets\docx-v7.js" "%RELEASE_DIR%\" >nul
copy "%PROJECT_ROOT%\assets\jszip.min.js" "%RELEASE_DIR%\" >nul
copy "%PROJECT_ROOT%\data\default-rules.json" "%RELEASE_DIR%\" >nul

echo [2/5] 复制使用说明...
copy "README.md" "%RELEASE_DIR%\README.md" >nul
copy "USER-GUIDE.html" "%RELEASE_DIR%\User-Guide.html" >nul

echo [3/5] 复制示例文件...
if exist "examples\*" (
    copy "examples\*.*" "%RELEASE_DIR%\examples\" >nul 2>&1
) else (
    echo. > "%RELEASE_DIR%\examples\readme.txt"
    echo Please place test documents here > "%RELEASE_DIR%\examples\readme.txt"
)

echo [4/5] 压缩发布包...
if exist "output\%ZIP_NAME%" del "output\%ZIP_NAME%"
powershell -Command "Compress-Archive -Path '%RELEASE_DIR%' -DestinationPath 'output\%ZIP_NAME%' -Force"

echo.
echo ========================================
echo   打包完成！
echo ========================================
echo.
echo 输出目录: %RELEASE_DIR%\
echo 压缩文件: output\%ZIP_NAME%
echo.
echo 您可以将 output\%ZIP_NAME% 发送给朋友使用
echo.

pause
