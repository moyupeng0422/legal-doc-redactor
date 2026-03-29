@echo off

echo ========================================
echo   法律文件脱敏处理 - 右键菜单卸载
echo ========================================
echo.

set "REG_BASE=HKCU\SOFTWARE\Classes\SystemFileAssociations\.docx\shell"

:: 删除"用脱敏工具打开"
echo [卸载] 用脱敏工具打开...
reg delete "%REG_BASE%\RedactOpen" /f >nul 2>&1

:: 删除"一键脱敏"
echo [卸载] 一键脱敏...
reg delete "%REG_BASE%\QuickRedact" /f >nul 2>&1

:: 删除"一键还原"
echo [卸载] 一键还原...
reg delete "%REG_BASE%\QuickRestore" /f >nul 2>&1

echo.
echo ========================================
echo   卸载完成！
echo ========================================
echo.
echo 已移除以下右键菜单：
echo   - 用脱敏工具打开
echo   - 一键脱敏
echo   - 一键还原
echo.
pause
