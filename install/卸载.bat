@echo off
chcp 65001 >nul
setlocal

REM ============================================================
REM  这个脚本已作废，故意不再执行任何操作。
REM
REM  作废原因：它会执行
REM      rd /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\"
REM  那个目录是【整台机器上所有 Office.js 加载项】的共享缓存，
REM  不是本加载项专属的。删掉它等于把别的公司、别的项目装的
REM  加载项缓存一并清空。
REM
REM  这个问题在 Excel 工具箱的安装程序里被评审判为阻塞级并已修掉：
REM  现在的卸载只删本加载项自己的目录和注册项，【完全不碰 Wef 缓存】。
REM
REM  正确做法：用 Excel 工具箱的一键安装包卸载。
REM      解压后双击「Excel工具箱.bat」，选择卸载。
REM ============================================================

echo.
echo   这个卸载脚本已作废，不会执行任何操作。
echo   ============================================
echo.
echo   请改用 Excel 工具箱的一键安装包：
echo     双击「Excel工具箱.bat」，选择卸载。
echo.
echo   原因：旧脚本会清空
echo     %%LOCALAPPDATA%%\Microsoft\Office\16.0\Wef\
echo   那是【所有 Office 网页加载项共用】的缓存目录，
echo   删掉会连带影响你装的其它加载项。
echo.
pause
exit /b 1
