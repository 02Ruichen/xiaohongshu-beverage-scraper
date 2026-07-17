@echo off
chcp 65001 >nul
cd /d "D:\claude\AI race forest"

echo.
echo ══════════════════════════════════════════════
echo  🧋 饮料赛道小红书抓取 v2
echo ══════════════════════════════════════════════
echo.
echo  前置检查...
echo.

REM 检查 OpenCLI
where opencli >nul 2>&1
if %errorlevel% neq 0 (
    echo  ❌ OpenCLI 未找到！请先安装 opencli
    pause
    exit /b 1
)

REM 检查登录态
echo  验证小红书登录...
for /f "tokens=*" %%i in ('opencli xiaohongshu whoami 2^>^&1 ^| findstr "logged_in"') do set LOGIN_STATUS=%%i
echo  %LOGIN_STATUS%

echo.
echo  ⚠️ 确认事项：
echo    1. Chrome 已登录小红书（建议用小号）
echo    2. opencli doctor 全绿
echo    3. 这是今天第一次跑
echo.
set /p CONFIRM="  确认无误按 Enter 开始，Ctrl+C 取消..."

echo.
echo  🚀 开始抓取...
echo.

REM 激活 venv 并运行
call "%USERPROFILE%\.agent-reach-venv\Scripts\activate.bat"
python beverage_scanner_v2.py

echo.
echo ══════════════════════════════════════════════
echo  完成！Excel 在 output 目录下~
echo ══════════════════════════════════════════════
echo.
pause
