@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PORT=8610

echo.
echo 正在启动「水印工坊」本地服务...
echo 浏览器将自动打开 http://localhost:%PORT%/
echo 使用完毕后，关闭本窗口即可停止服务。
echo.

start "" http://localhost:%PORT%/

where py >nul 2>nul
if %errorlevel%==0 (
    py -3 -m http.server %PORT%
) else (
    python -m http.server %PORT%
)
