@echo off
chcp 65001 >nul
title ARCHI-TRANSLATE STUDIO // 空间深化与定制家具工程资料翻译工作台
echo ===================================================================
echo   ARCHI-TRANSLATE ENTERPRISE WORKBENCH (Studio Edition v2.5)
echo   [01] CAD 施工与深化图纸模式 (视觉大模型读图 + PaddleOCR 检测)
echo   [02] 说明书 / 商务画册与技术合同 (BabelDoc 流式排版)
echo   [03] 企业标准术语库 (分类 / 使用次数 / 导入导出)
echo   [04] 项目记录 (历史持久化) 与实时资产看板
echo ===================================================================
echo.

cd /d "%~dp0"

if not exist "%~dp0venv\Scripts\activate.bat" (
    echo [错误] 未找到本地运行环境 venv，请确认 venv 文件夹与本程序在同一目录。
    pause
    exit /b 1
)

rem Check if service is already running (double-launch protection)
netstat -ano | findstr ":7860" | findstr "LISTENING" >nul 2>&1
if not errorlevel 1 (
    echo [提示] 检测到翻译服务已在运行，直接打开工作台页面...
    start "" http://127.0.0.1:7860
    pause
    exit /b 0
)

echo [1/2] 正在加载 Python 虚拟环境...
call "%~dp0venv\Scripts\activate.bat"

echo [2/2] 正在启动翻译服务（就绪后浏览器会自动打开，请勿关闭本窗口）...

rem Poll port 7860 in background; open browser only after service is ready
start "" powershell -NoProfile -WindowStyle Hidden -Command "for($i=0;$i -lt 120;$i++){try{$c=New-Object Net.Sockets.TcpClient('127.0.0.1',7860);$c.Close();Start-Process 'http://127.0.0.1:7860';exit}catch{Start-Sleep -Milliseconds 500}}"

python "%~dp0server.py"

echo.
echo [提示] 服务已退出。若上方出现报错信息，请把截图发给管理员。
pause
