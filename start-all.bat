@echo off
echo ================================
echo Excel 列删除工具 - 一键启动
echo ================================
echo.

echo 正在检查 Python 环境...
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到 Python，请先安装 Python 3.8+
    pause
    exit /b 1
)

echo 正在检查 Node.js 环境...
node --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到 Node.js，请先安装 Node.js 16+
    pause
    exit /b 1
)

echo.
echo 1. 启动后端服务...
start "Excel 后端服务" cmd /k "cd backend && pip install -r requirements.txt && cd app && python main.py"

echo 等待后端服务启动...
timeout /t 5 /nobreak >nul

echo.
echo 2. 启动前端服务...
start "Excel 前端服务" cmd /k "cd frontend && npm install && npm run dev"

echo.
echo ================================
echo 服务启动完成！
echo ================================
echo 后端服务: http://localhost:8001
echo 前端应用: http://localhost:3000
echo API 文档: http://localhost:8001/docs
echo.
echo 请等待前端服务完全启动后访问应用
echo 按任意键关闭此窗口...
pause >nul