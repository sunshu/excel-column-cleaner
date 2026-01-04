@echo off
echo 启动后端服务器...
cd backend
pip install -r requirements.txt
cd app
python main.py
pause