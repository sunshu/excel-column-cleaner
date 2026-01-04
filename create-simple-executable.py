"""
创建简单的可执行文件
直接使用PyInstaller的基本功能
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path

def create_simple_main():
    """创建一个简化的main.py文件用于打包"""
    
    # 创建临时目录
    temp_dir = Path("temp_build")
    if temp_dir.exists():
        shutil.rmtree(temp_dir)
    temp_dir.mkdir()
    
    # 复制所有后端文件到临时目录
    backend_src = Path("backend/app")
    for file_path in backend_src.rglob("*.py"):
        relative_path = file_path.relative_to(backend_src)
        dest_path = temp_dir / relative_path
        dest_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(file_path, dest_path)
    
    # 修改main.py以包含所有必要的导入
    main_content = '''"""
Excel 列删除工具 - 可执行版本
"""

import sys
import os
from pathlib import Path

# 添加当前目录到Python路径
if getattr(sys, 'frozen', False):
    # 运行在PyInstaller打包的环境中
    application_path = sys._MEIPASS
else:
    # 运行在正常Python环境中
    application_path = os.path.dirname(os.path.abspath(__file__))

sys.path.insert(0, application_path)

# 现在导入应用模块
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import JSONResponse, FileResponse
import logging
import socket

# 导入控制器
from controllers.excel_controller import router as excel_router

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# 创建 FastAPI 应用
app = FastAPI(
    title="Excel 列删除工具 API",
    description="提供 Excel 文件列删除功能的后端服务",
    version="1.0.0"
)

# 获取静态文件路径
if getattr(sys, 'frozen', False):
    # 打包后的可执行文件
    static_dir = os.path.join(sys._MEIPASS, 'static')
else:
    # 开发环境
    static_dir = os.path.join(os.path.dirname(__file__), '..', '..', 'frontend', 'dist')

# 配置 CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # 允许所有来源
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 注册路由
app.include_router(excel_router, prefix="/api")

# 静态文件服务
if os.path.exists(static_dir):
    app.mount("/assets", StaticFiles(directory=os.path.join(static_dir, "assets")), name="assets")
    
    @app.get("/")
    async def serve_frontend():
        """服务前端页面"""
        index_path = os.path.join(static_dir, "index.html")
        if os.path.exists(index_path):
            return FileResponse(index_path)
        return {"message": "Excel 列删除工具 API", "version": "1.0.0", "docs": "/docs"}
else:
    @app.get("/")
    async def root():
        """根路径，返回 API 信息"""
        return {
            "message": "Excel 列删除工具 API",
            "version": "1.0.0",
            "docs": "/docs"
        }

# 确保上传目录存在
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

@app.get("/health")
async def health_check():
    """健康检查接口"""
    return {"status": "healthy"}

# 全局异常处理
@app.exception_handler(Exception)
async def global_exception_handler(request, exc):
    """全局异常处理器"""
    logger.error(f"未处理的异常: {str(exc)}", exc_info=True)
    return JSONResponse(
        status_code=500,
        content={"detail": "服务器内部错误"}
    )

def get_free_port():
    """获取可用端口"""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('', 0))
        s.listen(1)
        port = s.getsockname()[1]
    return port

if __name__ == "__main__":
    import uvicorn
    
    # 获取可用端口
    port = 8001
    try:
        # 尝试使用默认端口
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.bind(('localhost', port))
    except OSError:
        # 端口被占用，使用随机端口
        port = get_free_port()
    
    print(f"\\n🚀 Excel Column Remover started successfully!")
    print(f"📱 Access URL: http://localhost:{port}")
    print(f"📚 API Docs: http://localhost:{port}/docs")
    print(f"❤️  Press Ctrl+C to stop service\\n")
    
    try:
        uvicorn.run(
            app,
            host="0.0.0.0",
            port=port,
            log_level="info"
        )
    except KeyboardInterrupt:
        print("\\n👋 Service stopped by user")
    except Exception as e:
        print(f"\\n❌ Error: {e}")
        input("Press Enter to exit...")
'''
    
    with open(temp_dir / "main.py", 'w', encoding='utf-8') as f:
        f.write(main_content)
    
    return temp_dir

def build_simple_executable():
    """构建简单的可执行文件"""
    print("🚀 开始构建简化版可执行文件...")
    
    # 检查前端构建
    if not Path("frontend/dist").exists():
        print("❌ 前端构建不存在，请先运行: npm run build")
        return False
    
    # 创建临时构建目录
    temp_dir = create_simple_main()
    print("✅ 创建临时构建目录")
    
    # 安装PyInstaller
    try:
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)
        print("✅ PyInstaller 已准备就绪")
    except subprocess.CalledProcessError:
        print("❌ PyInstaller 安装失败")
        return False
    
    # 构建可执行文件
    try:
        cmd = [
            sys.executable, "-m", "PyInstaller",
            "--onefile",
            "--console",
            "--name", "ExcelProcessor",
            "--add-data", "frontend/dist;static",
            "--add-data", "sample-data.xlsx;.",
            "--hidden-import", "uvicorn.lifespan.on",
            "--hidden-import", "uvicorn.lifespan.off", 
            "--hidden-import", "uvicorn.protocols.websockets.auto",
            "--hidden-import", "uvicorn.protocols.http.auto",
            "--hidden-import", "uvicorn.loops.auto",
            str(temp_dir / "main.py")
        ]
        
        subprocess.run(cmd, check=True)
        print("✅ 可执行文件构建成功")
        
        # 清理临时目录
        shutil.rmtree(temp_dir)
        print("✅ 清理临时文件")
        
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"❌ 构建失败: {e}")
        return False

def create_portable_package():
    """创建便携包"""
    print("📦 创建便携包...")
    
    # 创建便携包目录
    portable_dir = Path("ExcelProcessor_Simple")
    if portable_dir.exists():
        shutil.rmtree(portable_dir)
    
    portable_dir.mkdir()
    
    # 复制可执行文件
    exe_path = Path("dist/ExcelProcessor.exe")
    if exe_path.exists():
        shutil.copy2(exe_path, portable_dir / "ExcelProcessor.exe")
        print("✅ 复制可执行文件")
    else:
        print("❌ 找不到可执行文件")
        return None
    
    # 复制示例文件
    if Path("sample-data.xlsx").exists():
        shutil.copy2("sample-data.xlsx", portable_dir / "sample-data.xlsx")
    
    # 创建启动脚本
    start_script = '''@echo off
title Excel Column Remover
echo.
echo ========================================
echo          Excel Column Remover
echo          Portable Version v1.0  
echo ========================================
echo.
echo Starting service, please wait...
echo.

ExcelProcessor.exe

echo.
echo Program stopped
echo To use again, double-click this file
echo.
pause
'''
    
    with open(portable_dir / "Start Tool.bat", 'w', encoding='utf-8') as f:
        f.write(start_script)
    
    # 创建说明文件
    readme_content = '''Excel Column Remover - Portable Version

=== Quick Start ===

1. Double-click "Start Tool.bat" to start service
2. Copy the displayed URL to browser
3. Select Excel file, check columns to delete
4. Click process button, save new file

=== Features ===

✓ Support .xlsx and .xls formats
✓ Visual column selection interface
✓ Preserve original styles and formats
✓ Handle merged cells properly
✓ Mobile browser compatible
✓ Original file remains unchanged

=== Notes ===

• First startup may require firewall authorization
• Close the black window when finished
• Recommended browsers: Chrome or Edge
• Supports mobile browser access

=== Troubleshooting ===

Q: Startup failed?
A: Check firewall settings, allow program to run

Q: Browser won't open?
A: Confirm URL is copied correctly, try refreshing

Q: Processing large files is slow?
A: This is normal, please be patient

Contact technical support if you have issues.
'''
    
    with open(portable_dir / "README.txt", 'w', encoding='utf-8') as f:
        f.write(readme_content)
    
    print(f"✅ 便携包创建成功: {portable_dir}")
    return portable_dir

if __name__ == "__main__":
    if build_simple_executable():
        portable_dir = create_portable_package()
        if portable_dir:
            print(f"""
🎉 简化版可执行文件打包完成！

📦 便携包位置: {portable_dir.absolute()}
🚀 使用方法: 双击 "Start Tool.bat"

📁 包含文件:
- ExcelProcessor.exe (主程序)
- Start Tool.bat (启动脚本)
- sample-data.xlsx (示例文件)
- README.txt (使用说明)

✨ 现在可以将整个文件夹复制到任何 Windows 电脑上使用！
""")
    else:
        print("❌ 构建失败")