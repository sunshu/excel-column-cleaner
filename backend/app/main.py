"""
FastAPI 主应用程序 - 打包版本
Excel 列删除工具后端服务，包含静态文件服务
"""

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import JSONResponse, FileResponse
import os
import sys
import logging
from pathlib import Path

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

if __name__ == "__main__":
    import uvicorn
    
    # 获取可用端口
    import socket
    def get_free_port():
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.bind(('', 0))
            s.listen(1)
            port = s.getsockname()[1]
        return port
    
    port = 8001
    try:
        # 尝试使用默认端口
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.bind(('localhost', port))
    except OSError:
        # 端口被占用，使用随机端口
        port = get_free_port()
    
    print(f"\n🚀 Excel 列删除工具启动成功！")
    print(f"📱 访问地址: http://localhost:{port}")
    print(f"📚 API 文档: http://localhost:{port}/docs")
    print(f"❤️  按 Ctrl+C 停止服务\n")
    
    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=port,
        reload=False,
        log_level="info"
    )
