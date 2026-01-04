# Excel 处理后端服务

基于 FastAPI 的 Excel 列删除服务。

## 功能

- 接收 Excel 文件上传
- 删除指定列
- 保持原有样式和合并单元格
- 返回处理后的文件

## 安装依赖

```bash
pip install -r requirements.txt
```

## 启动服务

```bash
cd app
python main.py
```

服务将在 http://localhost:8000 启动

## API 文档

启动后访问 http://localhost:8000/docs 查看 Swagger 文档