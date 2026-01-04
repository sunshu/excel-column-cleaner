# Excel 列删除工具

一个可本地部署的 Excel 处理项目，支持删除指定列并保持原有样式和合并单元格。

## 技术架构

- **前端**: React + Vite + Hooks
- **后端**: Python FastAPI + openpyxl
- **架构**: 前后端分离

## 功能特性

- ✅ 支持 .xlsx 和 .xls 格式文件
- ✅ 删除指定列（支持多列，如：3,5,7）
- ✅ 保留 Excel 原有样式和格式
- ✅ 保持合并单元格结构
- ✅ 使用浏览器 File System Access API 保存文件
- ✅ 原始文件不被覆盖
- ✅ 响应式界面设计

## 项目结构

```
excel-processor/
├── frontend/                 # React 前端
│   ├── src/
│   │   ├── components/
│   │   │   └── ExcelProcessor.jsx
│   │   ├── services/
│   │   │   └── api.js
│   │   ├── App.jsx
│   │   ├── App.css
│   │   └── main.jsx
│   ├── package.json
│   ├── vite.config.js
│   └── index.html
├── backend/                  # FastAPI 后端
│   ├── app/
│   │   ├── controllers/
│   │   │   └── excel_controller.py
│   │   ├── services/
│   │   │   └── excel_service.py
│   │   ├── utils/
│   │   │   └── file_utils.py
│   │   └── main.py
│   └── requirements.txt
├── start-frontend.bat        # 前端启动脚本
├── start-backend.bat         # 后端启动脚本
└── README.md
```

## 快速开始

### 环境要求

- Node.js 16+ 
- Python 3.8+
- Chrome/Edge 浏览器（支持 File System Access API）

### 1. 启动后端服务

```bash
# 方式一：使用批处理文件（Windows）
双击 start-backend.bat

# 方式二：手动启动
cd backend
pip install -r requirements.txt
cd app
python main.py
```

后端服务将在 `http://localhost:8000` 启动

### 2. 启动前端服务

```bash
# 方式一：使用批处理文件（Windows）
双击 start-frontend.bat

# 方式二：手动启动
cd frontend
npm install
npm run dev
```

前端服务将在 `http://localhost:3000` 启动

### 3. 使用应用

1. 打开浏览器访问 `http://localhost:3000`
2. 选择要处理的 Excel 文件
3. 输入要删除的列索引（如：3,5,7）
4. 点击"删除列并保存"按钮
5. 在弹出的保存对话框中选择保存位置

## API 文档

### 删除列接口

**POST** `/api/excel/delete-columns`

**请求参数:**
- `file`: Excel 文件（multipart/form-data）
- `columns`: 要删除的列索引字符串（如："3,5,7"）

**响应:**
- 成功：返回处理后的 Excel 文件
- 失败：返回错误信息

**示例:**
```javascript
const formData = new FormData()
formData.append('file', file)
formData.append('columns', '3,5')

const response = await fetch('/api/excel/delete-columns', {
  method: 'POST',
  body: formData
})
```

## 开发说明

### 前端技术要点

- 使用 React Hooks 管理状态
- File System Access API 实现本地文件保存
- 表单验证和错误处理
- 响应式 UI 设计

### 后端技术要点

- FastAPI 异步框架
- openpyxl 处理 Excel 文件
- 保持样式和合并单元格
- 分层架构（Controller/Service/Utils）

### 核心功能实现

1. **文件上传**: 使用 FastAPI 的 `UploadFile` 处理文件上传
2. **列删除**: 使用 openpyxl 的 `delete_cols()` 方法
3. **样式保持**: 处理合并单元格的重新计算
4. **文件保存**: 使用浏览器原生 API 实现本地保存

## 浏览器兼容性

- ✅ Chrome 86+
- ✅ Edge 86+
- ❌ Firefox（不支持 File System Access API，会降级到下载）
- ❌ Safari（不支持 File System Access API，会降级到下载）

## 故障排除

### 常见问题

1. **端口占用**: 修改 `vite.config.js` 或 `main.py` 中的端口配置
2. **CORS 错误**: 检查后端 CORS 配置是否包含前端地址
3. **文件保存失败**: 确保使用支持的浏览器
4. **Excel 处理错误**: 检查文件格式和列索引是否正确

### 日志查看

- 前端：浏览器开发者工具 Console
- 后端：终端输出或日志文件

## 许可证

MIT License