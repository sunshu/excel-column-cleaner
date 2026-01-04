# Excel 列删除工具

一个可本地部署的 Excel 处理项目，支持删除指定列并保持原有样式和合并单元格。

## 技术架构

- **前端**: React + Vite + Hooks
- **后端**: Python FastAPI + openpyxl
- **架构**: 前后端分离

## 功能特性

- 支持 .xlsx 和 .xls 格式文件
- 可视化列选择：勾选要删除的列，支持预览数据
- 保留 Excel 原有样式和格式
- 保持合并单元格结构
- 使用浏览器 File System Access API 保存文件
- 原始文件不被覆盖
- 响应式界面设计，支持移动端

## 最终交付清单

#### 便携版文件夹：`ExcelProcessor_Simple`
- `ExcelProcessor.exe` - 主程序（约150MB）
- `Start Tool.bat` - 启动脚本
- `sample-data.xlsx` - 示例文件
- `README.txt` - 英文使用说明

#### 使用方法
1. 复制整个 `ExcelProcessor_Simple` 文件夹到任何电脑
2. 双击 `Start Tool.bat` 启动工具
3. 复制显示的网址到浏览器打开
4. 上传Excel文件，勾选要删除的列，处理并保存

### 开发者文件

#### 源代码
- `frontend/` - React前端源码
- `backend/` - FastAPI后端源码
- `sample-data.xlsx` - 示例Excel文件

#### 开发脚本
- `start-all.bat` - 一键启动开发环境
- `start-backend.bat` - 启动后端服务
- `start-frontend.bat` - 启动前端服务
- `create-simple-executable.py` - 重新构建可执行文件

## 项目结构

```
excel-processor/
├── frontend/                    # React 前端
│   ├── src/
│   │   ├── components/
│   │   │   └── ExcelProcessor.jsx    # 主处理组件
│   │   ├── services/
│   │   │   └── api.js               # API 服务
│   │   ├── App.jsx                  # 主应用
│   │   ├── App.css                  # 样式文件
│   │   └── main.jsx                 # 入口文件
│   ├── dist/                        # 构建输出
│   ├── package.json                 # 前端依赖
│   └── vite.config.js              # Vite 配置
├── backend/                     # FastAPI 后端
│   ├── app/
│   │   ├── controllers/
│   │   │   └── excel_controller.py  # Excel API 控制器
│   │   ├── services/
│   │   │   └── excel_service.py     # Excel 业务逻辑
│   │   ├── utils/
│   │   │   └── file_utils.py        # 文件工具函数
│   │   └── main.py                  # FastAPI 主程序
│   └── requirements.txt             # Python 依赖
├── ExcelProcessor_Simple/           # 便携版可执行文件
│   ├── ExcelProcessor.exe          # 主程序
│   ├── Start Tool.bat              # 启动脚本
│   ├── sample-data.xlsx            # 示例文件
│   └── README.txt                  # 英文说明
├── sample-data.xlsx                # 示例Excel文件
├── start-all.bat                   # 开发环境启动脚本
├── start-backend.bat               # 后端启动脚本
├── start-frontend.bat              # 前端启动脚本
└── README.md                       # 项目说明
```

## 快速开始

### 环境要求

- **Python 3.8+** （用于后端服务）
- **Node.js 16+** （用于前端构建）
- **现代浏览器** Chrome 86+ 或 Edge 86+（支持文件保存 API）

### 方法一：一键启动（推荐）
双击 `start-all.bat` 文件，系统会自动：
1. 检查 Python 和 Node.js 环境
2. 安装后端依赖
3. 启动后端服务（端口 8001）
4. 安装前端依赖  
5. 启动前端服务（端口 3000+）

### 方法二：分别启动
1. **启动后端**：双击 `start-backend.bat`
2. **启动前端**：双击 `start-frontend.bat`

### 方法三：手动启动
```bash
# 启动后端
cd backend
pip install -r requirements.txt
cd app
python main.py

# 启动前端（新终端）
cd frontend
npm install
npm run dev
```

## 使用步骤

1. **访问应用**
   - 打开浏览器访问前端地址（启动时会显示具体端口）

2. **选择文件**
   - 点击"选择 Excel 文件"按钮
   - 选择 .xlsx 或 .xls 格式的文件
   - 系统会自动预览文件的所有列信息

3. **选择要删除的列**
   - 查看每列的名称和示例数据
   - 勾选要删除的列（支持多选）
   - 可以使用"全选/取消全选"快速操作

4. **处理文件**
   - 点击"删除选中的X列并保存"按钮
   - 等待处理完成

5. **保存文件**
   - 系统会弹出保存对话框
   - 选择保存位置和文件名
   - 原文件不会被修改

## 界面特性

### 移动端适配
- **响应式设计**：自动适配手机、平板、电脑屏幕
- **触摸友好**：大按钮设计，易于点击
- **流畅动画**：现代化的交互效果

### 列预览功能
- **自动预览**：选择文件后自动显示所有列
- **列信息展示**：显示列名、列索引、示例数据
- **可视化选择**：通过勾选框选择要删除的列
- **批量操作**：支持全选/取消全选

### 用户体验
- **实时反馈**：操作状态实时显示
- **错误提示**：详细的错误信息和解决建议
- **进度显示**：处理过程中显示加载动画
- **成功确认**：操作完成后的明确提示

## 示例测试

项目包含一个示例文件 `sample-data.xlsx`，包含以下列：

| 列号 | 列名 | 示例数据 |
|------|------|----------|
| 1 | 姓名 | 张三, 李四, 王五 |
| 2 | 年龄 | 28, 32, 25 |
| 3 | 部门 | 技术部, 技术部, 设计部 |
| 4 | 职位 | 软件工程师, 高级工程师, UI设计师 |
| 5 | 工资 | 12000, 18000, 10000 |
| 6 | 入职日期 | 2023-01-15, 2022-03-20, 2023-05-10 |
| 7 | 备注 | React开发, Python后端, 界面设计 |

**测试建议：**
- 删除部门列：勾选"部门"
- 删除部门和工资列：勾选"部门"和"工资"
- 删除多个列：勾选"年龄"、"职位"、"入职日期"

## 离线部署指南

### 部署方案选择

根据不同的使用场景，提供多种离线部署方案：

#### 方案对比

| 方案 | 优点 | 缺点 | 适用场景 |
|------|------|------|----------|
| 可执行文件版本 | 双击即用，无需环境 | 文件较大(~150MB) | 普通用户，简单部署 |
| 开发环境版本 | 完整代码，可定制 | 需要Python/Node.js | 开发者，可修改 |

### 方案一：可执行文件版本（推荐）

#### 特点
- 双击即用，无需安装任何环境
- 包含完整的前端和后端
- 自动选择可用端口
- 适合普通用户

#### 构建步骤
```bash
# 1. 运行构建脚本
python create-simple-executable.py

# 2. 获得便携包文件夹
ExcelProcessor_Simple/
```

#### 使用方法
1. 构建完成后获得 `ExcelProcessor_Simple` 文件夹
2. 将整个文件夹复制到目标电脑
3. 双击 `Start Tool.bat` 即可使用
4. 浏览器访问显示的地址

### 方案二：开发环境部署

#### 特点
- 包含完整源代码
- 可以修改和定制
- 支持开发模式
- 需要Python和Node.js环境

#### 部署步骤
```bash
# 1. 启动后端
cd backend
pip install -r requirements.txt
cd app
python main.py

# 2. 启动前端（新终端）
cd frontend
npm install
npm run dev
```

## API 文档

### 预览列接口

**POST** `/api/excel/preview`

**请求参数:**
- `file`: Excel 文件（multipart/form-data）

**响应:**
- 成功：返回包含列信息的JSON对象
- 失败：返回错误信息

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

### 其他接口
- `GET /`: API 基本信息
- `GET /health`: 健康检查
- `GET /api/excel/info`: Excel API 信息
- `GET /docs`: Swagger API 文档

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

### 高级配置

#### 自定义端口
在可执行文件版本中，程序会自动选择可用端口。如需固定端口，可以修改源代码中的端口配置。

#### 修改界面
所有版本都支持修改前端界面：
1. 修改 `frontend/src/` 下的文件
2. 重新构建前端：`npm run build`
3. 重新打包

#### 添加功能
可以在后端添加新的API接口：
1. 在 `backend/app/controllers/` 添加控制器
2. 在 `backend/app/services/` 添加业务逻辑
3. 重新打包

## 浏览器兼容性

- Chrome 86+
- Edge 86+
- Firefox（不支持 File System Access API，会降级到下载）
- Safari（不支持 File System Access API，会降级到下载）

### 移动端使用

#### 手机浏览器
1. 使用 Chrome 或 Safari 浏览器
2. 访问前端地址
3. 界面会自动适配手机屏幕
4. 支持触摸操作和手势

#### 平板设备
1. 支持横屏和竖屏模式
2. 列信息以网格形式展示
3. 触摸友好的大按钮设计

## 故障排除

### 常见问题

**1. 端口被占用**
```
错误：Address already in use
解决：系统会自动选择可用端口，查看启动日志获取实际端口
```

**2. 文件预览失败**
```
错误：预览失败
解决：检查文件格式是否正确，确保是有效的Excel文件
```

**3. 列信息显示异常**
```
错误：列信息不完整
解决：确保Excel文件第一行包含列标题
```

**4. 移动端显示问题**
```
错误：页面显示不正常
解决：使用现代浏览器，清除缓存后重新访问
```

**5. 文件保存失败**
```
错误：showSaveFilePicker is not defined
解决：使用 Chrome 86+ 或 Edge 86+ 浏览器
```

**6. 可执行文件启动失败**
- 检查防火墙设置
- 确保端口未被占用
- 以管理员身份运行

**7. 开发环境安装失败**
- 确保Python和Node.js已正确安装
- 检查网络连接（首次安装需要下载依赖）
- 尝试手动安装依赖

**8. Excel 处理错误**
- 检查文件格式和列索引是否正确
- 确认文件没有被其他程序打开
- 查看浏览器控制台错误信息

### 性能优化

**减少文件大小**
- 删除不需要的示例文件
- 压缩前端资源
- 使用生产环境构建

**提高启动速度**
- 使用SSD存储
- 关闭不必要的防病毒软件扫描
- 预热Python环境

### 日志查看

- **前端日志**：浏览器开发者工具 Console
- **后端日志**：终端输出
- **可执行文件版本**：控制台输出
- **开发环境版本**：终端输出

## 技术支持

如遇到问题，请提供：
1. 使用的部署方案
2. 操作系统版本
3. 错误信息截图
4. 相关日志内容

## 许可证

MIT License