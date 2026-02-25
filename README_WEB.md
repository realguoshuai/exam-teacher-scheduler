# 监考老师排班系统 - Web版

## 功能特性

- 🌐 **Web可视化界面** - 现代化网页操作，无需命令行
- 📊 **实时数据展示** - 动态表格展示排班结果
- 📥 **一键导出** - Excel格式导出排班结果
- 🔄 **自动刷新** - 数据自动加载和更新
- 📱 **响应式设计** - 支持桌面和移动设备
- 🎨 **美观界面** - 渐变色设计，卡片式布局

## 快速开始

### 方式一：一键启动（推荐）

```batch
双击 start_web.bat
```

菜单选项：
1. 首次使用 - 安装Web依赖
2. 运行Web界面 (http://localhost:5000)
3. 运行CLI版本
4. 打开数据目录
5. 在浏览器中打开
0. 退出

### 方式二：分步安装

```batch
1. 双击 install_web.bat 安装Web依赖
2. 双击 run_web.bat 启动Web服务
```

### 方式三：手动启动

```bash
# 1. 安装依赖
pip install -r requirements_web.txt

# 2. 启动Web服务
python app.py

# 3. 浏览器访问
http://localhost:5000
```

## Web界面功能

### 排班管理
- 🔄 初始化数据 - 创建示例数据
- 🚀 自动排班 - 智能分配监考老师
- 📥 导出Excel - 下载排班结果
- 🗑️ 重置数据 - 清除所有数据

### 监考老师
- 📋 查看老师列表
- 🔄 刷新数据
- 📊 显示监考次数

### 考试信息
- 📋 查看考试列表
- 🔄 刷新数据
- 📅 显示考试时间

### 统计分析
- 📊 排班统计
- 👥 老师监考统计
- 📈 数据可视化

## API接口

### 健康检查
```
GET /api/health
```

### 数据初始化
```
GET /api/init
```

### 获取监考老师
```
GET /api/teachers
```

### 获取考试信息
```
GET /api/exams
```

### 执行排班
```
POST /api/schedule
```

### 获取排班结果
```
GET /api/schedule
```

### 获取统计信息
```
GET /api/statistics
```

### 导出Excel
```
GET /api/export
```

### 重置数据
```
POST /api/reset
```

## 技术栈

- **后端**: Python 3.8+ + Flask 3.0.0
- **前端**: HTML5 + CSS3 + JavaScript (原生)
- **数据处理**: pandas + openpyxl
- **API**: Flask RESTful API
- **跨域**: Flask-CORS

## 项目结构

```
exam-teacher-scheduler/
├── app.py                    # Web应用主程序
├── templates/
│   └── index.html            # Web界面
├── requirements_web.txt      # Web依赖包
├── install_web.bat          # 安装Web依赖
├── run_web.bat             # 运行Web服务
├── start_web.bat           # 一键启动
├── models.py               # 数据模型
├── scheduler.py            # 排班算法
├── utils.py                # 工具函数
└── data/                   # 数据目录
    ├── teachers.xlsx
    ├── exams.xlsx
    └── schedule.xlsx
```

## 浏览器支持

- ✅ Chrome 90+
- ✅ Firefox 88+
- ✅ Safari 14+
- ✅ Edge 90+
- ✅ 移动浏览器

## 常见问题

### Q: 如何修改端口号？
A: 编辑 `app.py` 文件最后一行：
```python
app.run(debug=True, host='0.0.0.0', port=5000)  # 修改 port 参数
```

### Q: 如何自定义数据？
A: 编辑 `data/teachers.xlsx` 和 `data/exams.xlsx`

### Q: Web版本和CLI版本有什么区别？
A: Web版本提供可视化界面，CLI版本提供命令行界面，功能相同

### Q: 支持多用户吗？
A: 当前版本为单用户版本，支持同一局域网访问

### Q: 如何停止Web服务？
A: 在命令行窗口按 `Ctrl+C`

## 开发说明

### 本地开发
```bash
# 安装依赖
pip install -r requirements_web.txt

# 启动开发服务器
python app.py
```

### 调试模式
```python
app.run(debug=True, host='0.0.0.0', port=5000)
```

## 更新日志

### v1.0.0 (2026-01-30)
- ✅ 初始版本发布
- ✅ Web可视化界面
- ✅ RESTful API
- ✅ Excel导出功能
- ✅ 实时数据展示

## 联系方式

- 项目地址: D:\code\git\exam-teacher-scheduler
- 版本: 1.0.0
- 技术支持: AI Assistant
