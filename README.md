# 工资表生成系统 (Salary Sheet Maker)

一个基于 Streamlit 的智能工资表生成系统，支持自动合并休假数据和加班数据，并智能判断考勤情况。

## ✨ 功能特点

- 📊 **自动数据合并**：智能合并工资表模板、休假数据和加班数据
- 🎯 **智能考勤判断**：自动根据休假类型判断考勤情况和全勤工资
- 📱 **现代化界面**：基于 Streamlit 的直观用户界面
- 📥 **文件上传支持**：支持 Excel 文件上传和处理
- 💾 **一键下载**：生成的工资表可直接下载
- 🔄 **实时预览**：数据处理过程实时显示

## 🚀 快速开始

### 本地运行

1. **克隆项目**
   ```bash
   git clone https://github.com/yourusername/salarysheet_maker.git
   cd salarysheet_maker
   ```

2. **安装依赖**
   ```bash
   pip install -r requirements.txt
   ```

3. **运行应用**
   ```bash
   streamlit run salary_generator.py
   ```

4. **访问应用**
   - 打开浏览器访问 `http://localhost:8501`

### 使用步骤

1. **准备模板文件**：确保 `工资表模板.xlsx` 在项目根目录
2. **上传数据文件**：
   - 休假表（可选）：包含员工休假信息
   - 加班表（可选）：包含员工加班信息
3. **生成工资表**：点击"生成工资表"按钮
4. **下载结果**：下载生成的工资表文件

## 🌐 在线部署

### 推荐平台

1. **Streamlit Cloud**（推荐）
   - 免费且易用
   - 与 GitHub 无缝集成
   - 访问：[share.streamlit.io](https://share.streamlit.io)

2. **Vercel**
   - 支持自定义域名
   - 全球 CDN 加速

3. **Heroku**
   - 传统 PaaS 平台
   - 支持自定义域名

详细部署指南请参考 [DEPLOYMENT.md](DEPLOYMENT.md)

## 📁 文件说明

```
工资表/
├── salary_generator.py    # 主程序文件
├── requirements.txt       # Python 依赖
├── README.md             # 项目说明
├── DEPLOYMENT.md         # 部署指南
├── .gitignore           # Git 忽略文件
├── vercel.json          # Vercel 配置
├── Procfile             # Heroku 配置
├── api/
│   └── index.py         # API 入口文件
├── 工资表模板.xlsx        # 工资表模板
├── 休假表模板.xlsx        # 休假表模板
└── 加班表模板.xlsx        # 加班表模板
```

## 🛠 技术栈

- **前端框架**：Streamlit
- **数据处理**：Pandas
- **文件处理**：openpyxl, xlrd
- **部署平台**：Streamlit Cloud, Vercel, Heroku

## 📋 系统要求

- Python 3.7+
- 支持的文件格式：Excel (.xlsx, .xls)
- 浏览器：Chrome, Firefox, Safari, Edge

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

## 📄 许可证

MIT License