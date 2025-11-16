# Qconverto - 多媒体文件格式转换工具

Qconverto 是一个基于 NiceGUI 的多媒体文件格式转换工具，支持图片、音频、视频和文档格式的转换。

## 功能特点

- 支持多种文件格式转换
- 现代化的用户界面
- 拖拽上传文件
- 实时转换进度显示
- 在线预览和下载转换结果

## 支持的格式

### 图片
- 输入: JPG, JPEG, PNG, BMP, GIF, WEBP, ICO
- 输出: JPG, JPEG, PNG, WEBP, PDF

### 音频
- 输入: MP3, WAV, FLAC, OGG, M4A, MP4, AAC, APE, WV
- 输出: MP3, WAV, FLAC, OGG, M4A

### 视频
- 输入: MP4, AVI, MOV, MKV, WMV, FLV
- 输出: MP4, AVI, MOV, MKV, WMV

### 文档
- 输入: PDF, DOC, DOCX, TXT
- 输出: PDF, DOCX, TXT, JPG

## 部署到 Railway

本项目已经配置好可以直接部署到 Railway 平台。

### 部署步骤

1. 在 GitHub 上 Fork 本项目或者推送你的代码到 GitHub 仓库
2. 访问 [Railway](https://railway.app/) 并登录或注册账户
3. 点击 "New Project"
4. 选择 "Deploy from GitHub repo"
5. 选择你的仓库
6. Railway 会自动检测这是一个 Python 项目并使用 Nixpacks 构建
7. 点击 "Deploy" 开始部署

### 配置说明

- 项目已包含 `railway.json` 配置文件，定义了构建和部署设置
- 项目已包含 `Procfile` 文件，告诉 Railway 如何运行应用
- 应用会自动使用 Railway 提供的 PORT 环境变量

## 本地开发

### 安装依赖

```bash
pip install -r requirements.txt
```

### 运行应用

```bash
python main.py
```

应用将在 http://localhost:8081 启动

## 技术栈

- Python 3.x
- NiceGUI - 用于构建现代Web界面
- Pillow - 图像处理
- PyMuPDF - PDF处理
- pydub - 音频处理
- moviepy - 视频处理

## 注意事项

- 视频转换可能需要较长时间，请耐心等待
- 大文件转换可能会消耗较多内存和处理时间
- 某些格式转换可能需要安装额外的系统依赖（如FFmpeg）

## 许可证

MIT License