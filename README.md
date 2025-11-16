# Qconverto - 多媒体文件格式转换工具

Qconverto 是一个基于 Python 和 NiceGUI 的多媒体文件格式转换工具，支持图片、音频、视频和文档格式的转换。

## 功能特点

- 支持多种文件格式转换
- 现代化的用户界面
- 拖拽上传文件
- 实时转换进度显示
- 支持在线预览和下载

## 支持的格式

### 图片
- 输入格式：JPG, JPEG, PNG, BMP, GIF, WEBP, ICO
- 输出格式：JPG, JPEG, PNG, WEBP, PDF

### 音频
- 输入格式：MP3, WAV, FLAC, OGG, M4A, MP4, AAC, APE, WV
- 输出格式：MP3, WAV, FLAC, OGG, M4A

### 视频
- 输入格式：MP4, AVI, MOV, MKV, WMV, FLV
- 输出格式：MP4, AVI, MOV, MKV, WMV

### 文档
- 输入格式：PDF, DOC, DOCX, TXT
- 输出格式：PDF, DOCX, TXT, JPG

## 本地运行

1. 安装依赖：
   ```bash
   pip install -r requirements.txt
   ```

2. 运行应用：
   ```bash
   python main.py
   ```

3. 在浏览器中访问 `http://localhost:8081`

## 部署到 Render

本项目已配置为可在 Render 平台上一键部署：

1. Fork 此仓库到你的 GitHub 账户
2. 登录 [Render](https://render.com/) 并连接你的 GitHub 账户
3. 点击 "New Web Service"
4. 选择你 Fork 的仓库
5. Render 会自动检测 `render.yaml` 配置文件
6. 点击 "Create Web Service" 完成部署

或者手动配置时使用以下设置：
- Build Command: `pip install -r requirements.txt`
- Start Command: `python app.py`
- Python Version: 3.9.15+

应用将在部署完成后通过 Render 提供的 URL 访问。

## 依赖说明

- NiceGUI: 用于构建现代 Web 界面
- Pillow: 图像处理
- PyMuPDF: PDF 处理
- pydub: 音频处理
- ffmpeg-python: 视频处理
- MoviePy: 视频处理备选方案

## 注意事项

- 视频转换可能需要较长时间，请耐心等待
- 大文件转换可能会消耗较多内存
- 某些格式转换可能需要额外安装系统级依赖（如 FFmpeg）

## 许可证

MIT License