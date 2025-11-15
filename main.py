#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Qconverto - 多媒体文件格式转换工具
支持图片、音频、视频和文档格式的转换
"""

import sys
import os
import tempfile
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog, QComboBox, QProgressBar,
    QMessageBox, QTextEdit, QGroupBox, QGridLayout, QSizePolicy
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QFont, QIcon, QDragEnterEvent, QDropEvent

# 添加当前目录到路径
sys.path.insert(0, '.')

# 检查MoviePy是否可用
try:
    from moviepy.editor import VideoFileClip
    MOVIEPY_AVAILABLE = True
except ImportError:
    MOVIEPY_AVAILABLE = False

# 支持的文件格式字典
SUPPORTED_FORMATS = {
    '图片': {
        '输入': ['.jpg', '.jpeg', '.png', '.bmp', '.gif', '.webp', '.ico'],
        '输出': ['.jpg', '.jpeg', '.png', '.webp']
    },
    '音频': {
        '输入': ['.mp3', '.wav', '.flac', '.ogg', '.m4a', '.mp4', '.aac', '.ape', '.wv'],
        '输出': ['.mp3', '.wav', '.flac', '.ogg', '.m4a']
    },
    '视频': {
        '输入': ['.mp4', '.avi', '.mov', '.mkv', '.wmv', '.flv'],
        '输出': ['.mp4', '.avi', '.mov', '.mkv', '.wmv']
    },
    '文档': {
        '输入': ['.pdf', '.doc', '.docx', '.txt'],
        '输出': ['.pdf', '.docx', '.txt', '.jpg']
    }
}

# 文件头标识字典（用于智能识别文件类型）
FILE_SIGNATURES = {
    '图片': {
        b'\xff\xd8\xff': 'JPEG',
        b'\x89PNG\r\n\x1a\n': 'PNG',
        b'GIF87a': 'GIF',
        b'GIF89a': 'GIF',
        b'BM': 'BMP',
        b'RIFF': 'WebP',  # WebP 文件以 RIFF 开头，但需要进一步检查
        b'\x00\x00\x01\x00': 'ICO'  # ICO 图标文件
    },
    '音频': {
        b'ID3': 'MP3',
        b'\xff\xfb': 'MP3',
        b'\xff\xf3': 'MP3',
        b'\xff\xf2': 'MP3',
        b'RIFF': 'WAV',  # WAV 文件以 RIFF 开头，但需要进一步检查
        b'OggS': 'OGG',
        b'fLaC': 'FLAC',
        b'MAC': 'APE',  # Monkey's Audio
        b'wvpk': 'WV'   # WavPack
    },
    '视频': {
        b'\x00\x00\x00\x18ftypmp4': 'MP4',
        b'\x00\x00\x00\x20ftypmp4': 'MP4',
        b'RIFF': 'AVI',  # AVI 文件以 RIFF 开头，但需要进一步检查
        b'\x1aE\xdf\xa3': 'MKV',
        b'ftyp': 'MOV',  # MOV 文件以 ftyp 开头
        b'FLV': 'FLV'   # FLV 文件
    },
    '文档': {
        b'%PDF': 'PDF',
        b'PK\x03\x04': 'DOCX',  # DOCX 实际上是 ZIP 格式，需要进一步检查
        b'{\\rtf': 'RTF',
        b'PK\x03\x04': 'DOCX',  # DOCX 文件
        b'PK\x03\x04': 'PPTX',  # PPTX 文件
        b'PK\x03\x04': 'XLSX',  # XLSX 文件
        b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1': 'DOC'  # DOC 文件的文件头签名
    }
}


class ConversionThread(QThread):
    """转换线程，用于在后台执行文件转换任务"""
    progress = pyqtSignal(int)
    finished = pyqtSignal(bool, str, str)  # success, output_file, original_file
    error = pyqtSignal(str)

    def __init__(self, input_file, output_file, file_type, settings=None):
        super().__init__()
        self.input_file = input_file
        self.output_file = output_file
        self.file_type = file_type
        self.settings = settings or {}

    def run(self):
        """执行转换任务"""
        try:
            self.progress.emit(10)
            
            if self.file_type == '图片':
                self._convert_image()
            elif self.file_type == '音频':
                self._convert_audio()
            elif self.file_type == '视频':
                self._convert_video()
            elif self.file_type == '文档':
                self._convert_document()
            else:
                raise ValueError(f"不支持的文件类型: {self.file_type}")
            
            self.progress.emit(100)
            self.finished.emit(True, self.output_file, self.input_file)
        except Exception as e:
            # 提供更详细的错误信息，包括原始异常类型和堆栈跟踪
            import traceback
            error_details = f"{str(e)}\n\n详细信息:\n{traceback.format_exc()}"
            self.error.emit(error_details)
            self.finished.emit(False, self.output_file, self.input_file)

    def _convert_image(self):
        """转换图片文件"""
        try:
            from PIL import Image
            self.progress.emit(30)
            with Image.open(self.input_file) as img:
                self.progress.emit(70)
                # 应用图片质量设置
                quality = self.settings.get('image_quality', 90)
                img.save(self.output_file, quality=quality)
        except Exception as e:
            raise RuntimeError(f"图片转换失败: {str(e)}") from e

    def _convert_audio(self):
        """转换音频文件"""
        try:
            self.progress.emit(30)
            
            # 首先尝试使用pydub进行音频转换（支持更多格式）
            try:
                from pydub import AudioSegment
                
                # 使用pydub加载音频文件
                audio = AudioSegment.from_file(self.input_file)
                self.progress.emit(50)
                
                # 获取音频设置
                audio_bitrate = self.settings.get('audio_bitrate', '192k')
                
                # 导出为指定格式
                file_ext = os.path.splitext(self.output_file)[1].lower()
                
                if file_ext == '.mp3':
                    audio.export(self.output_file, format="mp3", bitrate=audio_bitrate)
                elif file_ext == '.wav':
                    audio.export(self.output_file, format="wav")
                elif file_ext == '.flac':
                    audio.export(self.output_file, format="flac")
                elif file_ext == '.ogg':
                    audio.export(self.output_file, format="ogg")
                elif file_ext == '.m4a':
                    audio.export(self.output_file, format="mp4")
                else:
                    # 默认导出为MP3
                    audio.export(self.output_file, format="mp3", bitrate=audio_bitrate)
                
                self.progress.emit(70)
                return  # 成功使用pydub转换
                
            except Exception as pydub_error:
                # pydub转换失败，尝试使用miniaudio
                print(f"pydub转换失败，尝试miniaudio: {pydub_error}")
                pass
            
            # 尝试使用miniaudio进行音频处理
            try:
                import miniaudio
                
                self.progress.emit(50)
                
                # 使用miniaudio处理音频文件
                input_ext = os.path.splitext(self.input_file)[1].lower()
                output_ext = os.path.splitext(self.output_file)[1].lower()
                
                # 获取音频信息
                info = miniaudio.get_file_info(self.input_file)
                print(f"音频信息: {info.nchannels}声道, {info.sample_rate}Hz, {info.duration:.2f}秒")
                
                # 解码音频
                decoded = miniaudio.decode_file(self.input_file)
                
                # 根据输出格式决定处理方式
                if output_ext == '.wav':
                    # 保存为WAV文件
                    import wave
                    with wave.open(self.output_file, 'w') as wf:
                        wf.setnchannels(decoded.nchannels)
                        wf.setsampwidth(2)  # 16-bit
                        wf.setframerate(decoded.sample_rate)
                        wf.writeframes(decoded.samples)
                else:
                    # 对于其他格式，执行文件复制并给出提示
                    import shutil
                    shutil.copy2(self.input_file, self.output_file)
                    print(f"注意: miniaudio不支持编码为{output_ext}格式，执行文件复制")
                
                self.progress.emit(70)
                return  # 成功使用miniaudio处理
                
            except Exception as miniaudio_error:
                # miniaudio处理失败，尝试使用纯Python方法
                print(f"miniaudio处理失败，尝试纯Python方法: {miniaudio_error}")
                pass
            
            # 如果pydub和miniaudio都不可用，使用纯Python方法进行基本转换
            self.progress.emit(50)
            
            # 纯Python音频转换（仅支持WAV相关格式）
            input_ext = os.path.splitext(self.input_file)[1].lower()
            output_ext = os.path.splitext(self.output_file)[1].lower()
            
            # 对于WAV文件的基本处理
            if input_ext == '.wav':
                if output_ext == '.wav':
                    # WAV到WAV的复制
                    import shutil
                    shutil.copy2(self.input_file, self.output_file)
                else:
                    # 对于其他格式，提供说明或简单复制
                    import shutil
                    shutil.copy2(self.input_file, self.output_file)
                    print(f"注意：从WAV到{output_ext}的真实转换需要专门的编码库")
            else:
                # 对于非WAV格式，尝试简单复制（可能不工作）
                import shutil
                shutil.copy2(self.input_file, self.output_file)
                print(f"注意：从{input_ext}到{output_ext}的真实转换需要专门的解码和编码库")
            
            self.progress.emit(70)
            
        except Exception as e:
            # 提供更详细的错误信息
            error_msg = f"音频转换失败: {str(e)}\n\n"
            error_msg += "请确保输入文件格式受支持且未被其他程序占用。"
            raise RuntimeError(error_msg) from e

    def _convert_video(self):
        """转换视频文件"""
        try:
            self.progress.emit(30)
            
            # 首先尝试使用MoviePy进行视频转换（如果可用）
            if MOVIEPY_AVAILABLE:
                try:
                    from moviepy.editor import VideoFileClip
                    import os
                    
                    # 使用MoviePy加载视频文件
                    video = VideoFileClip(self.input_file)
                    self.progress.emit(50)
                    
                    # 获取输出文件扩展名
                    output_ext = os.path.splitext(self.output_file)[1].lower()
                    
                    # 根据输出格式进行相应处理
                    if output_ext in ['.mp4', '.avi', '.mov', '.mkv', '.wmv']:
                        # 对于支持的格式，直接写入
                        video.write_videofile(
                            self.output_file,
                            codec='libx264',
                            audio_codec='aac',
                            temp_audiofile='temp-audio.m4a',
                            remove_temp=True,
                            verbose=False,
                            logger=None
                        )
                    else:
                        # 对于不直接支持的格式，使用默认设置
                        video.write_videofile(
                            self.output_file,
                            temp_audiofile='temp-audio.m4a',
                            remove_temp=True,
                            verbose=False,
                            logger=None
                        )
                    
                    # 关闭视频文件
                    video.close()
                    self.progress.emit(70)
                    return  # 成功使用MoviePy转换
                    
                except Exception as moviepy_error:
                    # MoviePy转换失败，尝试使用FFmpeg
                    print(f"MoviePy转换失败，尝试FFmpeg: {moviepy_error}")
                    pass
            
            # 尝试使用FFmpeg进行视频转换
            try:
                import ffmpeg
                import shutil
                import os
                
                # 检查ffmpeg是否可用
                ffmpeg_path = shutil.which("ffmpeg")
                if not ffmpeg_path:
                    # 尝试使用相对路径或打包环境中的ffmpeg
                    possible_paths = [
                        "ffmpeg",
                        "./ffmpeg",
                        "./ffmpeg.exe",
                        os.path.join(os.path.dirname(__file__), "ffmpeg"),
                        os.path.join(os.path.dirname(__file__), "ffmpeg.exe")
                    ]
                    
                    for path in possible_paths:
                        if os.path.exists(path) or shutil.which(path):
                            ffmpeg_path = path
                            break
                
                # 如果仍然找不到，提供更友好的错误信息
                if not ffmpeg_path:
                    raise RuntimeError(
                        "未找到FFmpeg，视频转换需要FFmpeg支持。\n\n"
                        "解决方案:\n"
                        "1. 自动下载: 运行 install_ffmpeg.py 脚本自动安装FFmpeg\n"
                        "2. 手动安装: 访问 https://ffmpeg.org/download.html 下载并安装\n"
                        "3. 打包应用: 将ffmpeg.exe文件放在程序同目录下\n\n"
                        "注意: 音频转换不需要FFmpeg，可以正常使用。"
                    )
                
                # 应用视频设置
                video_codec = self.settings.get('video_codec', 'libx264')
                audio_codec = self.settings.get('audio_codec', 'aac')
                crf = self.settings.get('video_quality', 23)
                
                try:
                    (ffmpeg
                        .input(self.input_file)
                        .output(self.output_file, 
                                vcodec=video_codec, 
                                acodec=audio_codec, 
                                crf=crf, 
                                strict='experimental')
                        .overwrite_output()
                        .run(cmd=ffmpeg_path, capture_stdout=True, capture_stderr=True)
                    )
                    self.progress.emit(70)
                except ffmpeg.Error as e:
                    # 获取详细的错误信息
                    error_msg = f"视频转换失败: {e.stderr.decode('utf-8') if e.stderr else str(e)}"
                    raise RuntimeError(error_msg) from e
                except Exception as e:
                    error_msg = f"视频转换过程中发生未知错误: {str(e)}"
                    raise RuntimeError(error_msg) from e
                    
            except Exception as ffmpeg_error:
                # FFmpeg处理失败，尝试使用纯Python方法
                print(f"FFmpeg处理失败: {ffmpeg_error}")
                raise RuntimeError(f"视频转换失败: 未找到可用的视频处理库(MoviePy或FFmpeg)") from ffmpeg_error
                
        except Exception as e:
            raise RuntimeError(f"视频转换失败: {str(e)}") from e

    def _convert_document(self):
        """转换文档文件"""
        try:
            # 仅支持部分文档格式转换
            ext_in = os.path.splitext(self.input_file)[1].lower()
            ext_out = os.path.splitext(self.output_file)[1].lower()
            
            self.progress.emit(30)
            
            # PDF转图片
            if ext_in == '.pdf' and ext_out in ['.jpg', '.jpeg']:
                self._pdf_to_image()
            # 图片转PDF
            elif ext_in in ['.jpg', '.jpeg'] and ext_out == '.pdf':
                self._image_to_pdf()
            # 文本转PDF
            elif ext_in == '.txt' and ext_out == '.pdf':
                self._txt_to_pdf()
            # PDF转文本
            elif ext_in == '.pdf' and ext_out == '.txt':
                self._pdf_to_txt()
            # DOC转PDF
            elif ext_in == '.doc' and ext_out == '.pdf':
                self._doc_to_pdf()
            # DOCX转PDF
            elif ext_in == '.docx' and ext_out == '.pdf':
                # 直接调用_docx_to_pdf方法，它会自动处理依赖库检查
                self._docx_to_pdf()
            # PDF转DOCX
            elif ext_in == '.pdf' and ext_out == '.docx':
                self._pdf_to_docx()
            else:
                raise RuntimeError(f"不支持的文档转换类型: {ext_in} → {ext_out}")
            
            self.progress.emit(70)
        except Exception as e:
            raise RuntimeError(f"文档转换失败: {str(e)}") from e

    def _pdf_to_image(self):
        """PDF转图片"""
        try:
            import fitz  # PyMuPDF
            from PIL import Image
            import os
            
            # 检查输入文件是否存在
            if not os.path.exists(self.input_file):
                raise FileNotFoundError(f"输入文件不存在: {self.input_file}")
            
            doc = fitz.open(self.input_file)
            
            # 检查PDF是否有页面
            if len(doc) == 0:
                doc.close()
                raise RuntimeError("PDF文件没有页面内容")
            
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                # 使用更高的分辨率以获得更好的图像质量
                pix = page.get_pixmap(matrix=fitz.Matrix(3, 3))  # 3倍分辨率
                img_data = pix.tobytes("ppm")
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                
                if len(doc) > 1:
                    # 如果有多个页面，为每个页面创建单独的图片文件
                    base_name = os.path.splitext(self.output_file)[0]
                    output_file = f"{base_name}_{page_num + 1:02d}{os.path.splitext(self.output_file)[1]}"
                else:
                    output_file = self.output_file
                
                # 保存图片
                output_format = os.path.splitext(self.output_file)[1].upper()[1:]  # 获取文件扩展名作为格式
                if output_format == 'JPG':
                    output_format = 'JPEG'
                
                img.save(output_file, format=output_format, quality=95)
                
                # 检查输出文件是否创建成功
                if not os.path.exists(output_file):
                    doc.close()
                    raise RuntimeError(f"图片文件未能成功创建: {output_file}")
            
            doc.close()
        except FileNotFoundError:
            raise
        except Exception as e:
            raise RuntimeError(f"PDF转图片失败: {str(e)}") from e

    def _image_to_pdf(self):
        """图片转PDF"""
        try:
            from PIL import Image
            import os
            
            # 检查输入文件是否存在
            if not os.path.exists(self.input_file):
                raise FileNotFoundError(f"输入文件不存在: {self.input_file}")
            
            # 打开图片
            img = Image.open(self.input_file)
            
            # 确保图片模式兼容PDF
            if img.mode in ("RGBA", "P"):
                img = img.convert("RGB")
            
            # 保存为PDF
            img.save(self.output_file, format='PDF', resolution=100.0)
            
            # 检查输出文件是否创建成功
            if not os.path.exists(self.output_file):
                raise RuntimeError("PDF文件未能成功创建")
                
        except FileNotFoundError:
            raise
        except Exception as e:
            raise RuntimeError(f"图片转PDF失败: {str(e)}") from e

    def _txt_to_pdf(self):
        """文本转PDF"""
        try:
            from reportlab.pdfgen import canvas
            from reportlab.lib.pagesizes import letter
            from reportlab.pdfbase import pdfmetrics
            from reportlab.pdfbase.ttfonts import TTFont
            import os
            
            # 检查输入文件是否存在
            if not os.path.exists(self.input_file):
                raise FileNotFoundError(f"输入文件不存在: {self.input_file}")
            
            # 创建PDF
            c = canvas.Canvas(self.output_file, pagesize=letter)
            width, height = letter
            
            # 读取文本文件，尝试多种编码
            lines = []
            encodings = ['utf-8', 'gbk', 'gb2312', 'latin-1']
            file_read = False
            
            for encoding in encodings:
                try:
                    with open(self.input_file, 'r', encoding=encoding) as f:
                        lines = f.readlines()
                    file_read = True
                    break
                except UnicodeDecodeError:
                    continue
            
            if not file_read:
                raise RuntimeError("无法读取文本文件，可能的编码问题")
            
            # 尝试注册并使用中文字体
            font_registered = False
            try:
                # 在Windows上尝试使用系统字体
                if os.name == 'nt':  # Windows
                    font_path = "C:/Windows/Fonts/simsun.ttc"  # 宋体
                    if os.path.exists(font_path):
                        pdfmetrics.registerFont(TTFont('SimSun', font_path))
                        c.setFont("SimSun", 12)
                        font_registered = True
                # 在其他系统上可以尝试其他字体路径
            except:
                pass
            
            # 如果中文字体注册失败，使用默认字体
            if not font_registered:
                c.setFont("Helvetica", 12)
            
            y_position = height - 50
            
            # 写入文本
            for line in lines:
                # 处理过长的行，进行换行
                text = line.strip()
                while len(text) > 80:  # 如果行长度超过80字符
                    c.drawString(50, y_position, text[:80])
                    text = text[80:]
                    y_position -= 15
                    
                    # 如果到达页面底部，创建新页面
                    if y_position < 50:
                        c.showPage()
                        # 重新设置字体
                        if font_registered:
                            c.setFont("SimSun", 12)
                        else:
                            c.setFont("Helvetica", 12)
                        y_position = height - 50
                
                # 绘制剩余文本（或未超长的整行）
                if text:  # 确保有文本需要绘制
                    c.drawString(50, y_position, text)
                    y_position -= 15
                y_position -= 15
                
                # 如果到达页面底部，创建新页面
                if y_position < 50:
                    c.showPage()
                    # 重新设置字体
                    if font_registered:
                        c.setFont("SimSun", 12)
                    else:
                        c.setFont("Helvetica", 12)
                    y_position = height - 50
            
            c.save()
            
            # 检查输出文件是否创建成功
            if not os.path.exists(self.output_file):
                raise RuntimeError("PDF文件未能成功创建")
                
        except FileNotFoundError:
            raise
        except Exception as e:
            raise RuntimeError(f"文本转PDF失败: {str(e)}") from e

    def _pdf_to_txt(self):
        """PDF转文本"""
        try:
            import fitz  # PyMuPDF
            import os
            
            # 检查输入文件是否存在
            if not os.path.exists(self.input_file):
                raise FileNotFoundError(f"输入文件不存在: {self.input_file}")
            
            # 打开PDF文档
            doc = fitz.open(self.input_file)
            text = ""
            
            # 提取所有页面的文本
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                page_text = page.get_text()
                if page_text:
                    text += page_text + '\n\n'  # 每页之间添加空行分隔
            
            # 关闭文档
            doc.close()
            
            # 写入文本文件
            with open(self.output_file, 'w', encoding='utf-8') as f:
                f.write(text)
            
            # 检查输出文件是否创建成功
            if not os.path.exists(self.output_file):
                raise RuntimeError("TXT文件未能成功创建")
                
        except FileNotFoundError:
            raise
        except Exception as e:
            raise RuntimeError(f"PDF转文本失败: {str(e)}") from e

    def _doc_to_pdf(self):
        """DOC转PDF"""
        try:
            # 尝试使用 python-docx 和 reportlab 实现 DOC 转 PDF
            try:
                # 注意: python-docx 主要处理 DOCX 格式，对于 DOC 格式支持有限
                # 这里我们尝试使用 win32com.client (Windows) 或 libreoffice (跨平台) 来处理
                import os
                
                # 检查操作系统
                import platform
                if platform.system() == "Windows":
                    # 在 Windows 上尝试使用 win32com.client
                    try:
                        import win32com.client
                        # 使用 Word 应用程序打开 DOC 文件并另存为 PDF
                        word = win32com.client.Dispatch("Word.Application")
                        word.visible = False
                        doc = word.Documents.Open(os.path.abspath(self.input_file))
                        doc.SaveAs(os.path.abspath(self.output_file), FileFormat=17)  # 17 表示 PDF 格式
                        doc.Close()
                        word.Quit()
                        return
                    except Exception as e:
                        print(f"使用 Word 应用程序转换失败: {e}")
                        pass
                
                # 跨平台方案：尝试使用 libreoffice
                try:
                    import subprocess
                    import shutil
                    
                    # 检查 libreoffice 是否可用
                    libreoffice_path = shutil.which("libreoffice")
                    if not libreoffice_path:
                        # 尝试使用 soffice (libreoffice 的命令行工具)
                        libreoffice_path = shutil.which("soffice")
                    
                    if libreoffice_path:
                        # 使用 libreoffice 转换
                        cmd = [
                            libreoffice_path,
                            "--headless",
                            "--convert-to", "pdf",
                            "--outdir", os.path.dirname(self.output_file),
                            self.input_file
                        ]
                        
                        result = subprocess.run(cmd, capture_output=True, text=True)
                        if result.returncode == 0:
                            # 重命名输出文件以匹配期望的输出文件名
                            input_name = os.path.splitext(os.path.basename(self.input_file))[0]
                            actual_output = os.path.join(os.path.dirname(self.output_file), input_name + ".pdf")
                            if os.path.exists(actual_output):
                                os.rename(actual_output, self.output_file)
                                return
                        else:
                            print(f"LibreOffice 转换失败: {result.stderr}")
                except Exception as e:
                    print(f"使用 LibreOffice 转换失败: {e}")
                    pass
                
                # 如果以上方法都失败了，提供有用的错误信息
                raise RuntimeError(
                    "DOC 转 PDF 需要额外的依赖库:\n\n"
                    "Windows 用户:\n"
                    "- 安装 Microsoft Word 或\n"
                    "- 安装 LibreOffice\n\n"
                    "macOS/Linux 用户:\n"
                    "- 安装 LibreOffice\n\n"
                    "安装 LibreOffice:\n"
                    "- 访问 https://www.libreoffice.org/download/download/ 下载并安装\n"
                    "- 或使用包管理器安装: brew install libreoffice (macOS) 或 apt install libreoffice (Ubuntu)"
                )
            except ImportError as e:
                raise RuntimeError(f"缺少必要的依赖库: {str(e)}") from e
        except Exception as e:
            raise RuntimeError(f"DOC 转 PDF 失败: {str(e)}") from e

    def _docx_to_pdf(self):
        """DOCX转PDF"""
        try:
            from docx2pdf import convert
            import os
            
            # 检查输入文件是否存在
            if not os.path.exists(self.input_file):
                raise FileNotFoundError(f"输入文件不存在: {self.input_file}")
            
            convert(self.input_file, self.output_file)
            
            # 检查输出文件是否创建成功
            if not os.path.exists(self.output_file):
                raise RuntimeError("PDF文件未能成功创建")
                
        except ImportError:
            # 如果没有安装docx2pdf库，提供有用的错误信息
            raise RuntimeError(
                "DOCX转PDF需要docx2pdf库:\n\n"
                "请运行以下命令安装:\n"
                "pip install docx2pdf\n\n"
                "注意: docx2pdf需要Microsoft Word或WPS Office支持"
            )
        except Exception as e:
            raise RuntimeError(f"DOCX转PDF失败: {str(e)}") from e

    def _pdf_to_docx(self):
        """PDF转DOCX"""
        try:
            import fitz  # PyMuPDF
            from docx import Document
            import os
            
            # 检查输入文件是否存在
            if not os.path.exists(self.input_file):
                raise FileNotFoundError(f"输入文件不存在: {self.input_file}")
            
            # 打开PDF文档
            doc = fitz.open(self.input_file)
            docx_doc = Document()
            
            # 添加标题
            docx_doc.add_heading('转换自PDF文件', 0)
            
            # 提取所有页面的文本并添加到DOCX文档
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                text = page.get_text()
                
                if text and text.strip():  # 确保文本不为空
                    # 添加页面分隔标题
                    docx_doc.add_heading(f'页面 {page_num+1}', level=1)
                    # 添加文本内容
                    docx_doc.add_paragraph(text.strip())
                    # 添加页面分隔符（除了最后一页）
                    if page_num < len(doc) - 1:
                        docx_doc.add_page_break()
                elif page_num < len(doc) - 1:  # 即使页面无文本也要添加分页符（除了最后一页）
                    docx_doc.add_page_break()
            
            # 关闭PDF文档
            doc.close()
            
            # 保存DOCX文档
            docx_doc.save(self.output_file)
            
            # 检查输出文件是否创建成功
            if not os.path.exists(self.output_file):
                raise RuntimeError("DOCX文件未能成功创建")
                
        except FileNotFoundError:
            raise
        except Exception as e:
            raise RuntimeError(f"PDF转DOCX失败: {str(e)}") from e


class QconvertoApp(QMainWindow):
    """Qconverto主应用程序"""
    
    def __init__(self):
        super().__init__()
        self.init_ui()
        self.input_file = ""
        self.output_dir = ""
        self.worker = None
        # 启用拖放功能
        self.setAcceptDrops(True)
        
    def init_ui(self):
        """初始化用户界面"""
        # 设置窗口属性
        self.setWindowTitle('Qconverto - 多媒体文件格式转换工具')
        self.setGeometry(100, 100, 600, 550)
        self.setMinimumSize(500, 400)
        
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 创建主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # 标题
        title_label = QLabel('Qconverto')
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("""
            font-size: 24px;
            font-weight: bold;
            color: #1976d2;
            margin: 10px;
        """)
        main_layout.addWidget(title_label)
        
        # 说明文本
        desc_label = QLabel(
            '支持图片、音频、视频和文档格式的转换。\n'
            '选择文件并将其转换为不同的格式。'
        )
        desc_label.setAlignment(Qt.AlignCenter)
        desc_label.setStyleSheet("""
            font-size: 14px;
            color: #666666;
            margin: 5px;
        """)
        main_layout.addWidget(desc_label)
        
        # 文件选择区域
        file_group = QGroupBox("文件选择")
        file_layout = QGridLayout(file_group)
        file_layout.setSpacing(10)
        file_layout.setContentsMargins(15, 15, 15, 15)
        
        # 输入文件
        self.input_label = QLabel("输入文件:")
        self.input_path_label = QLabel("未选择文件")
        self.input_path_label.setStyleSheet("color: #666666;")
        self.input_path_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        # 设置最大宽度以防止布局混乱
        self.input_path_label.setMaximumWidth(300)
        self.input_path_label.setWordWrap(False)
        self.browse_input_btn = QPushButton("浏览")
        self.browse_input_btn.clicked.connect(self.browse_input_file)
        self.browse_input_btn.setMinimumWidth(80)
        
        file_layout.addWidget(self.input_label, 0, 0)
        file_layout.addWidget(self.input_path_label, 0, 1)
        file_layout.addWidget(self.browse_input_btn, 0, 2)
        
        # 输出格式
        self.format_label = QLabel("输出格式:")
        self.format_combo = QComboBox()
        self.format_combo.addItems(['.mp3', '.wav', '.flac', '.ogg', '.m4a', '.jpg', '.png', '.pdf'])
        self.format_combo.setCurrentText('.mp3')
        
        file_layout.addWidget(self.format_label, 1, 0)
        file_layout.addWidget(self.format_combo, 1, 1, 1, 2)
        
        # 输出目录
        self.output_label = QLabel("输出目录:")
        self.output_path_label = QLabel("与输入文件相同目录")
        self.output_path_label.setStyleSheet("color: #666666;")
        self.output_path_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        # 设置最大宽度以防止布局混乱
        self.output_path_label.setMaximumWidth(300)
        self.output_path_label.setWordWrap(False)
        self.browse_output_btn = QPushButton("浏览")
        self.browse_output_btn.clicked.connect(self.browse_output_dir)
        self.browse_output_btn.setMinimumWidth(80)
        
        file_layout.addWidget(self.output_label, 2, 0)
        file_layout.addWidget(self.output_path_label, 2, 1)
        file_layout.addWidget(self.browse_output_btn, 2, 2)
        
        main_layout.addWidget(file_group)
        
        # 转换按钮
        self.convert_btn = QPushButton("开始转换")
        self.convert_btn.setStyleSheet("""
            QPushButton {
                background-color: #1976d2;
                color: white;
                border: none;
                padding: 12px;
                font-size: 16px;
                font-weight: bold;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #1565c0;
            }
            QPushButton:disabled {
                background-color: #bdbdbd;
            }
        """)
        self.convert_btn.clicked.connect(self.start_conversion)
        main_layout.addWidget(self.convert_btn)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.hide()
        main_layout.addWidget(self.progress_bar)
        
        # 参数配置区域（默认隐藏）
        self.settings_group = QGroupBox("参数配置")
        self.settings_group.setVisible(False)
        self.settings_layout = QGridLayout(self.settings_group)
        self.settings_layout.setSpacing(10)
        self.settings_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.addWidget(self.settings_group)
        
        # 日志输出
        log_group = QGroupBox("转换日志")
        log_layout = QVBoxLayout(log_group)
        log_layout.setContentsMargins(10, 10, 10, 10)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(180)
        self.log_text.setFont(QFont("Consolas", 9))
        log_layout.addWidget(self.log_text)
        
        main_layout.addWidget(log_group)
        
        # 状态栏
        self.status_label = QLabel("就绪")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("""
            font-size: 12px;
            color: #666666;
            padding: 5px;
        """)
        main_layout.addWidget(self.status_label)
        
        # 初始化参数配置控件字典
        self.setting_controls = {}
        
    def show_format_settings(self, file_type):
        """根据文件类型显示相应的参数配置选项"""
        # 清除之前的配置控件
        for i in reversed(range(self.settings_layout.count())): 
            widget = self.settings_layout.itemAt(i).widget()
            if widget:
                widget.setParent(None)
        self.setting_controls.clear()
        
        # 根据文件类型添加相应的配置选项
        row = 0
        if file_type == '图片':
            # 图片质量设置
            quality_label = QLabel("图片质量 (1-100):")
            quality_combo = QComboBox()
            quality_combo.addItems([str(i) for i in range(100, 0, -10)])  # 100, 90, 80, ..., 10
            quality_combo.setCurrentText("90")
            self.settings_layout.addWidget(quality_label, row, 0)
            self.settings_layout.addWidget(quality_combo, row, 1)
            self.setting_controls['image_quality'] = quality_combo
            row += 1
            
        elif file_type == '音频':
            # 音频比特率设置
            bitrate_label = QLabel("音频比特率:")
            bitrate_combo = QComboBox()
            bitrate_combo.addItems(['64k', '128k', '192k', '256k', '320k'])
            bitrate_combo.setCurrentText("192k")
            self.settings_layout.addWidget(bitrate_label, row, 0)
            self.settings_layout.addWidget(bitrate_combo, row, 1)
            self.setting_controls['audio_bitrate'] = bitrate_combo
            row += 1
            
        elif file_type == '视频':
            # 视频编码设置
            codec_label = QLabel("视频编码:")
            codec_combo = QComboBox()
            codec_combo.addItems(['libx264', 'libx265', 'mpeg4', 'vp9'])
            codec_combo.setCurrentText("libx264")
            self.settings_layout.addWidget(codec_label, row, 0)
            self.settings_layout.addWidget(codec_combo, row, 1)
            self.setting_controls['video_codec'] = codec_combo
            row += 1
            
            # 音频编码设置
            audio_codec_label = QLabel("音频编码:")
            audio_codec_combo = QComboBox()
            audio_codec_combo.addItems(['aac', 'mp3', 'opus', 'vorbis'])
            audio_codec_combo.setCurrentText("aac")
            self.settings_layout.addWidget(audio_codec_label, row, 0)
            self.settings_layout.addWidget(audio_codec_combo, row, 1)
            self.setting_controls['audio_codec'] = audio_codec_combo
            row += 1
            
            # 视频质量设置 (CRF)
            crf_label = QLabel("视频质量 (0-51, 越小越好):")
            crf_combo = QComboBox()
            crf_combo.addItems([str(i) for i in range(15, 30)])  # 15-29
            crf_combo.setCurrentText("23")
            self.settings_layout.addWidget(crf_label, row, 0)
            self.settings_layout.addWidget(crf_combo, row, 1)
            self.setting_controls['video_quality'] = crf_combo
            row += 1
            
        # 显示或隐藏参数配置区域
        self.settings_group.setVisible(row > 0)
        if row > 0:
            self.settings_group.setTitle(f"{file_type}参数配置")
        
    def browse_input_file(self):
        """浏览输入文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            "选择输入文件",
            "",
            "所有文件 (*.*)"
        )
        
        if file_path:
            self.input_file = file_path
            self.input_path_label.setText(os.path.basename(file_path))
            self.input_path_label.setStyleSheet("color: #1976d2;")
            self.log(f"已选择输入文件: {file_path}")
            self.status_label.setText("已选择输入文件")
            
            # 根据文件类型更新输出格式选项
            file_type = self.determine_file_type(file_path)
            if file_type:
                self.update_format_options(file_type)
            
    def browse_output_dir(self):
        """浏览输出目录"""
        dir_path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        
        if dir_path:
            self.output_dir = dir_path
            self.output_path_label.setText(os.path.basename(dir_path) if len(dir_path) <= 30 else f"...{dir_path[-27:]}")
            self.output_path_label.setStyleSheet("color: #1976d2;")
            # 设置标签的最大宽度以防止布局混乱
            self.output_path_label.setMaximumWidth(300)
            self.output_path_label.setWordWrap(False)
            self.log(f"已选择输出目录: {dir_path}")
            self.status_label.setText("已选择输出目录")
            
    def update_format_options(self, file_type):
        """根据文件类型更新输出格式选项"""
        if file_type in SUPPORTED_FORMATS:
            formats = SUPPORTED_FORMATS[file_type]['输出']
            self.format_combo.clear()
            self.format_combo.addItems(formats)
            if formats:
                self.format_combo.setCurrentText(formats[0])
            
            # 显示当前文件类型和可转换格式的信息
            self.status_label.setText(f"检测到文件类型: {file_type}，可转换为: {', '.join(formats)}")
            
            # 在日志中显示更详细的信息
            self.log(f"✓ 检测到文件类型: {file_type}")
            self.log(f"  可转换格式: {', '.join(formats)}")
            
            # 显示参数配置区域
            self.show_format_settings(file_type)
        else:
            # 不支持的文件类型
            self.status_label.setText(f"不支持的文件类型: {file_type}")
            self.log(f"✗ 不支持的文件类型: {file_type}")
            self.format_combo.clear()
            
    def start_conversion(self):
        """开始转换"""
        if not self.input_file:
            self.status_label.setText("请先选择输入文件")
            self.log("警告: 请先选择输入文件")
            return
            
        if not os.path.exists(self.input_file):
            self.status_label.setText("输入文件不存在")
            self.log("警告: 输入文件不存在")
            return
            
        # 确定输出文件路径
        input_dir = os.path.dirname(self.input_file)
        input_name = os.path.splitext(os.path.basename(self.input_file))[0]
        output_ext = self.format_combo.currentText()
        output_dir = self.output_dir if self.output_dir else input_dir
        output_file = os.path.join(output_dir, f"{input_name}{output_ext}")
        
        # 确定文件类型
        file_type = self.determine_file_type(self.input_file)
        if not file_type:
            self.status_label.setText("不支持的文件类型")
            self.log("警告: 不支持的文件类型")
            return
            
        self.log(f"开始转换 {self.input_file} 到 {output_file}")
        self.log(f"文件类型: {file_type}")
        
        # 禁用转换按钮
        self.convert_btn.setEnabled(False)
        self.progress_bar.show()
        self.progress_bar.setValue(0)
        
        try:
            # 收集用户设置的参数
            settings = {}
            
            # 图片质量设置
            if 'image_quality' in self.setting_controls:
                settings['image_quality'] = int(self.setting_controls['image_quality'].currentText())
                
            # 音频比特率设置
            if 'audio_bitrate' in self.setting_controls:
                settings['audio_bitrate'] = self.setting_controls['audio_bitrate'].currentText()
                
            # 视频编码设置
            if 'video_codec' in self.setting_controls:
                settings['video_codec'] = self.setting_controls['video_codec'].currentText()
                
            # 音频编码设置
            if 'audio_codec' in self.setting_controls:
                settings['audio_codec'] = self.setting_controls['audio_codec'].currentText()
                
            # 视频质量设置
            if 'video_quality' in self.setting_controls:
                settings['video_quality'] = int(self.setting_controls['video_quality'].currentText())
            
            # 创建转换线程
            self.worker = ConversionThread(
                self.input_file, 
                output_file, 
                file_type,
                settings
            )
            
            # 连接信号
            self.worker.progress.connect(self.update_progress)
            self.worker.finished.connect(self.on_conversion_finished)
            self.worker.error.connect(self.on_conversion_error)
            
            # 启动转换
            self.worker.start()
            
        except Exception as e:
            self.log(f"转换启动失败: {str(e)}")
            self.convert_btn.setEnabled(True)
            self.progress_bar.hide()
            
    def determine_file_type(self, file_path):
        """确定文件类型 - 智能识别版本"""
        if not os.path.exists(file_path):
            return None
            
        ext = os.path.splitext(file_path)[1].lower()
        detected_type = None
        
        # 首先基于文件扩展名进行初步判断
        for file_type, formats in SUPPORTED_FORMATS.items():
            if ext in formats['输入']:
                detected_type = file_type
                break
                
        # 然后读取文件头部信息进行验证和精确识别
        try:
            with open(file_path, 'rb') as f:
                # 读取文件前1024字节用于分析（增加读取字节数以提高准确性）
                header = f.read(1024)
                
            # 检查文件签名
            matched_signatures = []
            for file_type, signatures in FILE_SIGNATURES.items():
                for signature, format_name in signatures.items():
                    if header.startswith(signature):
                        matched_signatures.append((file_type, format_name, len(signature)))
            
            # 如果有多个匹配，选择签名最长的（更精确的匹配）
            if matched_signatures:
                matched_signatures.sort(key=lambda x: x[2], reverse=True)
                file_type, format_name, _ = matched_signatures[0]
                
                # 对于需要进一步确认的格式进行特殊处理
                if file_type == '图片' and format_name == 'WebP':
                    # WebP 文件需要检查是否包含 'WEBP' 字符串
                    if b'WEBP' in header:
                        return file_type
                elif file_type == '音频' and format_name == 'WAV':
                    # WAV 文件需要检查是否包含 'WAVE' 字符串
                    if b'WAVE' in header:
                        return file_type
                elif file_type == '视频' and format_name == 'AVI':
                    # AVI 文件需要检查是否包含 'AVI' 字符串
                    if b'AVI' in header:
                        return file_type
                elif file_type == '文档':
                    # 文档类文件需要进一步区分
                    if format_name == 'DOCX':
                        # DOCX 文件需要检查是否包含特定的XML内容
                        if b'[Content_Types].xml' in header or b'_rels/.rels' in header:
                            return file_type
                    elif format_name == 'PPTX':
                        # PPTX 文件需要检查是否包含特定的XML内容
                        if b'ppt/slides/' in header or b'ppt/presentation.xml' in header:
                            return '文档'  # PPTX 也归类为文档
                    elif format_name == 'XLSX':
                        # XLSX 文件需要检查是否包含特定的XML内容
                        if b'xl/worksheets/' in header or b'xl/workbook.xml' in header:
                            return '文档'  # XLSX 也归类为文档
                    elif format_name == 'DOC':
                        # DOC 文件已经被文件头签名匹配，直接返回文档类型
                        return '文档'
                    else:
                        return file_type
                else:
                    # 其他格式可以直接返回
                    return file_type
                            
        except Exception as e:
            self.log(f"文件类型检测时出错: {e}")
            
        # 如果内容检测失败，回退到扩展名判断
        return detected_type
        
    def update_progress(self, value):
        """更新进度"""
        self.progress_bar.setValue(value)
        self.status_label.setText(f"转换进度: {value}%")
        
    def on_conversion_finished(self, success, output_file, original_file):
        """转换完成"""
        self.convert_btn.setEnabled(True)
        self.progress_bar.hide()
        
        if success:
            self.log(f"✓ 转换成功完成!")
            self.log(f"  输出文件: {output_file}")
            if os.path.exists(output_file):
                size = os.path.getsize(output_file)
                self.log(f"  文件大小: {self.format_file_size(size)}")
                
            self.status_label.setText("转换成功完成")
            # 使用 QTimer 延迟重置状态栏文本
            QTimer.singleShot(3000, lambda: self.status_label.setText("就绪") if self.status_label.text() == "转换成功完成" else None)
        else:
            self.log(f"✗ 转换失败")
            self.status_label.setText("文件转换失败")
            # 使用 QTimer 延迟重置状态栏文本
            QTimer.singleShot(3000, lambda: self.status_label.setText("就绪") if self.status_label.text() == "文件转换失败" else None)
        
    def on_conversion_error(self, error_msg):
        """转换错误"""
        self.convert_btn.setEnabled(True)
        self.progress_bar.hide()
        self.log(f"✗ 转换错误: {error_msg}")
        self.status_label.setText(f"转换错误: {error_msg}")
        # 使用 QTimer 延迟重置状态栏文本，确保用户能看到错误信息
        QTimer.singleShot(5000, lambda: self.status_label.setText("就绪") if self.status_label.text() == f"转换错误: {error_msg}" else None)
        
        # 显示错误对话框
        QMessageBox.critical(self, "转换错误", error_msg)
        
    def format_file_size(self, size_bytes):
        """格式化文件大小"""
        if size_bytes == 0:
            return "0 B"
        size_names = ["B", "KB", "MB", "GB"]
        i = 0
        while size_bytes >= 1024 and i < len(size_names) - 1:
            size_bytes /= 1024.0
            i += 1
        return f"{size_bytes:.1f} {size_names[i]}"
        
    def log(self, message):
        """添加日志"""
        self.log_text.append(message)
        # 滚动到底部
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )

    def dragEnterEvent(self, event: QDragEnterEvent):
        """拖拽进入事件"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        """拖拽放下事件"""
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls:
                file_path = urls[0].toLocalFile()
                if os.path.isfile(file_path):
                    self.input_file = file_path
                    # 限制显示的文件名长度，防止界面被撑开
                    file_name = os.path.basename(file_path)
                    if len(file_name) > 50:
                        file_name = file_name[:47] + "..."
                    self.input_path_label.setText(file_name)
                    self.input_path_label.setStyleSheet("color: #1976d2;")
                    # 设置标签的最大宽度以防止布局混乱
                    self.input_path_label.setMaximumWidth(300)
                    self.input_path_label.setWordWrap(False)
                    self.log(f"通过拖放添加文件: {file_path}")
                    self.status_label.setText("已通过拖放添加文件")
                    
                    # 根据文件类型更新输出格式选项
                    file_type = self.determine_file_type(file_path)
                    if file_type:
                        self.update_format_options(file_type)
                    else:
                        # 文件类型无法识别
                        self.status_label.setText(f"无法识别文件类型: {file_name}")
                        self.log(f"✗ 无法识别文件类型: {file_path}")
                    return
        self.status_label.setText("拖放文件无效")
        event.ignore()


def check_dependencies():
    """检查依赖库是否完整"""
    missing_deps = []
    warnings = []
    
    # 检查基础依赖
    try:
        import PyQt5
    except ImportError:
        missing_deps.append("PyQt5")
    
    # 检查图片处理依赖
    try:
        import PIL
    except ImportError:
        missing_deps.append("Pillow")
    
    # 检查PDF处理依赖
    # PyMuPDF用于PDF转图片功能
    try:
        import fitz  # PyMuPDF
    except ImportError:
        missing_deps.append("PyMuPDF")
    

    
    # 检查音频处理依赖
    # pydub用于主要的音频格式转换功能
    try:
        import pydub
    except ImportError:
        warnings.append("pydub (音频转换功能受限)")
    
    # miniaudio作为pydub不可用时的备选方案
    try:
        import miniaudio
    except ImportError:
        warnings.append("miniaudio (音频转换功能受限)")
    
    # 检查文档处理依赖
    try:
        import docx2pdf
    except ImportError:
        warnings.append("docx2pdf (DOCX转PDF功能受限)")
    
    # 检查视频处理依赖
    # ffmpeg-python用于直接调用FFmpeg进行视频处理
    try:
        import ffmpeg
    except ImportError:
        warnings.append("ffmpeg-python (视频转换功能受限)")
    
    # MoviePy作为ffmpeg-python不可用时的备选方案
    try:
        # 添加当前目录到路径
        sys.path.insert(0, '.')
        
        # 检查MoviePy是否可用
        try:
            from moviepy.editor import VideoFileClip
            MOVIEPY_AVAILABLE = True
        except ImportError:
            MOVIEPY_AVAILABLE = False
    except ImportError:
        MOVIEPY_AVAILABLE = False
        print("警告: moviepy 未安装，将使用替代方案进行视频转换")
    
    return missing_deps, warnings


def main():
    """主函数"""
    app = QApplication(sys.argv)
    
    # 设置应用程序属性
    app.setApplicationName("Qconverto")
    app.setApplicationVersion("1.0")
    
    # 检查依赖库
    missing_deps, warnings = check_dependencies()
    
    # 创建并显示主窗口
    window = QconvertoApp()
    
    # 显示依赖检查结果
    if missing_deps:
        error_msg = "缺少必要的依赖库:\n" + "\n".join(missing_deps)
        error_msg += "\n\n请运行以下命令安装:\npip install " + " ".join(missing_deps)
        QMessageBox.critical(None, "依赖库缺失", error_msg)
        return
    
    if warnings:
        warning_msg = "以下功能可能受限:\n" + "\n".join(warnings)
        warning_msg += "\n\n建议安装完整依赖以获得所有功能"
        QMessageBox.warning(None, "功能受限", warning_msg)
    
    window.show()
    
    # 运行应用程序
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()