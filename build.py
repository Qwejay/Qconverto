import os
import sys
import subprocess
import shutil
from pathlib import Path

# 项目根目录
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))

# 检查Python版本
if sys.version_info < (3, 7):
    print("错误: 需要Python 3.7或更高版本")
    sys.exit(1)

def run_command(command, cwd=None):
    """运行命令并返回执行结果"""
    print(f"执行命令: {' '.join(command)}")
    try:
        result = subprocess.run(
            command,
            cwd=cwd or PROJECT_ROOT,
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=300  # 设置5分钟超时
        )
        if result.stdout:
            print(f"命令执行成功: {result.stdout}")
        else:
            print("命令执行成功")
        return True
    except subprocess.CalledProcessError as e:
        print(f"命令执行失败: {e.stderr}")
        return False
    except subprocess.TimeoutExpired:
        print("命令执行超时")
        return False

def install_dependencies():
    """安装项目依赖"""
    print("安装项目依赖...")
    # 首先尝试安装不依赖编译的包
    # 视频处理依赖：ffmpeg-python用于直接调用FFmpeg进行视频处理，MoviePy作为备选方案
    # 图片处理依赖：Pillow用于主要的图片处理功能
    # PDF处理依赖：PyMuPDF用于PDF处理功能
    basic_packages = [
        'PyQt5==5.15.9',
        'Pillow==10.2.0',
        'pydub==0.25.1',
        'ffmpeg-python==0.2.0',
        'docx2pdf==0.1.8',
        'PyMuPDF==1.23.6',
        'moviepy==1.0.3'
    ]
    
    # 安装基本包
    for package in basic_packages:
        print(f"安装包: {package}")
        if not run_command([sys.executable, '-m', 'pip', 'install', package]):
            print(f"警告: 包 {package} 安装失败，将继续安装其他包")
    
    # 尝试安装miniaudio，如果失败则给出提示
    # miniaudio作为pydub不可用时的备选方案
    print("安装miniaudio包...")
    if not run_command([sys.executable, '-m', 'pip', 'install', 'miniaudio==1.56']):
        print("警告: miniaudio包安装失败")
        print("提示: 如果需要音频转换功能，请安装Microsoft Visual C++ 14.0或更高版本")
        print("可以从 https://visualstudio.microsoft.com/visual-cpp-build-tools/ 下载")
    
    return True

def convert_svg_to_ico():
    """将SVG图标转换为ICO格式"""
    print("将SVG图标转换为ICO格式...")
    try:
        from PIL import Image
        # 打开SVG文件
        svg_path = os.path.join(PROJECT_ROOT, 'icon.svg')
        ico_path = os.path.join(PROJECT_ROOT, 'icon.ico')
        
        # 检查SVG文件是否存在
        if not os.path.exists(svg_path):
            print("警告: 未找到icon.svg文件")
            return False
            
        # 尝试使用PIL打开SVG并保存为ICO
        # 注意：PIL对SVG的支持有限，可能需要安装额外的库
        img = Image.open(svg_path)
        # 调整大小以适应ICO格式要求
        img = img.resize((256, 256), Image.Resampling.LANCZOS)
        img.save(ico_path, format='ICO')
        print(f"图标已转换: {ico_path}")
        return True
    except ImportError:
        print("警告: PIL库未正确安装，无法转换图标")
        print("提示: 可以手动将icon.svg转换为icon.ico")
        return False
    except Exception as e:
        print(f"图标转换失败: {str(e)}")
        print("提示: 可以手动将icon.svg转换为icon.ico，或者在打包时移除--icon参数")
        return False

def build_executable():
    """使用PyInstaller打包可执行文件"""
    print("开始打包可执行文件...")
    
    # 确保icon.ico存在
    icon_path = os.path.join(PROJECT_ROOT, 'icon.ico')
    if not os.path.exists(icon_path):
        print("警告: 未找到icon.ico文件，将使用默认图标")
        icon_arg = []
    else:
        icon_arg = ['--icon', 'icon.ico']
    
    # 构建PyInstaller命令
    cmd = [
        'pyinstaller',
        '--onefile',
        '--windowed',
        '--name', 'Qconverto',
        '--noconfirm',
        '--optimize=2',  # 优化级别2，进一步减小文件大小
        '--strip',  # 移除符号表和调试信息
        # 排除一些可能不需要的大型依赖
        '--exclude-module', 'matplotlib',
        '--exclude-module', 'numpy',
        '--exclude-module', 'scipy',
        '--exclude-module', 'pandas',
        # 减少收集的文件
        '--collect-submodules=PIL',
        '--collect-submodules=pydub',
        '--collect-submodules=PyQt5.QtCore',
        '--collect-submodules=PyQt5.QtGui',
        '--collect-submodules=PyQt5.QtWidgets',
        *icon_arg,
        'main.py'
    ]
    
    return run_command(cmd)

def cleanup():
    """清理构建过程中生成的临时文件"""
    print("清理临时文件...")
    
    # 需要清理的目录和文件
    cleanup_items = [
        'dist/Qconverto.spec',
        'Qconverto.spec',
    ]
    
    for item in cleanup_items:
        item_path = os.path.join(PROJECT_ROOT, item)
        if os.path.isdir(item_path):
            try:
                shutil.rmtree(item_path)
                print(f"已删除目录: {item_path}")
            except Exception as e:
                print(f"删除目录失败 {item_path}: {str(e)}")
        elif os.path.isfile(item_path):
            try:
                os.remove(item_path)
                print(f"已删除文件: {item_path}")
            except Exception as e:
                print(f"删除文件失败 {item_path}: {str(e)}")

def main():
    """主函数"""
    print("=== Qconverto 构建脚本 ===")
    
    # 步骤1: 安装依赖
    if not install_dependencies():
        print("依赖安装失败，退出构建")
        sys.exit(1)
    
    # 步骤2: 转换图标
    # 即使图标转换失败，也继续执行打包过程
    convert_svg_to_ico()
    
    # 步骤3: 打包可执行文件
    if not build_executable():
        print("打包失败，退出构建")
        sys.exit(1)
    
    # 步骤4: 清理临时文件
    cleanup()
    
    # 检查输出文件
    output_file = os.path.join(PROJECT_ROOT, 'dist', 'Qconverto.exe')
    if os.path.exists(output_file):
        print(f"\n构建成功！")
        print(f"可执行文件位置: {output_file}")
        print("\n使用说明:")
        print("1. 双击Qconverto.exe即可运行")
        print("2. 无需安装任何依赖")
        print("3. 支持拖放文件或点击选择文件进行转换")
    else:
        print("构建过程完成，但未找到输出文件，请检查错误信息")

if __name__ == '__main__':
    main()