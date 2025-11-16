#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Qconverto - Web服务入口点
用于在Render等云平台上部署
"""

import os
import sys
from pathlib import Path

# 添加当前目录到Python路径
sys.path.insert(0, str(Path(__file__).parent))

# 导入主应用
from main import *

if __name__ == "__main__":
    # 获取Render提供的端口，如果没有则使用默认端口
    port = int(os.environ.get("PORT", 8080))
    
    # 运行应用
    ui.run(
        title='Qconverto - 多媒体文件格式转换工具',
        reload=False,
        favicon='icon.svg',
        dark=False,
        port=port,
        host="0.0.0.0"  # 绑定到所有接口以在Render上正常工作
    )