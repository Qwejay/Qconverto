#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Qconverto - UI组件模块
包含优化后的用户界面组件
"""

import os
from pathlib import Path
from nicegui import ui, events
from typing import Optional, Callable


class ModernUIComponents:
    """现代化UI组件类"""
    
    def __init__(self, app_instance):
        self.app = app_instance
        self.setup_modern_ui()
        
    def setup_modern_ui(self):
        """设置现代化用户界面"""
        # 清除原有界面
        self.app.main_container.clear()
        
        # 顶部导航栏
        with ui.header(elevated=True).classes('items-center justify-between p-4'):
            with ui.row().classes('items-center'):
                ui.icon('transform').classes('text-2xl')
                ui.label('Qconverto').classes('text-xl font-bold')
            ui.label('多功能文件格式转换工具').classes('text-sm opacity-80')
        
        # 主内容区域
        with ui.column().classes('w-full max-w-4xl mx-auto p-4 gap-6'):
            # 文件拖拽区域
            self.create_drag_drop_area()
            
            # 进度和状态区域
            self.create_progress_status_area()
            
            # 操作按钮组
            self.create_action_buttons()
    
    def create_drag_drop_area(self):
        """创建文件拖拽区域"""
        with ui.column().classes('w-full items-center'):
            # 新的三段式布局容器
            with ui.row().classes('w-full justify-between items-center gap-8') as conversion_layout:
                # 左侧：原始文件显示区域（整合拖拽功能）
                with ui.column().classes('flex-1 items-center p-6 border-2 border-dashed border-gray-300 rounded-lg min-h-64 cursor-pointer transition-all duration-300 file-area-hover') as left_area:
                    ui.label('原始文件').classes('font-medium mb-4')
                    self.app.original_file_icon = ui.icon('insert_drive_file').classes('text-6xl text-gray-400 mb-2')
                    self.app.original_file_name = ui.label('未选择文件').classes('text-gray-500 text-center')
                    
                    # 添加选择文件按钮到原始文件区域内
                    ui.button(
                        '选择文件', 
                        icon='folder_open',
                        on_click=lambda: self.app.uploader.run_method('pickFiles')
                    ).classes('px-6 py-3 bg-blue-500 text-white rounded-lg hover:bg-blue-600 transition-colors mt-4')
                    
                    # 添加拖拽上传组件到原始文件区域
                    self.app.uploader = ui.upload(
                        multiple=False,
                        auto_upload=True,
                        on_upload=self.app.handle_file_upload
                    ).props('accept=*/*').classes('hidden')
                    
                    # 添加拖拽事件处理
                    left_area.on('dragover.prevent', lambda e: left_area.classes(replace='flex-1 items-center p-6 border-2 border-dashed border-blue-500 rounded-lg min-h-64 cursor-pointer transition-all duration-300 bg-blue-50 file-area-hover'))
                    left_area.on('dragleave.prevent', lambda e: left_area.classes(replace='flex-1 items-center p-6 border-2 border-dashed border-gray-300 rounded-lg min-h-64 cursor-pointer transition-all duration-300 file-area-hover'))
                    left_area.on('drop.prevent', lambda e: (
                        self.app.uploader.run_method('pickFiles'),
                        left_area.classes(replace='flex-1 items-center p-6 border-2 border-dashed border-green-500 rounded-lg min-h-64 cursor-pointer transition-all duration-300 bg-green-50 file-area-hover')
                    ))
                
                # 中间：转换箭头动画区域
                with ui.column().classes('items-center justify-center') as center_area:
                    self.app.conversion_arrow = ui.icon('arrow_forward').classes('text-4xl text-blue-500')
                    self.app.conversion_spinner = ui.spinner(size='lg', thickness=5).classes('hidden')
                    
                    # 将开始转换按钮移到箭头下方
                    self.app.convert_button = ui.button(
                        '开始转换', 
                        icon='play_arrow',
                        on_click=self.app.start_conversion
                    ).classes('px-6 py-3 bg-green-500 text-white rounded-lg hover:bg-green-600 transition-colors mt-4').bind_visibility_from(self.app, 'input_file')
                
                # 右侧：转换后文件显示区域（添加格式选择功能）
                with ui.column().classes('flex-1 items-center p-6 border-2 border-dashed border-gray-300 rounded-lg min-h-64') as right_area:
                    ui.label('转换后文件').classes('font-medium mb-4')
                    self.app.converted_file_icon = ui.icon('insert_drive_file').classes('text-6xl text-gray-400 mb-2')
                    self.app.converted_file_name = ui.label('未转换').classes('text-gray-500 text-center')
                    
                    # 添加输出格式选择器
                    with ui.row().classes('w-full items-center mt-4'):
                        ui.label('输出格式:').classes('text-sm')
                        self.app.format_select = ui.select([], value=None).classes('flex-grow')
                    
                    # 添加输出格式选择器（用于向后兼容）
                    self.app.output_format_select = ui.select([], value=None).classes('hidden')
                    
                    # 添加下载按钮占位符
                    with ui.column().classes('items-center mt-4') as download_area:
                        self.app.download_button_container = ui.column().classes('w-full items-center')
    
    def create_progress_status_area(self):
        """创建进度和状态区域"""
        with ui.column().classes('w-full').bind_visibility_from(self.app, 'conversion_in_progress'):
            # 进度条
            with ui.card().classes('w-full'):
                ui.label('转换进度').classes('font-medium mb-2')
                self.app.progress_bar = ui.linear_progress(value=0, show_value=True).classes('w-full')
                self.app.progress_text = ui.label('准备就绪').classes('text-sm text-gray-600 mt-1')
            
            # 实时预览（适用于图片）
            self.app.preview_container = ui.column().classes('w-full mt-4')
    
    def create_action_buttons(self):
        """创建操作按钮组"""
        with ui.row().classes('w-full justify-center gap-4 mt-4'):
            # 取消按钮保留在这里
            self.app.cancel_button = ui.button(
                '取消', 
                icon='cancel',
                on_click=self.app.cancel_conversion
            ).classes('px-6 py-3 text-lg').bind_visibility_from(self.app, 'conversion_in_progress')
    
    def update_format_settings(self, file_type: str):
        """根据文件类型更新参数设置界面"""
        # 转换设置功能已移除
        pass