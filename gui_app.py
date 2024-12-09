#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from script_generator import ScriptGenerator
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("拍摄脚本处理程序")
        self.root.geometry("500x300")
        
        # 设置窗口样式
        style = ttk.Style()
        style.configure("TButton", padding=6)
        style.configure("TLabel", padding=6)
        
        # 创建主框架
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(expand=True, fill=tk.BOTH)
        
        # 创建标题
        title_label = ttk.Label(main_frame, text="拍摄脚本处理程序", font=("STHeiti", 24))
        title_label.pack(pady=20)
        
        # 创建上传按钮
        self.upload_button = ttk.Button(
            main_frame,
            text="选择PPTX文件",
            command=self.select_file,
            style="Upload.TButton"
        )
        self.upload_button.pack(pady=20)
        
        # 创建文件名标签
        self.file_label = ttk.Label(main_frame, text="未选择文件")
        self.file_label.pack(pady=10)
        
        # 创建进度条
        self.progress = ttk.Progressbar(main_frame, length=300, mode='determinate')
        self.progress.pack(pady=10)
        
        # 创建状态标签
        self.status_label = ttk.Label(main_frame, text="就绪")
        self.status_label.pack(pady=10)
        
        # 存储选择的文件
        self.selected_file = None
        
        # 自定义按钮样式
        style.configure(
            "Upload.TButton",
            font=("STHeiti", 14),
            padding=10
        )

    def select_file(self):
        """选择PPTX文件"""
        file = filedialog.askopenfilename(
            title="选择PPTX文件",
            filetypes=[("PowerPoint文件", "*.pptx")]
        )
        if file:
            self.selected_file = file
            self.file_label['text'] = os.path.basename(file)
            self.status_label['text'] = "就绪"
            self.progress['value'] = 0
            self.process_file()

    def process_file(self):
        """处理选中的文件"""
        if not self.selected_file:
            return
            
        try:
            # 注册字体
            font_path = "/System/Library/Fonts/STHeiti Light.ttc"
            pdfmetrics.registerFont(TTFont("STHeiti", font_path))
            
            # 创建生成器实例
            generator = ScriptGenerator()
            
            # 更新状态
            self.status_label['text'] = f"正在处理: {os.path.basename(self.selected_file)}"
            self.progress['value'] = 20
            self.root.update()
            
            # 处理文件
            generator.process_file(self.selected_file)
            
            # 更新进度
            self.progress['value'] = 100
            self.status_label['text'] = "处理完成！"
            self.root.update()
            
            # 显示完成消息
            output_file = os.path.splitext(self.selected_file)[0] + "_拍摄需求.pdf"
            messagebox.showinfo(
                "完成",
                f"文件处理完成！\n\n生成的PDF文件：\n{os.path.basename(output_file)}"
            )
            
            # 重置状态
            self.selected_file = None
            self.file_label['text'] = "未选择文件"
            
        except Exception as e:
            messagebox.showerror("错误", f"处理文件时发生错误：\n{str(e)}")
            self.status_label['text'] = "处理出错"
            self.progress['value'] = 0

def main():
    root = tk.Tk()
    app = App(root)
    root.mainloop()

if __name__ == "__main__":
    main() 