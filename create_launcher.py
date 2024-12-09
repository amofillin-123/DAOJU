#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

class LauncherCreator:
    def __init__(self, root):
        self.root = root
        self.root.title("快捷启动工具创建器")
        self.root.geometry("500x300")
        
        # 设置窗口样式
        style = ttk.Style()
        style.configure("TButton", padding=6)
        style.configure("TLabel", padding=6)
        
        # 创建主框架
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 创建标题
        title_label = ttk.Label(main_frame, text="快捷启动工具创建器", font=("STHeiti", 20))
        title_label.grid(row=0, column=0, columnspan=2, pady=20)
        
        # 创建目标目录选择框
        self.dir_frame = ttk.Frame(main_frame)
        self.dir_frame.grid(row=1, column=0, columnspan=2, pady=10)
        
        self.dir_label = ttk.Label(self.dir_frame, text="目标目录：")
        self.dir_label.grid(row=0, column=0)
        
        self.dir_entry = ttk.Entry(self.dir_frame, width=40)
        self.dir_entry.grid(row=0, column=1, padx=5)
        
        self.dir_button = ttk.Button(self.dir_frame, text="选择", command=self.select_directory)
        self.dir_button.grid(row=0, column=2)
        
        # 创建按钮
        self.create_button = ttk.Button(main_frame, text="创建启动工具", command=self.create_launcher)
        self.create_button.grid(row=2, column=0, columnspan=2, pady=20)
        
        # 创建状态标签
        self.status_label = ttk.Label(main_frame, text="就绪")
        self.status_label.grid(row=3, column=0, columnspan=2)

    def select_directory(self):
        """选择目标目录"""
        directory = filedialog.askdirectory(title="选择目标目录")
        if directory:
            self.dir_entry.delete(0, tk.END)
            self.dir_entry.insert(0, directory)

    def create_launcher(self):
        """创建启动工具"""
        target_dir = self.dir_entry.get().strip()
        if not target_dir:
            messagebox.showwarning("警告", "请先选择目标目录！")
            return
            
        try:
            # 复制必要文件
            source_files = [
                "script_generator.py",
                "gui_app.py",
                "requirements.txt",
                "README.md"
            ]
            
            for file in source_files:
                if os.path.exists(file):
                    shutil.copy2(file, target_dir)
                    self.status_label['text'] = f"复制文件: {file}"
                    self.root.update()
            
            # 创建启动脚本
            launcher_path = os.path.join(target_dir, "start_app.command")
            with open(launcher_path, "w") as f:
                f.write("""#!/bin/bash
cd "$(dirname "$0")"
if ! command -v python3 &> /dev/null; then
    echo "错误：未安装Python 3"
    exit 1
fi

if [ ! -f "requirements.txt" ]; then
    echo "错误：找不到requirements.txt文件"
    exit 1
fi

# 检查是否需要安装依赖
if ! python3 -c "import pptx" 2>/dev/null; then
    echo "正在安装依赖包..."
    pip3 install -r requirements.txt
fi

# 启动GUI应用
python3 gui_app.py
""")
            
            # 设置执行权限
            os.chmod(launcher_path, 0o755)
            
            self.status_label['text'] = "创建完成！"
            messagebox.showinfo("完成", f"启动工具已创建在：\n{launcher_path}\n\n双击该文件即可运行程序。")
            
        except Exception as e:
            messagebox.showerror("错误", f"创建启动工具时发生错误：\n{str(e)}")
            self.status_label['text'] = "创建失败"

def main():
    root = tk.Tk()
    app = LauncherCreator(root)
    root.mainloop()

if __name__ == "__main__":
    main() 