#!/bin/bash
cd "$(dirname "$0")"

# 检查Python环境
if ! command -v python3 &> /dev/null; then
    echo "错误：未安装Python 3"
    exit 1
fi

# 检查依赖
if ! python3 -c "import pptx" 2>/dev/null; then
    echo "正在安装依赖包..."
    pip3 install -r requirements.txt
fi

# 启动GUI应用
python3 gui_app.py 