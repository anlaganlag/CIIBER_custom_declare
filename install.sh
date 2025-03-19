#!/bin/bash

# 检查是否安装了Python 3.7或更高版本
if command -v python3 >/dev/null 2>&1; then
    python3 install.py
else
    echo "❌ 请先安装Python 3.7或更高版本"
    exit 1
fi