#!/bin/bash

# 切换到脚本所在目录
cd "$(dirname "$0")"

echo "==================================================="
echo " Excel Converter for Declaration List - Setup Tool"
echo "==================================================="
echo ""

# 检查git是否已安装
if command -v git &> /dev/null; then
    echo "检查代码更新..."
    if [ -d ".git" ]; then
        # 如果是git仓库，拉取最新代码
        echo "从远程仓库拉取最新代码..."
        git pull
    else
        echo "当前目录不是git仓库，跳过更新检查"
    fi
else
    echo "Git未安装，跳过代码更新检查"
fi
echo ""

# 检查Python是否已安装
if ! command -v python3 &> /dev/null; then
    echo "Python 3未安装！请先安装Python 3.8或更高版本"
    echo "Mac用户可使用Homebrew: brew install python3"
    echo "Linux用户可使用包管理器: sudo apt install python3 python3-pip (Ubuntu/Debian)"
    echo ""
    exit 1
fi

# 显示Python版本
echo "检测到Python版本:"
python3 --version
echo ""

# 检查虚拟环境是否存在
if [ ! -d "venv" ]; then
    echo "创建虚拟环境..."
    python3 -m venv venv
else
    echo "虚拟环境已存在..."
fi

# 激活虚拟环境
echo "激活虚拟环境..."
source venv/bin/activate

# 升级pip
echo "升级pip..."
pip install --upgrade pip

# 检查requirements.txt是否存在
if [ ! -f "requirements.txt" ]; then
    echo "创建requirements.txt文件..."
    echo "streamlit>=1.22.0" > requirements.txt
    echo "pandas>=1.5.0" >> requirements.txt
    echo "openpyxl>=3.1.0" >> requirements.txt
    echo "numpy>=1.22.0" >> requirements.txt
fi

# 安装依赖
echo "安装依赖项..."
pip install -r requirements.txt

echo ""
echo "==================================================="
echo " 安装完成！启动应用程序..."
echo "==================================================="
echo ""

# 检查8502端口是否被占用
PORT=8502
echo "检查端口 $PORT 是否被占用..."
if lsof -i :$PORT > /dev/null; then
    echo "端口 $PORT 已被占用，尝试释放..."
    PID=$(lsof -ti :$PORT)
    echo "杀死进程 PID: $PID"
    kill -9 $PID
    sleep 1
    echo "端口已释放"
fi

# 启动Streamlit应用，指定端口8502
streamlit run "app.py" --server.port=$PORT

# 使脚本可执行
# chmod +x setup_mac.sh