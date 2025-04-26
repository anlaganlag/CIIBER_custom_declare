@echo off
echo ===================================================
echo  Excel Converter for Declaration List - Setup Tool
echo ===================================================
echo.

:: 检查Python是否已安装
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo Python未安装！请先安装Python 3.8或更高版本
    echo 可以从 https://www.python.org/downloads/ 下载
    echo.
    pause
    exit /b 1
)

:: 显示Python版本
echo 检测到Python版本:
python --version
echo.

:: 检查虚拟环境是否存在
if not exist venv (
    echo 创建虚拟环境...
    python -m venv venv
) else (
    echo 虚拟环境已存在...
)

:: 激活虚拟环境
echo 激活虚拟环境...
call venv\Scripts\activate.bat

:: 升级pip
echo 升级pip...
python -m pip install --upgrade pip

:: 检查requirements.txt是否存在
if not exist requirements.txt (
    echo 创建requirements.txt文件...
    echo streamlit>=1.22.0 > requirements.txt
    echo pandas>=1.5.0 >> requirements.txt
    echo openpyxl>=3.1.0 >> requirements.txt
    echo numpy>=1.22.0 >> requirements.txt
)

:: 安装依赖
echo 安装依赖项...
pip install -r requirements.txt

echo.
echo ===================================================
echo  安装完成！启动应用程序...
echo ===================================================
echo.

:: 启动Streamlit应用
streamlit run app.py

:: 保持窗口打开
pause 