# 快速开始指南

## 环境要求
- Python 3.7 或更高版本
- Windows 操作系统

## 安装步骤

1. 如果您还没有安装Python，请先从[Python官网](https://www.python.org/downloads/)下载并安装Python 3.7或更高版本
2. 下载本程序的所有文件到本地文件夹
3. 双击运行`install.py`，它会自动安装所需的依赖包

## 使用方法

### 方式一：命令行操作（适合批量处理）

1. 打开命令提示符（CMD）
2. 进入程序所在目录
3. 运行以下命令：
   ```
   python merge_excel.py <输入文件路径> <输出文件路径>
   ```
   例如：
   ```
   python merge_excel.py input.xlsx output.xlsx
   ```

### 方式二：Web界面操作（适合单次处理）

1. 打开命令提示符（CMD）
2. 进入程序所在目录
3. 运行以下命令：
   ```
   python streamlit_app.py
   ```
4. 等待浏览器自动打开Web界面
5. 在Web界面上传文件并进行处理

## 常见问题

### 1. 安装依赖包失败
- 检查网络连接是否正常
- 尝试使用管理员权限运行`install.py`
- 如果使用公司网络，可能需要配置代理

### 2. 文件处理出错
- 确保输入文件格式正确（Excel文件）
- 检查文件是否被其他程序占用
- 查看程序目录下的`merge_excel.log`文件了解详细错误信息

### 3. Web界面无法打开
- 确保端口8501没有被其他程序占用
- 检查防火墙设置
- 尝试使用管理员权限运行

## 反馈与支持

如果您遇到任何问题：
1. 查看`merge_excel.log`文件了解详细错误信息
2. 将错误信息和问题描述发送给技术支持