import sys
import subprocess
import os

def print_banner():
    """打印启动横幅"""
    print("""
╔═══════════════════════════════════════╗
║     Python环境配置助手 v1.0           ║
║     作者: Omnitek.AI技术团队          ║
╚═══════════════════════════════════════╝
""")

def check_python_version():
    """检查Python版本是否满足要求"""
    required_version = (3, 7)
    current_version = sys.version_info
    
    print("\n[1/2] 检查Python版本...")
    if current_version < required_version:
        print(f"❌ Python版本要求3.7或更高，当前版本为{sys.version.split()[0]}")
        return False
    print(f"✓ Python版本检查通过：{sys.version.split()[0]}")
    print("\n提示：按任意键继续...")
    input()
    return True

def install_requirements():
    """安装所需的依赖包"""
    try:
        print("\n[2/2] 正在安装依赖包...")
        # 先尝试安装openpyxl的依赖
        subprocess.check_call([sys.executable, "-m", "pip", "install", "lxml"])
        # 安装其他依赖包
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        print("✓ 依赖包安装成功！")
        print("\n提示：按任意键继续...")
        input()
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ 依赖包安装失败：{str(e)}")
        return False

def main():
    print_banner()
    
    print("\n欢迎使用Python环境配置助手！")
    print("本程序将帮助您配置所需的Python环境。")
    print("\n提示：按任意键开始配置...")
    input()
    
    # 检查Python版本
    if not check_python_version():
        print("\n❗ 请确保您已安装Python 3.7或更高版本。")
        print("\n提示：按Enter键退出...")
        input()
        return
    
    # 检查requirements.txt是否存在
    if not os.path.exists("requirements.txt"):
        print("\n❌ 未找到requirements.txt文件")
        print("\n❗ 请确保requirements.txt文件在当前目录中。")
        print("\n提示：按Enter键退出...")
        input()
        return
    
    # 安装依赖
    if not install_requirements():
        print("\n❗ 安装依赖包失败，请检查网络连接或尝试手动安装。")
        print("\n提示：按Enter键退出...")
        input()
        return
    
    print("\n✨ 环境配置完成！现在您可以运行程序了。")
    print("\n运行方式：")
    print("1. 处理Excel文件：")
    print("   python merge_excel.py <输入文件> <输出文件>")
    print("2. 启动Web界面：")
    print("   python streamlit_app.py")
    print("\n如需帮助，请参考README-快速开始.md文件。")
    print("\n提示：按Enter键退出...")
    input()

if __name__ == "__main__":
    main()