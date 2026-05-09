"""依赖检查和自动安装模块"""
import sys
import subprocess
from pathlib import Path

# Windows 下强制 UTF-8 输出
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")


def check_and_install_dependencies():
    """
    检查必需的依赖是否已安装，如果缺失则自动安装

    Returns:
        bool: True 表示所有依赖已就绪，False 表示安装失败
    """
    # 定义必需的依赖包
    # 格式: (import_name, package_name)
    required_packages = [
        ("anthropic", "anthropic>=0.39.0"),
        ("zhipuai", "zhipuai>=2.0.0"),
        ("openai", "openai>=1.0.0"),
        ("pptx", "python-pptx>=1.0.0"),
        ("edge_tts", "edge-tts>=6.1.0"),
        ("dotenv", "python-dotenv>=1.0.0"),
    ]

    # Windows 平台才需要 pywin32
    if sys.platform == "win32":
        required_packages.append(("win32com", "pywin32>=306"))

    missing_packages = []

    # 检查每个依赖
    for import_name, package_spec in required_packages:
        try:
            # 使用 importlib 避免实际导入模块
            import importlib.util
            spec = importlib.util.find_spec(import_name)
            if spec is None:
                missing_packages.append(package_spec)
        except (ImportError, ModuleNotFoundError):
            missing_packages.append(package_spec)

    # 如果没有缺失的包，直接返回
    if not missing_packages:
        return True

    # 发现缺失的包，询问用户是否自动安装
    print("=" * 60)
    print("检测到缺失的依赖包:")
    print("=" * 60)
    for pkg in missing_packages:
        print(f"  - {pkg}")
    print()

    # 自动安装（可以改为询问用户）
    response = input("是否自动安装缺失的依赖？(Y/n): ").strip().lower()

    if response in ('', 'y', 'yes'):
        print("\n开始安装依赖...")
        print("-" * 60)

        try:
            # 使用 pip 安装缺失的包
            cmd = [sys.executable, "-m", "pip", "install"] + missing_packages

            result = subprocess.run(
                cmd,
                check=True,
                capture_output=True,
                text=True
            )

            print("✓ 依赖安装完成")
            print("-" * 60)
            return True

        except subprocess.CalledProcessError as e:
            print(f"❌ 依赖安装失败: {e}")
            print(f"错误信息: {e.stderr}")
            print("\n请手动运行以下命令安装依赖:")
            print(f"  pip install {' '.join(missing_packages)}")
            return False
    else:
        print("\n已取消自动安装。")
        print("请手动运行以下命令安装依赖:")
        print(f"  pip install {' '.join(missing_packages)}")
        print("\n或者运行:")
        print("  pip install -r requirements.txt")
        return False


def check_optional_dependencies():
    """
    检查可选依赖（不会阻止程序运行）
    """
    optional_packages = [
        ("pdf2image", "pdf2image", "用于 LibreOffice PDF 转图片"),
    ]

    missing_optional = []

    for import_name, package_name, description in optional_packages:
        try:
            import importlib.util
            spec = importlib.util.find_spec(import_name)
            if spec is None:
                missing_optional.append((package_name, description))
        except (ImportError, ModuleNotFoundError):
            missing_optional.append((package_name, description))

    if missing_optional:
        print("\n提示: 以下可选依赖未安装（不影响基本功能）:")
        for pkg, desc in missing_optional:
            print(f"  - {pkg}: {desc}")
        print()


if __name__ == "__main__":
    # 测试依赖检查
    if check_and_install_dependencies():
        print("✓ 所有必需依赖已就绪")
        check_optional_dependencies()
    else:
        print("❌ 依赖检查失败")
        sys.exit(1)
