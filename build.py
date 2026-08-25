#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PyInstaller 打包脚本

使用方法:
    python build.py

生成文件:
    dist/fcp.exe - 单文件可执行程序
"""

import subprocess
import sys
import os

def build():
    """执行 PyInstaller 打包"""
    
    # 检查 PyInstaller 是否安装
    try:
        import PyInstaller
        print(f"✓ PyInstaller 版本: {PyInstaller.__version__}")
    except ImportError:
        print("✗ PyInstaller 未安装，正在安装...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    # PyInstaller 参数
    args = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",           # 单文件模式
        "--win-private-assemblies",
        "--windowed",          # GUI 模式，不显示控制台
        "--name", "fcp",       # 输出文件名
        "--clean",             # 清理临时文件
        
        # 隐藏导入（PyInstaller 可能无法自动检测）
        "--hidden-import", "rapidfuzz",
        "--hidden-import", "rapidfuzz.fuzz",
        "--hidden-import", "rapidfuzz.process",
        "--hidden-import", "openpyxl",
        "--hidden-import", "pandas",
        "--hidden-import", "tkinter",
        "--hidden-import", "tkinter.ttk",
        "--hidden-import", "tkinter.filedialog",
        "--hidden-import", "tkinter.messagebox",
        "--hidden-import", "tkinter.scrolledtext",
        
        # 主脚本
        "fcp.py"
    ]
    
    print("\n开始打包...")
    print(f"命令: {' '.join(args)}\n")
    
    try:
        result = subprocess.run(args, check=True, capture_output=False)
        print("\n✓ 打包成功！")
        print(f"输出文件: dist/fcp.exe")
        
        # 显示文件大小
        exe_path = "dist/fcp.exe"
        if os.path.exists(exe_path):
            size_mb = os.path.getsize(exe_path) / (1024 * 1024)
            print(f"文件大小: {size_mb:.1f} MB")
        
    except subprocess.CalledProcessError as e:
        print(f"\n✗ 打包失败，错误代码: {e.returncode}")
        sys.exit(1)
    except Exception as e:
        print(f"\n✗ 打包失败: {e}")
        sys.exit(1)

if __name__ == "__main__":
    build()
