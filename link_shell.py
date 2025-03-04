#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Python Link Shell Extension
用于在Windows系统中创建和管理硬链接、符号链接和目录联接的工具
"""

import os
import sys
import ctypes
import argparse
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.font import Font
import winreg
import subprocess
from pathlib import Path

# 确保以管理员权限运行
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    if not is_admin():
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, " ".join(sys.argv), None, 1
        )
        sys.exit(0)

# 链接操作函数
def create_hardlink(source, target):
    """
    创建硬链接
    :param source: 源文件路径
    :param target: 目标文件路径
    :return: 是否成功
    """
    try:
        if os.path.isfile(source):
            result = ctypes.windll.kernel32.CreateHardLinkW(target, source, None)
            return result != 0
        else:
            messagebox.showerror("错误", "硬链接只能用于文件，不能用于目录")
            return False
    except Exception as e:
        messagebox.showerror("错误", f"创建硬链接失败: {str(e)}")
        return False

def create_symlink(source, target, is_directory=False):
    """
    创建符号链接
    :param source: 源路径
    :param target: 目标路径
    :param is_directory: 是否是目录
    :return: 是否成功
    """
    try:
        if is_directory:
            os.symlink(source, target, target_is_directory=True)
        else:
            os.symlink(source, target)
        return True
    except Exception as e:
        messagebox.showerror("错误", f"创建符号链接失败: {str(e)}")
        return False

def create_junction(source, target):
    """
    创建目录联接
    :param source: 源目录路径
    :param target: 目标目录路径
    :return: 是否成功
    """
    try:
        if os.path.isdir(source):
            # 使用mklink /J命令创建目录联接
            result = subprocess.run(
                ["cmd", "/c", f"mklink /J \"{target}\" \"{source}\""], 
                capture_output=True, 
                text=True, 
                shell=True
            )
            return result.returncode == 0
        else:
            messagebox.showerror("错误", "目录联接只能用于目录，不能用于文件")
            return False
    except Exception as e:
        messagebox.showerror("错误", f"创建目录联接失败: {str(e)}")
        return False

# 检查路径是否是链接
def is_link(path):
    """
    检查路径是否是链接
    :param path: 要检查的路径
    :return: 链接类型或None
    """
    try:
        if os.path.islink(path):
            return "符号链接"
        
        # 检查是否是目录联接
        if os.path.isdir(path):
            result = subprocess.run(
                ["fsutil", "reparsepoint", "query", path], 
                capture_output=True, 
                text=True
            )
            if "重解析点" in result.stdout or "Reparse" in result.stdout:
                return "目录联接"
        
        # 检查是否是硬链接
        if os.path.isfile(path):
            result = subprocess.run(
                ["fsutil", "hardlink", "list", path], 
                capture_output=True, 
                text=True
            )
            if len(result.stdout.strip().split('\n')) > 2:
                return "硬链接"
        
        return None
    except:
        return None

# 获取链接目标
def get_link_target(path):
    """
    获取链接的目标路径
    :param path: 链接路径
    :return: 目标路径或None
    """
    try:
        if os.path.islink(path):
            return os.readlink(path)
        
        # 获取目录联接的目标
        if os.path.isdir(path):
            result = subprocess.run(
                ["fsutil", "reparsepoint", "query", path], 
                capture_output=True, 
                text=True
            )
            for line in result.stdout.split('\n'):
                if "Substitute Name" in line or "替代名称" in line:
                    target = line.split(':')[-1].strip()
                    # 处理输出格式，移除可能的前缀
                    if target.startswith("\\??\\"): 
                        target = target[4:]
                    return target
        
        return None
    except:
        return None

# 图形界面类
class LinkShellApp(tk.Tk):
    def __init__(self):
        super().__init__()
        
        self.title("Python Link Shell Extension")
        self.geometry("700x500")
        self.resizable(True, True)
        
        # 设置样式
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # 创建主框架
        self.main_frame = ttk.Frame(self)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建选项卡
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # 创建链接选项卡
        self.create_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.create_tab, text="创建链接")
        
        # 创建管理选项卡
        self.manage_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.manage_tab, text="管理链接")
        
        # 创建设置选项卡
        self.settings_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_tab, text="设置")
        
        # 初始化各选项卡的内容
        self.init_create_tab()
        self.init_manage_tab()
        self.init_settings_tab()
        
        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        self.status_bar = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def init_create_tab(self):
        # 链接类型选择
        ttk.Label(self.create_tab, text="链接类型:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        
        self.link_type = tk.StringVar()
        self.link_type.set("hardlink")
        
        ttk.Radiobutton(self.create_tab, text="硬链接 (仅文件)", variable=self.link_type, 
                       value="hardlink").grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Radiobutton(self.create_tab, text="符号链接", variable=self.link_type, 
                       value="symlink").grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Radiobutton(self.create_tab, text="目录联接 (仅目录)", variable=self.link_type, 
                       value="junction").grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 源路径选择
        ttk.Label(self.create_tab, text="源路径:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        
        self.source_path = tk.StringVar()
        source_entry = ttk.Entry(self.create_tab, textvariable=self.source_path, width=50)
        source_entry.grid(row=3, column=1, sticky=tk.W+tk.E, padx=5, pady=5)
        
        ttk.Button(self.create_tab, text="浏览...", command=self.browse_source).grid(
            row=3, column=2, sticky=tk.W, padx=5, pady=5)
        
        # 目标路径选择
        ttk.Label(self.create_tab, text="目标路径:").grid(row=4, column=0, sticky=tk.W, padx=5, pady=5)
        
        self.target_path = tk.StringVar()
        target_entry = ttk.Entry(self.create_tab, textvariable=self.target_path, width=50)
        target_entry.grid(row=4, column=1, sticky=tk.W+tk.E, padx=5, pady=5)
        
        ttk.Button(self.create_tab, text="浏览...", command=self.browse_target).grid(
            row=4, column=2, sticky=tk.W, padx=5, pady=5)
        
        # 创建按钮
        create_button = ttk.Button(self.create_tab, text="创建链接", command=self.create_link)
        create_button.grid(row=5, column=1, pady=20)
        
        # 说明文本
        info_text = "\n链接类型说明:\n\n"
        info_text += "硬链接: 只能用于文件，不能跨驱动器，多个文件名指向同一个文件内容\n"
        info_text += "符号链接: 可用于文件或目录，可以跨驱动器，类似快捷方式但更透明\n"
        info_text += "目录联接: 只能用于目录，不能跨驱动器，将一个目录挂载到另一个位置\n"
        
        info_label = ttk.Label(self.create_tab, text=info_text, justify=tk.LEFT, wraplength=650)
        info_label.grid(row=6, column=0, columnspan=3, sticky=tk.W, padx=5, pady=10)
        
        # 配置网格布局
        self.create_tab.columnconfigure(1, weight=1)
    
    def init_manage_tab(self):
        # 路径输入
        ttk.Label(self.manage_tab, text="检查路径:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        
        self.check_path = tk.StringVar()
        check_entry = ttk.Entry(self.manage_tab, textvariable=self.check_path, width=50)
        check_entry.grid(row=0, column=1, sticky=tk.W+tk.E, padx=5, pady=5)
        
        ttk.Button(self.manage_tab, text="浏览...", command=self.browse_check).grid(
            row=0, column=2, sticky=tk.W, padx=5, pady=5)
        
        # 检查按钮
        check_button = ttk.Button(self.manage_tab, text="检查链接", command=self.check_link)
        check_button.grid(row=1, column=1, pady=10)
        
        # 结果显示区域
        ttk.Label(self.manage_tab, text="链接信息:").grid(row=2, column=0, sticky=tk.NW, padx=5, pady=5)
        
        self.result_text = tk.Text(self.manage_tab, height=15, width=70, wrap=tk.WORD)
        self.result_text.grid(row=2, column=1, columnspan=2, sticky=tk.W+tk.E+tk.N+tk.S, padx=5, pady=5)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(self.manage_tab, orient="vertical", command=self.result_text.yview)
        scrollbar.grid(row=2, column=3, sticky=tk.NS)
        self.result_text.configure(yscrollcommand=scrollbar.set)
        
        # 配置网格布局
        self.manage_tab.columnconfigure(1, weight=1)
        self.manage_tab.rowconfigure(2, weight=1)
    
    def init_settings_tab(self):
        # 右键菜单集成选项
        self.integrate_var = tk.BooleanVar()
        self.integrate_var.set(self.check_integration())
        
        ttk.Checkbutton(self.settings_tab, text="集成到Windows资源管理器右键菜单", 
                       variable=self.integrate_var).grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        
        # 应用按钮
        apply_button = ttk.Button(self.settings_tab, text="应用设置", command=self.apply_settings)
        apply_button.grid(row=1, column=0, pady=10)
        
        # 说明文本
        note_text = "\n注意: 集成到右键菜单需要管理员权限，且可能需要重启资源管理器才能生效。\n"
        note_label = ttk.Label(self.settings_tab, text=note_text, justify=tk.LEFT, wraplength=650)
        note_label.grid(row=2, column=0, columnspan=2, sticky=tk.W, padx=5, pady=10)
    
    # 浏览文件/文件夹的方法
    def browse_source(self):
        path = filedialog.askdirectory() if self.link_type.get() == "junction" else filedialog.askopenfilename()
        if path:
            self.source_path.set(path)
    
    def browse_target(self):
        if self.link_type.get() == "hardlink":
            # 硬链接目标是文件
            path = filedialog.asksaveasfilename()
        else:
            # 符号链接和目录联接目标可以是新位置
            path = filedialog.askdirectory() if self.link_type.get() == "junction" else filedialog.asksaveasfilename()
        if path:
            self.target_path.set(path)
    
    def browse_check(self):
        path = filedialog.askopenfilename()
        if path:
            self.check_path.set(path)
    
    # 创建链接
    def create_link(self):
        source = self.source_path.get()
        target = self.target_path.get()
        
        if not source or not target:
            messagebox.showerror("错误", "请指定源路径和目标路径")
            return
        
        link_type = self.link_type.get()
        
        # 检查是否需要管理员权限
        if link_type in ["symlink", "junction"] and not is_admin():
            run_as_admin()
            return
        
        success = False
        if link_type == "hardlink":
            success = create_hardlink(source, target)
        elif link_type == "symlink":
            is_dir = os.path.isdir(source)
            success = create_symlink(source, target, is_dir)
        elif link_type == "junction":
            success = create_junction(source, target)
        
        if success:
            self.status_var.set(f"成功创建{link_type}链接")
            messagebox.showinfo("成功", f"成功创建链接: {target}")
        else:
            self.status_var.set("创建链接失败")
    
    # 检查链接
    def check_link(self):
        path = self.check_path.get()
        
        if not path:
            messagebox.showerror("错误", "请指定要检查的路径")
            return
        
        if not os.path.exists(path):
            messagebox.showerror("错误", "指定的路径不存在")
            return
        
        self.result_text.delete(1.0, tk.END)
        
        link_type = is_link(path)
        if link_type:
            target = get_link_target(path)
            self.result_text.insert(tk.END, f"路径: {path}\n")
            self.result_text.insert(tk.END, f"类型: {link_type}\n")
            if target:
                self.result_text.insert(tk.END, f"目标: {target}\n")
                self.result_text.insert(tk.END, f"目标存在: {'是' if os.path.exists(target) else '否'}\n")
        else:
            self.result_text.insert(tk.END, f"路径 {path} 不是链接")
    
    # 检查是否已集成到右键菜单
    def check_integration(self):
        try:
            key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "*\\shell\\PythonLinkShell")
            winreg.CloseKey(key)
            return True
        except:
            return False
    
    # 应用设置
    def apply_settings(self):
        if self.integrate_var.get():
            self.integrate_to_context_menu()
        else:
            self.remove_from_context_menu()
    
    # 集成到右键菜单
    def integrate_to_context_menu(self):
        if not is_admin():
            run_as_admin()
            return
        
        try:
            # 为文件添加右键菜单
            key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, "*\\shell\\PythonLinkShell")
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "Python Link Shell")
            winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, sys.executable)
            winreg.CloseKey(key)
            
            key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, "*\\shell\\PythonLinkShell\\command")
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, f'"{sys.executable}" "{os.path.abspath(__file__)}" --source "%1"')
            winreg.CloseKey(key)
            
            # 为目录添加右键菜单
            key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, "Directory\\shell\\PythonLinkShell")
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "Python Link Shell")
            winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, sys.executable)
            winreg.CloseKey(key)
            
            key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, "Directory\\shell\\PythonLinkShell\\command")
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, f'"{sys.executable}" "{os.path.abspath(__file__)}" --source "%1"')
            winreg.CloseKey(key)
            
            messagebox.showinfo("成功", "已成功集成到右键菜单")
        except Exception as e:
            messagebox.showerror("错误", f"集成到右键菜单失败: {str(e)}")
    
    # 从右键菜单移除
    def remove_from_context_menu(self):
        if not is_admin():
            run_as_admin()
            return
        
        try:
            # 删除文件右键菜单
            winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, "*\\shell\\PythonLinkShell\\command")
            winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, "*\\shell\\PythonLinkShell")
            
            # 删除目录右键菜单
            winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, "Directory\\shell\\PythonLinkShell\\command")
            winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, "Directory\\shell\\PythonLinkShell")
            
            messagebox.showinfo("成功", "已从右键菜单移除")
        except Exception as e:
            messagebox.showerror("错误", f"从右键菜单移除失败: {str(e)}")

# 命令行参数处理
def parse_arguments():
    # 确保sys.stderr存在，防止'NoneType' object has no attribute 'write'错误
    if sys.stderr is None:
        sys.stderr = open(os.devnull, 'w')
        
    parser = argparse.ArgumentParser(description="Python Link Shell Extension")
    parser.add_argument("--type", choices=["hardlink", "symlink", "junction"], help="链接类型")
    parser.add_argument("--source", help="源路径")
    parser.add_argument("--target", help="目标路径")
    parser.add_argument("files", nargs="*", help="直接指定的文件路径")
    return parser.parse_args()

# 主函数
def main():
    args = parse_arguments()
    
    # 如果提供了命令行参数，则使用命令行模式
    if args.source:
        # 如果只提供了源路径，则启动GUI并预填充源路径
        if not args.type or not args.target:
            app = LinkShellApp()
            app.source_path.set(args.source)
            app.mainloop()
        else:
            # 完整的命令行参数，直接创建链接
            if args.type == "hardlink":
                if create_hardlink(args.source, args.target):
                    print(f"成功创建硬链接: {args.target}")
                else:
                    print("创建硬链接失败")
            elif args.type == "symlink":
                is_dir = os.path.isdir(args.source)
                if create_symlink(args.source, args.target, is_dir):
                    print(f"成功创建符号链接: {args.target}")
                else:
                    print("创建符号链接失败")
            elif args.type == "junction":
                if create_junction(args.source, args.target):
                    print(f"成功创建目录联接: {args.target}")
                else:
                    print("创建目录联接失败")
    # 检查是否有直接指定的文件路径
    elif args.files:
        # 启动GUI并预填充第一个文件路径
        app = LinkShellApp()
        app.source_path.set(args.files[0])
        app.mainloop()
    else:
        # 无命令行参数，启动GUI
        app = LinkShellApp()
        app.mainloop()

if __name__ == "__main__":
    main()