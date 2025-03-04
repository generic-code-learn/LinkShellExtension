import PyInstaller.__main__
import os

# 获取当前目录的绝对路径
base_path = os.path.abspath(os.path.dirname(__file__))

# 配置PyInstaller参数
PyInstaller.__main__.run([
    'link_shell.py',  # 主程序文件
    '--name=LinkShellExtension',  # 生成的exe名称
    '--onefile',  # 打包成单个exe文件
    '--windowed',  # 使用Windows子系统
    '--uac-admin',  # 请求管理员权限
    '--icon=NONE',  # 默认图标
    f'--distpath={os.path.join(base_path, "dist")}',  # 输出目录
    f'--workpath={os.path.join(base_path, "build")}',  # 临时工作目录
    '--clean',  # 清理临时文件
])