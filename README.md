# Python Link Shell Extension

这是一个用Python实现的类似LinkShellExtension的软件，用于在Windows系统中创建和管理硬链接、符号链接和目录联接。

## 功能特点

- 创建硬链接（Hard Links）
- 创建符号链接（Symbolic Links）
- 创建目录联接（Directory Junctions）
- 图形用户界面，操作简单直观
- 集成到Windows资源管理器右键菜单（可选）
- 查看和管理现有链接

## 系统要求

- Windows 7/8/10/11
- Python 3.6+
- 管理员权限（创建某些类型的链接需要）

## 安装方法

1. 确保已安装Python 3.6或更高版本
2. 安装所需依赖：
   ```
   pip install -r requirements.txt
   ```
3. 运行主程序：
   ```
   python link_shell.py
   ```

## 使用方法

### 图形界面使用

1. 启动程序后，选择要创建的链接类型
2. 选择源文件/文件夹
3. 选择目标位置
4. 点击创建按钮

### 命令行使用

```
python link_shell.py --type [hardlink|symlink|junction] --source [源路径] --target [目标路径]
```

## 注意事项

- 创建符号链接和目录联接可能需要管理员权限
- 删除链接时只会删除链接本身，不会影响原始文件
- 硬链接只能用于同一驱动器上的文件，不能用于目录

## 许可证

MIT