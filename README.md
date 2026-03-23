# cloudreve-manager

# Cloudreve 服务管理工具

[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.6%2B-green.svg)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)

🚀 One-Click, Truly “Out of the Box”
Cloudreve WebDisk Manager is a Windows‑focused tool that encapsulates all the complex deployment and management tasks of Cloudreve into a graphical interface. No need to memorize commands or write scripts—just a few clicks and you’re done:

🚀 一键搞定，真正的“开箱即用”
Cloudreve 网盘服务管理工具，专为 Windows 用户打造，将 Cloudreve 复杂的部署和管理工作全部封装在图形化界面中。无需记忆命令，无需编写脚本，点点鼠标，就能轻松完成：

✅ 一键安装/卸载 Cloudreve 服务（自动配置防火墙、端口、离开模式）

✅ 自动检测端口占用，智能同步配置文件端口

✅ 一键切换数据库：从 SQLite 平滑迁移到 MySQL，或一键切回默认 SQLite

✅ 自动升级：从 GitHub 获取最新版，自动下载、解压、替换、重启服务，全程可视化进度

✅ 配置备份与恢复：支持手动/自动备份，自动排除 uploads 文件夹，并自动导出 MySQL 数据库 SQL 文件

✅ 实时状态监控：底部状态栏始终显示服务运行状态和数据库类型（含 MySQL 用户名/密码）

✅ 多语言支持：自动检测系统语言，支持简体中文和英文，用户可随时切换



## 📸 界面预览

![主界面预览](screenshot.png) 
![主界面预览](screenshot2.png) 
## 🚀 快速开始

### 环境要求

- Windows 7/8/10/11（64 位）
- 已安装 [Cloudreve](https://github.com/cloudreve/cloudreve) 程序文件（cloudreve.exe）
- 如需使用 MySQL，需安装 MySQL 8.0 或 5.7 并确保服务正常运行

### 安装步骤

1. 从 [Releases](https://github.com/wfkwdw/cloudreve-manager/releases) 下载最新版本的整合包。
2. **以管理员身份运行** `CloudreveManager.exe`。
3. 在界面中配置端口（默认 5212），点击“安装网盘”即可完成部署。
4. 安装完成后，工具会自动启动服务并打开网盘页面。

### 使用说明

- **启动/停止服务**：主界面右侧按钮会根据服务状态动态显示“启动服务”或“停止服务”。
- **修改端口**：在端口输入框中输入新端口，点击“检测端口”可检查是否可用，安装时会自动同步至配置文件。
- **打开网盘**：点击“打开网盘”按钮，工具会自动启动服务（如未运行）并打开浏览器访问网盘。
- **升级网盘**：点击“升级网盘”可选择“自动升级”（从 GitHub 拉取最新版）或“手动升级”（选择本地 cloudreve.exe 文件）。
- **MySQL 配置**：点击菜单“文件” → “安装 MySQL 数据库”，按照提示输入 root 密码，工具将自动创建数据库及用户，并更新配置文件。
- **配置备份/恢复**：通过菜单“文件”可备份/恢复配置文件（排除 uploads 目录），若数据库为 MySQL 则同时导出 SQL 文件。

🎉 谁在使用？

个人用户：想快速搭建个人网盘，不想折腾命令行

中小企业：需要稳定可靠的内部文件分享系统

开发者/运维：需要一个可视化的 Cloudreve 管理辅助工具

🛠️ 开发与编译

环境准备：Python 3.6+

安装依赖：
```batch
pip install ttkbootstrap
```
运行源码
```batch
python cloudreve_manager.py
```

### 打包为可执行文件

使用 PyInstaller 打包为单文件无控制台窗口的 EXE（需根据实际 Python 安装路径调整 `--add-data` 参数）：

```batch
pyinstaller -F -w -i cloudreve.ico ^
--add-data "C:\Users\admin\AppData\Local\Programs\Python\Python314\Lib\site-packages\ttkbootstrap;ttkbootstrap" ^
--add-data "cloudreve.ico;." ^
cloudreve_manager.py
```

> **注意**：请将 `C:\Users\admin\AppData\Local\Programs\Python\Python314\Lib\site-packages\ttkbootstrap` 替换为你的 Python 环境中 `ttkbootstrap` 包的实际路径。  
> 可以通过命令 `pip show ttkbootstrap` 查看安装位置。

📄 许可证

本项目基于 MIT 协议开源，详情请参阅 LICENSE 文件。

🙏 致谢

Cloudreve – 优秀的开源网盘系统

winsw – Windows 服务包装器

ttkbootstrap – 现代 tkinter 主题库

注意：本工具仅适用于 Windows 平台，且需要与 Cloudreve 主程序配合使用。使用前请确保已阅读并同意 Cloudreve 的许可协议。
