import sys
import os
import time
import re
import subprocess
import ctypes
import threading
import webbrowser
import socket
import shutil
import tkinter.filedialog as filedialog
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import messagebox, simpledialog
import winreg
import json
import urllib.request
import urllib.error
import zipfile
import logging
import atexit
import configparser

# ========== 全局变量 ==========
opened_urls = set()

# ========== 常量定义 ==========
SERVICE_NAME = "100088"
PROCESS_NAME = "cloudreve.exe"
DEFAULT_PORT = 5212
REG_PATH = r"SYSTEM\CurrentControlSet\Control\Session Manager\Power"
REG_VALUE = "AwayModeEnabled"
GITHUB_API_URL = "https://api.github.com/repos/cloudreve/cloudreve/releases/latest"
USER_AGENT = "CloudreveManager/1.0"
RETRY_COUNT = 3
RETRY_BACKOFF = 2

TEMP_DIR = os.path.join(os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__), "temp")

# ========== 日志配置（仅控制台，不生成文件） ==========
def setup_logging():
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    console = logging.StreamHandler()
    formatter = logging.Formatter('[%(asctime)s] %(levelname)s - %(message)s')
    console.setFormatter(formatter)
    logger.addHandler(console)

setup_logging()
log = logging.getLogger(__name__)

# ========== 临时目录清理 ==========
def cleanup_temp():
    try:
        if os.path.exists(TEMP_DIR):
            shutil.rmtree(TEMP_DIR, ignore_errors=True)
            log.info("临时目录已清理")
    except Exception as e:
        log.error(f"清理临时目录失败: {e}")

atexit.register(cleanup_temp)

# ========== 管理员权限提升 ==========
def run_as_admin():
    try:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin()
    except:
        is_admin = False
    if not is_admin:
        log.info("当前非管理员权限，尝试提权...")
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, " ".join(sys.argv), None, 1
        )
        sys.exit(0)

def check_admin_permission():
    if not is_admin():
        messagebox.showwarning("权限不足", "此操作需要管理员权限，请以管理员身份重新运行程序。")
        return False
    return True

# ========== 离开模式注册表操作 ==========
def enable_away_mode():
    try:
        key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            REG_PATH,
            0,
            winreg.KEY_WRITE
        )
        winreg.SetValueEx(key, REG_VALUE, 0, winreg.REG_DWORD, 1)
        winreg.CloseKey(key)
        log.info("离开模式已启用")
        return True
    except PermissionError:
        log.warning("启用离开模式失败：权限不足")
        return False
    except Exception as e:
        log.error(f"启用离开模式异常: {e}")
        return False

def disable_away_mode():
    try:
        key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            REG_PATH,
            0,
            winreg.KEY_WRITE
        )
        try:
            winreg.DeleteValue(key, REG_VALUE)
        except FileNotFoundError:
            pass
        winreg.CloseKey(key)
        log.info("离开模式已禁用")
        return True
    except PermissionError:
        log.warning("禁用离开模式失败：权限不足")
        return False
    except Exception as e:
        log.error(f"禁用离开模式异常: {e}")
        return False

# ========== 带重试的子进程执行（已包含 CREATE_NO_WINDOW） ==========
def run_cmd_with_retry(cmd, timeout=10, retries=RETRY_COUNT):
    for attempt in range(retries):
        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding='gbk',
                errors='ignore',
                creationflags=0x08000000,
                timeout=timeout
            )
            return result
        except subprocess.TimeoutExpired:
            log.warning(f"命令超时（第{attempt+1}次尝试）: {cmd}")
            if attempt == retries - 1:
                raise
            time.sleep(RETRY_BACKOFF)
        except Exception as e:
            log.error(f"命令执行异常（第{attempt+1}次尝试）: {e}")
            if attempt == retries - 1:
                raise
            time.sleep(RETRY_BACKOFF)

# ========== 端口占用进程识别 ==========
def get_process_using_port(port):
    try:
        result = run_cmd_with_retry(['netstat', '-ano'], timeout=5)
        lines = result.stdout.splitlines()
        pid = None
        for line in lines:
            if f":{port}" in line and "LISTENING" in line:
                parts = line.split()
                pid = parts[-1].strip()
                if pid.isdigit():
                    pid = int(pid)
                    break
        if pid is None:
            return None
        result = run_cmd_with_retry(['tasklist', '/FI', f'PID eq {pid}'], timeout=5)
        lines = result.stdout.splitlines()
        for line in lines:
            if str(pid) in line:
                parts = line.split()
                if parts:
                    proc_name = parts[0]
                    return proc_name.lower()
        return None
    except Exception as e:
        log.error(f"获取端口占用进程失败: {e}")
        return None

# ========== 网络检测工具函数 ==========
def is_port_open(host='localhost', port=80, timeout=2):
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(timeout)
        result = sock.connect_ex((host, port))
        sock.close()
        return result == 0
    except Exception as e:
        log.error(f"检测端口{host}:{port}异常: {e}")
        return False

def wait_for_port_open(host='localhost', port=80, max_wait=60, check_interval=2, progress_callback=None):
    start_time = time.time()
    log.info(f"开始等待端口{host}:{port}开放，最长等待{max_wait}秒")
    while time.time() - start_time < max_wait:
        if is_port_open(host, port):
            elapsed = round(time.time() - start_time, 1)
            log.info(f"端口{host}:{port}已开放（等待{elapsed}秒）")
            if progress_callback:
                progress_callback(elapsed, max_wait, 0)
            return True
        elapsed = round(time.time() - start_time, 1)
        remaining = max_wait - elapsed
        if progress_callback:
            progress_callback(elapsed, max_wait, remaining)
        log.debug(f"端口{host}:{port}未开放，{check_interval}秒后重试... (剩余{remaining}秒)")
        time.sleep(check_interval)
    raise TimeoutError(f"等待端口{host}:{port}开放超时（{max_wait}秒）")

def open_url_safely(url, port, max_wait=60):
    global opened_urls
    if url in opened_urls:
        log.info(f"关闭已打开的URL: {url}")
        try:
            if sys.platform == 'win32':
                subprocess.run(['taskkill', '/F', '/IM', 'msedge.exe'], capture_output=True, errors='ignore', timeout=5, creationflags=0x08000000)
                subprocess.run(['taskkill', '/F', '/IM', 'chrome.exe'], capture_output=True, errors='ignore', timeout=5, creationflags=0x08000000)
                subprocess.run(['taskkill', '/F', '/IM', 'firefox.exe'], capture_output=True, errors='ignore', timeout=5, creationflags=0x08000000)
        except:
            pass
        opened_urls.remove(url)
    try:
        wait_for_port_open(host='localhost', port=port, max_wait=max_wait)
        log.info(f"端口localhost:{port}已就绪，打开URL: {url}")
        webbrowser.open(url)
        opened_urls.add(url)
        return True
    except TimeoutError as e:
        log.error(f"打开URL失败: {e}")
        messagebox.showwarning("警告", f"无法连接到 {url}\n端口{port}未开放，请检查服务是否正常启动")
        return False

def is_port_occupied(port, host='localhost'):
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(1)
        result = sock.connect_ex((host, port))
        sock.close()
        if result == 0:
            log.debug(f"端口{port}：有服务在监听（已被占用）")
            return True
        sock2 = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            sock2.bind(('', port))
            sock2.close()
            log.debug(f"端口{port}：未被占用（可正常使用）")
            return False
        except OSError as e:
            if "address already in use" in str(e).lower() or e.errno == 98:
                log.debug(f"端口{port}：绑定失败（已被占用）")
                return True
            else:
                log.debug(f"端口{port}：绑定异常 - {e}（不视为占用）")
                return False
        except Exception as e:
            log.error(f"端口{port}：检测异常 - {e}（默认视为未占用）")
            return False
    except Exception as e:
        log.error(f"端口{port}：检测失败 - {e}（默认视为未占用）")
        return False

# ========== 配置文件操作函数（修复端口重复） ==========
def get_port_from_conf(conf_path, default_port=DEFAULT_PORT):
    try:
        if not os.path.exists(conf_path):
            log.info(f"配置文件{conf_path}不存在，使用默认端口{default_port}")
            return default_port
        with open(conf_path, 'r', encoding='utf-8') as f:
            content = f.read()
        match = re.search(r'Listen\s*=\s*:(\d+)', content, re.IGNORECASE)
        if match:
            port = int(match.group(1))
            log.info(f"从配置文件读取到端口: {port}")
            return port
        else:
            log.info(f"配置文件中未找到有效端口配置，使用默认端口{default_port}")
            return default_port
    except Exception as e:
        log.error(f"读取配置文件端口失败: {e}，使用默认端口{default_port}")
        return default_port

def modify_conf_port(conf_path, port):
    """修改配置文件中的端口，并自动备份原文件，确保只保留一个正确的 Listen 配置"""
    try:
        # 备份原文件
        if os.path.exists(conf_path):
            backup_path = conf_path + ".bak"
            shutil.copy2(conf_path, backup_path)
            log.info(f"已备份原配置文件至 {backup_path}")

        # 使用 configparser 读取，保留大小写和注释
        config = configparser.ConfigParser()
        config.optionxform = str  # 保持键的大小写（默认会转小写）
        config.read(conf_path, encoding='utf-8')

        # 确保 [System] 段存在
        if 'System' not in config:
            config['System'] = {}

        # 删除 [System] 段中所有名称类似 'listen' 的键（不区分大小写）
        keys_to_remove = []
        for key in config['System'].keys():
            if key.lower() == 'listen':
                keys_to_remove.append(key)
        for key in keys_to_remove:
            del config['System'][key]
            log.info(f"已删除旧的 Listen 配置项: {key}")

        # 设置新的 Listen 端口
        config['System']['Listen'] = f":{port}"

        # 写回文件
        with open(conf_path, 'w', encoding='utf-8') as f:
            config.write(f)

        # 验证写入结果
        with open(conf_path, 'r', encoding='utf-8') as f:
            content = f.read()
            if f"Listen = :{port}" in content:
                log.info(f"验证通过：conf.ini中已成功设置 Listen = :{port} 在 [System] 段")
                return True
            else:
                log.error(f"验证失败：conf.ini中未找到 Listen = :{port}")
                return False
    except Exception as e:
        log.error(f"修改conf.ini文件失败: {e}")
        import traceback
        log.error(f"异常详情: {traceback.format_exc()}")
        return False

# ========== 核心功能函数 ==========
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def check_service_exists(service_name):
    try:
        result = run_cmd_with_retry(['sc', 'query', service_name], timeout=10)
        return "STATE" in result.stdout and service_name in result.stdout
    except Exception as e:
        log.error(f"检查服务异常: {e}")
        return False

def get_service_status(service_name):
    try:
        result = run_cmd_with_retry(['sc', 'query', service_name], timeout=10)
        output = result.stdout.lower()
        if "state" not in output:
            return "不存在"
        elif "running" in output:
            return "RUNNING"
        elif "stopped" in output:
            return "STOPPED"
        else:
            return "未知"
    except Exception as e:
        log.error(f"获取服务状态异常: {e}")
        return "未知"

def stop_service(service_name):
    try:
        run_cmd_with_retry(['sc', 'stop', service_name], timeout=10)
        time.sleep(2)
        if get_service_status(service_name) == "STOPPED":
            return True
        log.warning("服务停止超时，尝试强制结束进程")
        kill_process(PROCESS_NAME)
        return True
    except Exception as e:
        log.error(f"停止服务异常: {e}")
        try:
            kill_process(PROCESS_NAME)
            return True
        except:
            return False

def start_service(service_name):
    try:
        run_cmd_with_retry(['sc', 'start', service_name], timeout=10)
        time.sleep(2)
        return True
    except Exception as e:
        log.error(f"启动服务异常: {e}")
        return False

def check_process_exists(process_name):
    try:
        result = run_cmd_with_retry(['tasklist', '/FI', f'IMAGENAME eq {process_name}'], timeout=10)
        return process_name.lower() in result.stdout.lower()
    except Exception as e:
        log.error(f"检查进程异常: {e}")
        return False

def kill_process(process_name):
    try:
        run_cmd_with_retry(['taskkill', '/F', '/IM', process_name], timeout=10)
        time.sleep(2)
        return True
    except Exception as e:
        log.error(f"结束进程异常: {e}")
        return False

def add_firewall_rule(port, rule_name="Cloudreve_Port"):
    try:
        run_cmd_with_retry(['cmd', '/c', f'netsh advfirewall firewall delete rule name="{rule_name}_Inbound"'], timeout=10)
        run_cmd_with_retry(['cmd', '/c', f'netsh advfirewall firewall delete rule name="{rule_name}_Outbound"'], timeout=10)
        inbound_cmd = f'netsh advfirewall firewall add rule name="{rule_name}_Inbound" dir=in action=allow protocol=TCP localport={port} enable=yes'
        inbound_result = run_cmd_with_retry(['cmd', '/c', inbound_cmd], timeout=10)
        outbound_cmd = f'netsh advfirewall firewall add rule name="{rule_name}_Outbound" dir=out action=allow protocol=TCP localport={port} enable=yes'
        outbound_result = run_cmd_with_retry(['cmd', '/c', outbound_cmd], timeout=10)
        success = True
        error_info = ""
        if "错误" in inbound_result.stderr or "失败" in inbound_result.stderr:
            success = False
            error_info += f"入站规则：{inbound_result.stderr.strip()}\n"
        if "错误" in outbound_result.stderr or "失败" in outbound_result.stderr:
            success = False
            error_info += f"出站规则：{outbound_result.stderr.strip()}\n"
        if success and (inbound_result.returncode != 0 or outbound_result.returncode != 0):
            error_info = "命令返回码非0，但未检测到明确错误"
        return (success, error_info.strip())
    except Exception as e:
        log.error(f"添加防火墙规则异常: {e}")
        return (False, f"执行异常：{str(e)}")

def wait_for_service_start(service_name, process_name, timeout=30, check_interval=2, status_callback=None, progress_callback=None):
    start_time = time.time()
    max_retries = int(timeout / check_interval)
    retry_count = 0
    log.info(f"开始等待服务启动，最长等待{timeout}秒")
    while time.time() - start_time < timeout:
        service_status = get_service_status(service_name)
        process_running = check_process_exists(process_name)
        elapsed = round(time.time() - start_time, 1)
        retry_count += 1
        if status_callback:
            status_callback()
        if progress_callback:
            progress_callback(retry_count, max_retries, elapsed, timeout)
        if service_status == "RUNNING" and process_running:
            log.info(f"服务启动成功，耗时{elapsed}秒")
            return (True, f"服务状态：{service_status}，进程：存在", elapsed)
        log.debug(f"服务未就绪（{elapsed}s），{check_interval}秒后重试...")
        time.sleep(check_interval)
    final_status = get_service_status(service_name)
    final_process = check_process_exists(process_name)
    return (
        False,
        f"超时（{timeout}s）| 服务状态：{final_status} | 进程：{'存在' if final_process else '不存在'}",
        timeout
    )

def wait_for_conf_file(conf_path, timeout=60, check_interval=2, progress_callback=None):
    start_time = time.time()
    elapsed_time = 0
    while elapsed_time < timeout:
        if os.path.exists(conf_path):
            log.info(f"检测到conf.ini文件生成，耗时{elapsed_time}秒")
            if progress_callback:
                progress_callback(elapsed_time, timeout, 0)
            return True
        elapsed_time = round(time.time() - start_time, 1)
        remaining = timeout - elapsed_time
        if progress_callback:
            progress_callback(elapsed_time, timeout, remaining)
        log.debug(f"等待conf.ini文件生成... ({elapsed_time}/{timeout}秒)")
        time.sleep(check_interval)
    log.error(f"等待conf.ini文件生成超时（{timeout}秒）")
    return False

# ========== 自动升级相关函数 ==========
def get_current_version(exe_path):
    try:
        result = subprocess.run(
            [exe_path, '--version'],
            capture_output=True,
            text=True,
            timeout=10,
            encoding='utf-8',
            errors='ignore',
            creationflags=0x08000000
        )
        output = result.stdout.strip() or result.stderr.strip()
        match = re.search(r'v?(\d+\.\d+\.\d+)', output)
        if match:
            return match.group(1)
        return None
    except Exception as e:
        log.error(f"获取当前版本失败: {e}")
        return None

def get_latest_version_from_github():
    for attempt in range(RETRY_COUNT):
        try:
            req = urllib.request.Request(GITHUB_API_URL, headers={'User-Agent': USER_AGENT})
            with urllib.request.urlopen(req, timeout=10) as response:
                if response.getcode() != 200:
                    raise ConnectionError(f"GitHub API返回HTTP {response.getcode()}")
                data = json.loads(response.read().decode('utf-8'))
                tag = data.get('tag_name', '')
                if tag.startswith('v'):
                    tag = tag[1:]
                return tag
        except urllib.error.HTTPError as e:
            if e.code == 403:
                raise ConnectionError("GitHub API访问受限，可能是请求频率过高，请稍后再试或手动升级") from e
            raise
        except Exception as e:
            log.warning(f"获取最新版本失败（第{attempt+1}次尝试）: {e}")
            if attempt == RETRY_COUNT - 1:
                raise
            time.sleep(RETRY_BACKOFF * (attempt + 1))

def download_file(url, dest_path, progress_callback=None):
    for attempt in range(RETRY_COUNT):
        try:
            req = urllib.request.Request(url, headers={'User-Agent': USER_AGENT})
            with urllib.request.urlopen(req, timeout=30) as response:
                total_size = int(response.headers.get('Content-Length', 0))
                downloaded = 0
                with open(dest_path, 'wb') as f:
                    while True:
                        chunk = response.read(8192)
                        if not chunk:
                            break
                        f.write(chunk)
                        downloaded += len(chunk)
                        if progress_callback and total_size:
                            percent = int(100 * downloaded / total_size)
                            progress_callback(percent, downloaded, total_size)
            return True
        except Exception as e:
            log.warning(f"下载失败（第{attempt+1}次尝试）: {e}")
            if os.path.exists(dest_path):
                os.remove(dest_path)
            if attempt == RETRY_COUNT - 1:
                raise
            time.sleep(RETRY_BACKOFF * (attempt + 1))

def compare_versions(ver1, ver2):
    def parse(v):
        return [int(x) for x in v.split('.')]
    v1 = parse(ver1)
    v2 = parse(ver2)
    while len(v1) < len(v2):
        v1.append(0)
    while len(v2) < len(v1):
        v2.append(0)
    for a, b in zip(v1, v2):
        if a < b:
            return True
        elif a > b:
            return False
    return False

def find_asset_url(assets):
    candidates = []
    for asset in assets:
        name = asset.get('name', '')
        if 'windows' in name.lower() and 'amd64' in name.lower() and name.endswith('.zip'):
            return asset.get('browser_download_url')
    for asset in assets:
        name = asset.get('name', '')
        if 'windows' in name.lower() and name.endswith('.zip'):
            candidates.append(asset)
    if not candidates:
        for asset in assets:
            name = asset.get('name', '')
            if name.endswith('.zip'):
                candidates.append(asset)
        if not candidates:
            return None
    for asset in candidates:
        name = asset.get('name', '').lower()
        if 'amd64' in name or '64bit' in name or 'x64' in name:
            return asset.get('browser_download_url')
    return candidates[0].get('browser_download_url')

def get_download_url_from_github():
    for attempt in range(RETRY_COUNT):
        try:
            req = urllib.request.Request(GITHUB_API_URL, headers={'User-Agent': USER_AGENT})
            with urllib.request.urlopen(req, timeout=10) as response:
                if response.getcode() != 200:
                    raise ConnectionError(f"GitHub API返回HTTP {response.getcode()}")
                data = json.loads(response.read().decode('utf-8'))
                assets = data.get('assets', [])
                asset_names = [a.get('name') for a in assets]
                log.info(f"GitHub assets: {asset_names}")
                if not assets:
                    log.warning("警告：GitHub API 返回的 assets 列表为空")
                    raise RuntimeError("GitHub 返回的发布资产列表为空，请稍后重试或手动升级")
                download_url = find_asset_url(assets)
                if not download_url:
                    raise RuntimeError("未找到适用于Windows的下载链接，请检查GitHub发布页")
                return download_url
        except Exception as e:
            log.warning(f"获取下载链接失败（第{attempt+1}次尝试）: {e}")
            if attempt == RETRY_COUNT - 1:
                raise
            time.sleep(RETRY_BACKOFF * (attempt + 1))

# ========== MySQL 备份辅助函数 ==========
def find_mysqldump():
    """尝试从常见路径或注册表查找 mysqldump.exe"""
    common_paths = [
        r"C:\Program Files\MySQL\MySQL Server 8.0\bin\mysqldump.exe",
        r"C:\Program Files\MySQL\MySQL Server 5.7\bin\mysqldump.exe",
        r"C:\Program Files (x86)\MySQL\MySQL Server 8.0\bin\mysqldump.exe",
        r"C:\Program Files (x86)\MySQL\MySQL Server 5.7\bin\mysqldump.exe",
    ]
    for path in common_paths:
        if os.path.isfile(path):
            return path
    # 尝试从注册表读取
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\MySQL AB\MySQL Server 8.0")
        install_dir = winreg.QueryValueEx(key, "Location")[0]
        mysqldump = os.path.join(install_dir, "bin", "mysqldump.exe")
        if os.path.isfile(mysqldump):
            return mysqldump
    except:
        pass
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\MySQL AB\MySQL Server 5.7")
        install_dir = winreg.QueryValueEx(key, "Location")[0]
        mysqldump = os.path.join(install_dir, "bin", "mysqldump.exe")
        if os.path.isfile(mysqldump):
            return mysqldump
    except:
        pass
    return None

def dump_mysql_database(conf_path, output_sql_path):
    """从配置文件读取 MySQL 连接信息，导出数据库到指定 SQL 文件。返回 (成功, 错误信息)"""
    try:
        config = configparser.ConfigParser()
        config.read(conf_path, encoding='utf-8')
        if 'Database' not in config or config['Database'].get('Type', '').lower() != 'mysql':
            return False, "当前不是 MySQL 数据库"
        host = config['Database'].get('Host', '127.0.0.1')
        port = config['Database'].get('Port', '3306')
        user = config['Database'].get('User', '')
        password = config['Database'].get('Password', '')
        dbname = config['Database'].get('Name', 'cloudreve')
        if not user or not password:
            return False, "数据库用户名或密码为空"
        mysqldump = find_mysqldump()
        if not mysqldump:
            return False, "未找到 mysqldump.exe，请确保 MySQL 已安装并正确配置"
        cmd = [mysqldump, f"--host={host}", f"--port={port}", f"--user={user}", f"--password={password}", dbname]
        with open(output_sql_path, 'w', encoding='utf-8') as f:
            result = subprocess.run(cmd, stdout=f, stderr=subprocess.PIPE, timeout=60,
                                    text=True, encoding='utf-8', errors='ignore',
                                    creationflags=0x08000000)
        if result.returncode == 0:
            return True, ""
        else:
            return False, result.stderr
    except Exception as e:
        return False, str(e)

# ========== GUI类 ==========
class CloudreveManagerGUI:
    def __init__(self, root):
        log.info("初始化Cloudreve服务管理工具")
        self.root = root
        self.root.title("Cloudreve 网盘服务管理工具")
        
        # 启用高DPI感知
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except:
            try:
                ctypes.windll.shcore.SetProcessDpiAwareness(1)
            except:
                pass
        
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        win_height = int(screen_height * 0.5)
        win_width = int(win_height * 0.6)
        min_width = 620
        min_height = 700
        if win_width < min_width:
            win_width = min_width
        if win_height < min_height:
            win_height = min_height
        x = (screen_width - win_width) // 2
        y = (screen_height - win_height) // 2
        self.root.geometry(f"{win_width}x{win_height}+{x}+{y}")
        self.root.minsize(min_width, min_height)
        
        # 图标
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            icon_path = os.path.join(script_dir, "cloudreve.ico")
            self.root.iconbitmap(icon_path)
            log.info(f"窗口图标加载成功：{icon_path}")
        except Exception as e:
            log.error(f"窗口图标加载失败: {e}")
        
        self.SERVICE_NAME = SERVICE_NAME
        self.PROCESS_NAME = PROCESS_NAME
        self.DEFAULT_PORT = DEFAULT_PORT
        self.custom_port = ttk.StringVar(value=str(self.DEFAULT_PORT))
        
        try:
            if getattr(sys, 'frozen', False):
                self.APP_DIR = os.path.dirname(sys.executable)
            else:
                self.APP_DIR = os.path.dirname(os.path.abspath(__file__))
            self.WINSW_EXE = os.path.join(self.APP_DIR, "winsw.exe")
            self.WINSW_XML = os.path.join(self.APP_DIR, "winsw.xml")
            self.DATA_DIR = os.path.join(self.APP_DIR, "data")
            self.CONF_PATH = os.path.join(self.DATA_DIR, "conf.ini")
            self.CLOUDREVE_EXE = os.path.join(self.APP_DIR, self.PROCESS_NAME)
        except Exception as e:
            log.error(f"获取程序路径异常: {e}")
            messagebox.showerror("错误", f"获取程序路径失败：{e}")
        
        # 用于 MySQL 配置的线程同步
        self.mysql_app_user = None
        self.mysql_app_pass = None
        self.mysql_cred_event = threading.Event()
        
        self.create_optimized_ui()
        self.refresh_service_status()

    def create_optimized_ui(self):
        main_container = ttk.Frame(self.root, padding=15)
        main_container.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        title_frame = ttk.LabelFrame(main_container, text="安装提醒")
        title_frame.pack(fill=X, pady=(0, 8))
        is_on_c_drive = self.APP_DIR.startswith("C:") or self.APP_DIR.startswith("c:")
        if is_on_c_drive:
            reminder_text = "⚠️ 因数据文件保存在本文件夹的data目录，勿在C盘运行安装程序"
            reminder_style = WARNING
        else:
            reminder_text = "✅ 数据文件保存在本文件夹的data目录，检测到不是C盘，可以安装"
            reminder_style = SUCCESS
        sub_title = ttk.Label(title_frame, text=reminder_text, font=("微软雅黑", 9), bootstyle=reminder_style)
        sub_title.pack(pady=8, padx=10)
        
        integrated_frame = ttk.LabelFrame(main_container, text="端口配置|服务状态")
        integrated_frame.pack(fill=X, pady=(0, 8))
        
        left_frame = ttk.Frame(integrated_frame)
        left_frame.pack(side=LEFT, fill=BOTH, expand=True, padx=10, pady=8)
        right_frame = ttk.Frame(integrated_frame)
        right_frame.pack(side=RIGHT, fill=BOTH, expand=True, padx=10, pady=8)
        
        port_label = ttk.Label(left_frame, text="配置端口：", font=("微软雅黑", 11))
        port_label.grid(row=0, column=0, padx=(0, 8), sticky="w")
        self.port_entry = ttk.Entry(left_frame, textvariable=self.custom_port, font=("微软雅黑", 11), width=10, justify=CENTER)
        self.port_entry.grid(row=0, column=1, padx=(0, 10), sticky="w")
        check_btn = ttk.Button(left_frame, text="检测端口", command=self.check_port_status, bootstyle=INFO, width=10)
        check_btn.grid(row=0, column=2, sticky="w")
        
        ttk.Label(right_frame, text="网盘服务：", font=("微软雅黑", 11)).grid(row=0, column=0, padx=(0, 5), sticky="w")
        self.service_status_label = ttk.Label(right_frame, text="未知", font=("微软雅黑", 10, "bold"))
        self.service_status_label.grid(row=0, column=1, padx=(0, 10), sticky="w")
        self.action_btn = ttk.Button(right_frame, text="检测中", command=lambda: None, width=10)
        self.action_btn.grid(row=0, column=2, sticky="w")
        
        progress_frame = ttk.LabelFrame(main_container, text="执行进度")
        progress_frame.pack(fill=X, pady=(0, 8))
        progress_inner = ttk.Frame(progress_frame)
        progress_inner.pack(fill=X, padx=10, pady=8)
        self.progress_var = ttk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_inner, variable=self.progress_var, maximum=100, bootstyle="SUCCESS.PROGRESS")
        self.progress_bar.pack(fill=X, padx=5, pady=(5, 3))
        self.progress_label = ttk.Label(progress_inner, text="准备就绪", font=("微软雅黑", 9))
        self.progress_label.pack(fill=X, padx=5, pady=(0, 5))
        
        result_frame = ttk.LabelFrame(main_container, text="执行结果")
        result_frame.pack(fill=BOTH, expand=True, pady=(0, 10))
        result_inner = ttk.Frame(result_frame)
        result_inner.pack(fill=BOTH, expand=True, padx=10, pady=6)
        result_scroll = ttk.Scrollbar(result_inner)
        result_scroll.pack(side=RIGHT, fill=Y)
        self.result_text = ttk.Text(result_inner, font=("微软雅黑", 9), wrap=WORD, yscrollcommand=result_scroll.set, height=8)
        self.result_text.pack(fill=BOTH, expand=True, padx=(0, 5))
        result_scroll.config(command=self.result_text.yview)
        self.result_text.config(state=DISABLED)
        self.append_result_text("📌 准备就绪，等待操作...", "info")
        
        menubar = ttk.Menu(self.root)
        self.root.config(menu=menubar)
        
        file_menu = ttk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="配置备份", command=self.backup_data)
        file_menu.add_command(label="配置恢复", command=self.restore_data)
        file_menu.add_separator()
        file_menu.add_command(label="编辑配置文件", command=self.edit_config)
        file_menu.add_separator()
        file_menu.add_command(label="安装 MySQL 数据库", command=self.install_mysql_database)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.root.quit)
        
        help_menu = ttk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助", menu=help_menu)
        help_menu.add_command(label="关于", command=self.show_about)
        
        button_container = ttk.Frame(main_container)
        btn_grid = ttk.Frame(button_container)
        btn_grid.pack(fill=X, padx=20)
        btn_grid.grid_columnconfigure(0, weight=1)
        btn_grid.grid_columnconfigure(1, weight=1)
        btn_grid.grid_columnconfigure(2, weight=1)
        btn_grid.grid_columnconfigure(3, weight=1)
        
        self.install_btn = ttk.Button(btn_grid, text="安装网盘", command=self.start_install, bootstyle=SUCCESS, padding=(15, 8), width=12)
        self.install_btn.grid(row=0, column=0, padx=(0, 6), pady=8, sticky="ew")
        
        self.uninstall_btn = ttk.Button(btn_grid, text="卸载网盘", command=self.start_uninstall, bootstyle=DANGER, padding=(15, 8), width=12)
        self.uninstall_btn.grid(row=0, column=1, padx=(6, 6), pady=8, sticky="ew")
        
        self.open_btn = ttk.Button(btn_grid, text="打开网盘", command=self.start_open_cloudreve, bootstyle=INFO, padding=(15, 8), width=12)
        self.open_btn.grid(row=0, column=2, padx=(6, 6), pady=8, sticky="ew")
        
        self.upgrade_btn = ttk.Button(btn_grid, text="升级网盘", command=self.start_upgrade, bootstyle=INFO, padding=(15, 8), width=12)
        self.upgrade_btn.grid(row=0, column=3, padx=(6, 0), pady=8, sticky="ew")
        
        button_container.pack(side=BOTTOM, fill=X, pady=(0, 5))
        
        self.status_bar = ttk.Label(self.root, text="就绪", relief=SUNKEN, anchor=W)
        self.status_bar.pack(side=BOTTOM, fill=X)

    # ========== 辅助方法 ==========
    def update_statusbar(self, text, bootstyle=None):
        """更新状态栏文字，并可设置样式（success/warning/danger/info）"""
        self.status_bar.config(text=text)
        if bootstyle:
            style = ttk.Style()
            if bootstyle == "success":
                self.status_bar.config(foreground="#2F855A")
            elif bootstyle == "warning":
                self.status_bar.config(foreground="#DD6B20")
            elif bootstyle == "danger":
                self.status_bar.config(foreground="#C53030")
            elif bootstyle == "info":
                self.status_bar.config(foreground="#2B6CB0")
            else:
                self.status_bar.config(foreground="#2D3748")
        self.root.update_idletasks()

    def append_result_text(self, text, bootstyle=None):
        try:
            self.result_text.config(state=NORMAL)
            if bootstyle == "success":
                self.result_text.tag_configure("colored", foreground="#2F855A")
            elif bootstyle == "warning":
                self.result_text.tag_configure("colored", foreground="#DD6B20")
            elif bootstyle == "danger":
                self.result_text.tag_configure("colored", foreground="#C53030")
            elif bootstyle == "info":
                self.result_text.tag_configure("colored", foreground="#2B6CB0")
            else:
                self.result_text.tag_configure("colored", foreground="#2D3748")
            self.result_text.insert(END, text + "\n", "colored")
            self.result_text.see(END)
            self.result_text.config(state=DISABLED)
        except Exception as e:
            log.error(f"追加结果文本异常: {e}")

    def clear_result_text(self):
        try:
            self.result_text.config(state=NORMAL)
            self.result_text.delete(1.0, END)
            self.result_text.config(state=DISABLED)
        except Exception as e:
            log.error(f"清空结果文本异常: {e}")

    def update_progress(self, value, text):
        try:
            self.progress_var.set(value)
            self.progress_label.config(text=text)
            self.root.update_idletasks()
            progress_info = f"[{value}%] {text}"
            log.info(progress_info)
            self.append_result_text(progress_info, "info")
        except Exception as e:
            log.error(f"更新进度异常: {e}")

    # ========== 动态按钮相关 ==========
    def update_action_button(self):
        conf_exists = os.path.exists(self.CONF_PATH)
        service_exists = check_service_exists(self.SERVICE_NAME)
        service_running = (get_service_status(self.SERVICE_NAME) == "RUNNING" and 
                           check_process_exists(self.PROCESS_NAME))
        
        if not conf_exists:
            self.action_btn.config(text="未装网盘", bootstyle="secondary", state=DISABLED, command=None)
        elif not service_exists or not service_running:
            self.action_btn.config(text="启动服务", bootstyle="success", state=NORMAL, command=self.start_service_action)
        else:
            self.action_btn.config(text="停止服务", bootstyle="danger", state=NORMAL, command=self.stop_service_action)

    def start_service_action(self):
        self.append_result_text("⏳ 正在启动服务...", "info")
        self.update_statusbar("启动服务中...")
        self.action_btn.config(state=DISABLED)
        threading.Thread(target=self._start_service_worker, daemon=True).start()

    def _start_service_worker(self):
        try:
            self.root.after(0, lambda: self.update_progress(5, "检查服务状态..."))
            service_exists = check_service_exists(self.SERVICE_NAME)
            if not service_exists:
                self.root.after(0, lambda: self.update_progress(10, "安装服务..."))
                self._install_service_only()
                self.root.after(0, lambda: self.update_progress(30, "服务安装完成，准备启动..."))
            else:
                self.root.after(0, lambda: self.update_progress(30, "准备启动服务..."))
            
            success = start_service(self.SERVICE_NAME)
            if not success:
                self.root.after(0, lambda: self.append_result_text("❌ 启动服务命令执行失败", "danger"))
                self.root.after(0, lambda: messagebox.showerror("启动失败", "启动服务命令执行失败，请检查系统服务。"))
                return

            def refresh_callback():
                self.root.after(0, self.refresh_service_status)

            def progress_callback(retry_count, max_retries, elapsed, timeout):
                progress = 30 + int(60 * retry_count / max_retries)
                self.root.after(0, self.update_progress, progress, f"等待服务启动... ({retry_count}/{max_retries})")

            service_started, status_msg, elapsed = wait_for_service_start(
                self.SERVICE_NAME,
                self.PROCESS_NAME,
                status_callback=refresh_callback,
                progress_callback=progress_callback
            )
            if not service_started:
                self.root.after(0, lambda: self.append_result_text(f"⚠️ 服务启动超时：{status_msg}", "warning"))
                self.root.after(0, lambda: messagebox.showwarning("启动失败", f"网盘服务启动失败：{status_msg}"))
            else:
                self.root.after(0, lambda: self.append_result_text(f"✅ 服务启动成功（耗时{elapsed}秒）", "success"))
                self.root.after(0, lambda: self.update_progress(100, "启动完成"))
        except Exception as e:
            log.error(f"启动服务失败: {e}")
            self.root.after(0, lambda: self.append_result_text(f"❌ 启动服务失败：{e}", "danger"))
            self.root.after(0, lambda: messagebox.showerror("启动失败", f"网盘服务启动失败：{e}"))
        finally:
            self.root.after(0, self.refresh_service_status)

    def _install_service_only(self):
        try:
            if check_process_exists(self.PROCESS_NAME):
                kill_process(self.PROCESS_NAME)
                self.root.after(0, lambda: self.append_result_text("✅ 已结束残留进程", "success"))
            run_cmd_with_retry([self.WINSW_EXE, "install"], timeout=30)
            log.info("服务安装完成")
            self.root.after(0, lambda: self.append_result_text("✅ 服务安装完成", "success"))
        except Exception as e:
            log.error(f"安装服务失败: {e}")
            self.root.after(0, lambda: self.append_result_text(f"❌ 安装服务失败：{e}", "danger"))
            raise

    def stop_service_action(self):
        self.append_result_text("⏳ 正在停止服务...", "info")
        self.update_statusbar("停止服务中...")
        self.action_btn.config(state=DISABLED)
        threading.Thread(target=self._stop_service_worker, daemon=True).start()

    def _stop_service_worker(self):
        try:
            self.root.after(0, lambda: self.update_progress(20, "停止服务..."))
            success = stop_service(self.SERVICE_NAME)
            if not success:
                self.root.after(0, lambda: self.append_result_text("❌ 停止服务失败", "danger"))
                self.root.after(0, lambda: messagebox.showerror("停止失败", "停止服务命令执行失败，请手动检查。"))
                return

            self.root.after(0, lambda: self.update_progress(40, "等待服务停止..."))
            for i in range(10):
                if get_service_status(self.SERVICE_NAME) != "RUNNING":
                    break
                progress = 40 + int(50 * (i + 1) / 10)
                self.root.after(0, self.update_progress, progress, f"等待服务停止... ({i+1}/10)")
                time.sleep(1)
            
            self.root.after(0, lambda: self.append_result_text("✅ 服务已停止", "success"))
            self.root.after(0, lambda: self.update_progress(100, "停止完成"))
        except Exception as e:
            log.error(f"停止服务失败: {e}")
            self.root.after(0, lambda: self.append_result_text(f"❌ 停止服务失败：{e}", "danger"))
            self.root.after(0, lambda: messagebox.showerror("停止失败", f"停止服务失败：{e}"))
        finally:
            self.root.after(0, self.refresh_service_status)

    # ========== 服务状态刷新（包含数据库类型显示） ==========
    def refresh_service_status(self):
        try:
            service_status = get_service_status(self.SERVICE_NAME)
            process_exists = check_process_exists(self.PROCESS_NAME)
            if service_status == "RUNNING" and process_exists:
                status_text = "已运行"
                bootstyle = "success"
            else:
                status_text = "未运行"
                bootstyle = "danger"
            self.service_status_label.config(text=status_text, bootstyle=bootstyle)

            # 获取数据库类型及用户名密码
            db_type = "未知"
            db_color = "secondary"
            db_extra = ""
            if os.path.exists(self.CONF_PATH):
                try:
                    config = configparser.ConfigParser()
                    config.read(self.CONF_PATH, encoding='utf-8')
                    if 'Database' in config and 'Type' in config['Database']:
                        db_type = config['Database']['Type']
                        if db_type.lower() == 'mysql':
                            db_color = "success"
                            # 读取用户名和密码
                            user = config['Database'].get('User', '')
                            pwd = config['Database'].get('Password', '')
                            if user and pwd:
                                db_extra = f" 用户名:{user} 密码:{pwd}"
                        else:
                            db_color = "info"
                    else:
                        db_type = "sqlite3"
                        db_color = "info"
                except Exception as e:
                    log.error(f"读取配置文件失败: {e}")
                    db_type = "读取失败"
                    db_color = "danger"
            status_msg = f"服务状态：{status_text} | 数据库：{db_type}{db_extra}"
            self.update_statusbar(status_msg, bootstyle=db_color)

            log.info(f"刷新状态 - 网盘服务：{status_text}，数据库：{db_type}")
        except Exception as e:
            log.error(f"刷新状态异常: {e}")
            self.service_status_label.config(text="错误", bootstyle="danger")
            self.update_statusbar("状态获取失败", bootstyle="danger")
        finally:
            self.update_action_button()

    def validate_port(self):
        try:
            port_str = self.custom_port.get().strip()
            if not port_str.isdigit():
                return False, "端口号必须是数字"
            port = int(port_str)
            if port < 1 or port > 65535:
                return False, "端口号必须在1-65535之间"
            return True, port
        except Exception as e:
            return False, str(e)

    # ========== 备份恢复增强 ==========
    def backup_data(self):
        """配置备份：排除 uploads 文件夹，如果数据库为 MySQL 则导出 SQL 文件。默认保存到程序目录"""
        if not os.path.exists(self.DATA_DIR):
            messagebox.showwarning("备份失败", "data 目录不存在，无法备份。")
            return
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        default_filename = f"cloudreve_config_backup_{timestamp}.zip"
        backup_path = filedialog.asksaveasfilename(
            title="选择配置文件备份保存位置",
            defaultextension=".zip",
            filetypes=[("ZIP 压缩文件", "*.zip"), ("所有文件", "*.*")],
            initialdir=self.APP_DIR,  # 默认保存到程序当前目录
            initialfile=default_filename
        )
        if not backup_path:
            return
        self.update_statusbar("正在备份配置...")
        try:
            # 创建临时目录用于存放要打包的内容
            temp_backup_dir = os.path.join(TEMP_DIR, "backup_temp")
            if os.path.exists(temp_backup_dir):
                shutil.rmtree(temp_backup_dir)
            os.makedirs(temp_backup_dir, exist_ok=True)

            # 复制 data 目录（排除 uploads）
            dest_data = os.path.join(temp_backup_dir, "data")
            shutil.copytree(self.DATA_DIR, dest_data, ignore=shutil.ignore_patterns("uploads"))

            # 如果数据库是 MySQL，导出 SQL
            sql_file = None
            if os.path.exists(self.CONF_PATH):
                config = configparser.ConfigParser()
                config.read(self.CONF_PATH, encoding='utf-8')
                if 'Database' in config and config['Database'].get('Type', '').lower() == 'mysql':
                    self.append_result_text("📦 检测到 MySQL 数据库，正在导出 SQL...", "info")
                    sql_temp = os.path.join(temp_backup_dir, f"mysql_dump_{config['Database'].get('Name', 'cloudreve')}.sql")
                    success, err = dump_mysql_database(self.CONF_PATH, sql_temp)
                    if success:
                        sql_file = sql_temp
                        self.append_result_text("✅ MySQL 数据库导出成功", "success")
                    else:
                        self.append_result_text(f"⚠️ MySQL 导出失败：{err}，备份将不包含数据库数据", "warning")
                else:
                    self.append_result_text("ℹ️ 当前使用 SQLite 数据库，数据已包含在 data 目录中", "info")

            # 打包
            with zipfile.ZipFile(backup_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, dirs, files in os.walk(temp_backup_dir):
                    for file in files:
                        full_path = os.path.join(root, file)
                        arcname = os.path.relpath(full_path, start=temp_backup_dir)
                        zipf.write(full_path, arcname)

            self.append_result_text(f"✅ 配置备份成功：{backup_path}", "success")
            self.update_statusbar("备份完成")
            # 成功弹窗提醒
            messagebox.showinfo("备份成功", f"配置已成功备份到：\n{backup_path}", parent=self.root)
        except Exception as e:
            log.error(f"备份配置失败: {e}")
            messagebox.showerror("备份失败", f"备份配置时发生错误：{str(e)}")
            self.append_result_text(f"❌ 配置备份失败：{str(e)}", "danger")
            self.update_statusbar("备份失败")
        finally:
            # 清理临时目录
            if os.path.exists(temp_backup_dir):
                shutil.rmtree(temp_backup_dir, ignore_errors=True)

    def restore_data(self):
        """配置恢复：保留现有 uploads 文件夹，检测 SQL 文件并提示"""
        zip_path = filedialog.askopenfilename(
            title="选择备份文件",
            filetypes=[("ZIP 压缩文件", "*.zip"), ("所有文件", "*.*")]
        )
        if not zip_path:
            return
        if not os.path.exists(self.DATA_DIR):
            if not messagebox.askyesno("确认恢复", "data 目录不存在，将创建新目录并恢复，是否继续？"):
                return
        else:
            uploads_path = os.path.join(self.DATA_DIR, "uploads")
            if os.path.exists(uploads_path):
                if not messagebox.askyesno("确认恢复", "恢复操作将覆盖除 uploads 外的所有配置和数据库文件。\n是否继续？"):
                    return
            else:
                if not messagebox.askyesno("确认恢复", "恢复操作将覆盖现有配置文件，是否继续？"):
                    return

        self.update_statusbar("正在恢复配置...")
        temp_uploads = None
        try:
            uploads_path = os.path.join(self.DATA_DIR, "uploads")
            if os.path.exists(uploads_path):
                temp_uploads = uploads_path + ".backup_temp"
                shutil.move(uploads_path, temp_uploads)

            # 解压备份到临时目录，再选择性覆盖
            temp_restore_dir = os.path.join(TEMP_DIR, "restore_temp")
            if os.path.exists(temp_restore_dir):
                shutil.rmtree(temp_restore_dir)
            os.makedirs(temp_restore_dir, exist_ok=True)
            with zipfile.ZipFile(zip_path, 'r') as zipf:
                zipf.extractall(temp_restore_dir)

            # 检查是否有 SQL 文件
            sql_files = [f for f in os.listdir(temp_restore_dir) if f.endswith('.sql')]
            if sql_files:
                self.append_result_text(f"ℹ️ 备份包中包含 SQL 文件：{', '.join(sql_files)}", "info")
                self.append_result_text("⚠️ 数据库恢复请使用 MySQL 客户端（如 phpMyAdmin）手动导入，本工具暂不支持自动导入", "warning")
                # 将 SQL 文件复制到 APP_DIR 方便用户手动恢复
                for sql in sql_files:
                    shutil.copy2(os.path.join(temp_restore_dir, sql), os.path.join(self.APP_DIR, sql))
                    self.append_result_text(f"📄 已将 SQL 文件复制到程序目录：{sql}", "info")

            # 恢复 data 目录（覆盖除 uploads 外的文件）
            if os.path.exists(os.path.join(temp_restore_dir, "data")):
                # 先删除现有 data 目录中除了 uploads 的所有内容
                for item in os.listdir(self.DATA_DIR):
                    if item != "uploads":
                        item_path = os.path.join(self.DATA_DIR, item)
                        if os.path.isfile(item_path):
                            os.remove(item_path)
                        elif os.path.isdir(item_path):
                            shutil.rmtree(item_path)
                # 复制备份中的 data 目录内容
                src_data = os.path.join(temp_restore_dir, "data")
                for item in os.listdir(src_data):
                    src_path = os.path.join(src_data, item)
                    dst_path = os.path.join(self.DATA_DIR, item)
                    if os.path.isfile(src_path):
                        shutil.copy2(src_path, dst_path)
                    else:
                        shutil.copytree(src_path, dst_path)

            if temp_uploads and os.path.exists(temp_uploads):
                shutil.move(temp_uploads, uploads_path)

            self.append_result_text(f"✅ 配置恢复成功：{zip_path}", "success")
            self.update_statusbar("恢复完成")
            self.refresh_service_status()
            # 成功弹窗提醒
            messagebox.showinfo("恢复成功", f"配置已成功从备份文件恢复：\n{zip_path}", parent=self.root)
        except Exception as e:
            log.error(f"恢复配置失败: {e}")
            if temp_uploads and os.path.exists(temp_uploads):
                try:
                    if os.path.exists(uploads_path):
                        shutil.rmtree(uploads_path)
                    shutil.move(temp_uploads, uploads_path)
                except:
                    pass
            messagebox.showerror("恢复失败", f"恢复配置时发生错误：{str(e)}")
            self.append_result_text(f"❌ 配置恢复失败：{str(e)}", "danger")
            self.update_statusbar("恢复失败")
        finally:
            # 清理临时目录
            if os.path.exists(temp_restore_dir):
                shutil.rmtree(temp_restore_dir, ignore_errors=True)

    def auto_backup_config(self):
        """自动备份配置（不含 uploads），如果数据库为 MySQL 则包含 SQL 导出。保存到程序目录"""
        if not os.path.exists(self.DATA_DIR):
            log.warning("自动备份失败：data 目录不存在")
            return
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        backup_filename = f"cloudreve_config_backup_{timestamp}.zip"
        backup_path = os.path.join(self.APP_DIR, backup_filename)
        try:
            # 创建临时目录
            temp_backup_dir = os.path.join(TEMP_DIR, "auto_backup_temp")
            if os.path.exists(temp_backup_dir):
                shutil.rmtree(temp_backup_dir)
            os.makedirs(temp_backup_dir, exist_ok=True)

            # 复制 data 目录（排除 uploads）
            dest_data = os.path.join(temp_backup_dir, "data")
            shutil.copytree(self.DATA_DIR, dest_data, ignore=shutil.ignore_patterns("uploads"))

            # 如果数据库是 MySQL，导出 SQL
            if os.path.exists(self.CONF_PATH):
                config = configparser.ConfigParser()
                config.read(self.CONF_PATH, encoding='utf-8')
                if 'Database' in config and config['Database'].get('Type', '').lower() == 'mysql':
                    self.append_result_text("📦 自动备份：检测到 MySQL 数据库，正在导出 SQL...", "info")
                    sql_temp = os.path.join(temp_backup_dir, f"mysql_dump_{config['Database'].get('Name', 'cloudreve')}.sql")
                    success, err = dump_mysql_database(self.CONF_PATH, sql_temp)
                    if success:
                        self.append_result_text("✅ 自动备份：MySQL 数据库导出成功", "success")
                    else:
                        self.append_result_text(f"⚠️ 自动备份：MySQL 导出失败：{err}", "warning")

            # 打包
            with zipfile.ZipFile(backup_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, dirs, files in os.walk(temp_backup_dir):
                    for file in files:
                        full_path = os.path.join(root, file)
                        arcname = os.path.relpath(full_path, start=temp_backup_dir)
                        zipf.write(full_path, arcname)

            self.append_result_text(f"📦 已自动备份配置至：{backup_path}", "success")
            log.info(f"自动备份成功：{backup_path}")
        except Exception as e:
            log.error(f"自动备份失败: {e}")
            self.append_result_text(f"⚠️ 自动备份失败：{e}", "warning")
        finally:
            if os.path.exists(temp_backup_dir):
                shutil.rmtree(temp_backup_dir, ignore_errors=True)

    def edit_config(self):
        if not os.path.exists(self.CONF_PATH):
            try:
                os.makedirs(self.DATA_DIR, exist_ok=True)
                with open(self.CONF_PATH, 'w', encoding='utf-8') as f:
                    f.write("# Cloudreve 配置文件\n")
                log.info(f"创建空配置文件：{self.CONF_PATH}")
            except Exception as e:
                log.error(f"创建配置文件失败: {e}")
                messagebox.showerror("错误", f"无法创建配置文件：{str(e)}")
                return
        try:
            os.startfile(self.CONF_PATH)
            self.append_result_text(f"📄 已打开配置文件：{self.CONF_PATH}", "info")
        except Exception as e:
            log.error(f"打开配置文件失败: {e}")
            messagebox.showerror("错误", f"无法打开配置文件：{str(e)}")

    def show_about(self):
        messagebox.showinfo(
            "关于 Cloudreve 管理工具",
            "Cloudreve 网盘服务管理工具\n版本 2.0\n\n功能：\n"
            "- 安装/卸载 Cloudreve 服务\n"
            "- 自动/手动升级\n"
            "- 服务状态监控\n"
            "- 端口检测与配置\n"
            "- 离开模式管理\n"
            "- 配置备份与恢复（含 MySQL 数据导出）\n"
            "- 配置文件编辑\n"
            "- 一键配置 MySQL 数据库\n\n"
            "使用 ttkbootstrap 主题"
        )

    # ========== 安装 MySQL 数据库（后台线程版） ==========
    def install_mysql_database(self):
        """利用现有 MySQL 服务配置 Cloudreve 数据库（后台线程版）"""
        # 禁用相关按钮，防止并发操作
        self.install_btn.config(state=DISABLED)
        self.uninstall_btn.config(state=DISABLED)
        self.open_btn.config(state=DISABLED)
        self.upgrade_btn.config(state=DISABLED)
        # 启动后台线程
        threading.Thread(target=self._install_mysql_worker, daemon=True).start()

    def _install_mysql_worker(self):
        """后台线程：执行 MySQL 配置的完整流程"""
        try:
            # 检查配置文件是否存在
            if not os.path.exists(self.CONF_PATH):
                self.root.after(0, lambda: self.append_result_text("❌ 未找到 Cloudreve 配置文件，请先安装网盘。", "danger"))
                self.root.after(0, lambda: messagebox.showerror("错误", "未找到 Cloudreve 配置文件，请先安装网盘。"))
                return

            # 检测 MySQL 服务
            self.root.after(0, lambda: self.append_result_text("🔍 检测 MySQL 服务...", "info"))
            mysql_service_name = None
            for name in ["MySQL80", "MySQL"]:
                if check_service_exists(name):
                    mysql_service_name = name
                    break

            if not mysql_service_name:
                self.root.after(0, lambda: self.append_result_text("❌ 未检测到 MySQL 服务。请先安装 MySQL 8.0（服务名 MySQL80 或 MySQL）并确保服务已启动。", "danger"))
                self.root.after(0, lambda: messagebox.showerror("错误", "未检测到 MySQL 服务。\n请安装 MySQL 8.0（服务名 MySQL80 或 MySQL）后重试。"))
                return

            self.root.after(0, lambda: self.append_result_text(f"✅ 检测到 MySQL 服务：{mysql_service_name}", "success"))

            # 获取 MySQL 安装路径
            try:
                result = subprocess.run(['sc', 'qc', mysql_service_name], capture_output=True, text=True, timeout=10, creationflags=0x08000000)
                if result.returncode == 0:
                    match = re.search(r'BINARY_PATH_NAME\s*:\s*(.+\.exe)', result.stdout, re.IGNORECASE)
                    if match:
                        mysqld_path = match.group(1).strip().strip('"')
                        mysql_bin_dir = os.path.dirname(mysqld_path)
                        mysql_client = os.path.join(mysql_bin_dir, "mysql.exe")
                        if not os.path.isfile(mysql_client):
                            raise Exception(f"未找到 mysql.exe，路径：{mysql_client}")
                        self.root.after(0, lambda: self.append_result_text(f"ℹ️ 找到 MySQL 客户端：{mysql_client}", "info"))
                    else:
                        raise Exception("无法解析服务二进制路径")
                else:
                    raise Exception("sc qc 命令失败")
            except Exception as e:
                self.root.after(0, lambda: self.append_result_text(f"❌ 获取 MySQL 路径失败：{e}", "danger"))
                self.root.after(0, lambda: messagebox.showerror("错误", f"无法获取 MySQL 安装路径，请确保 MySQL 服务已正确安装。\n错误：{e}"))
                return

            # 测试 root 连接（默认密码 cloudreve）
            self.root.after(0, lambda: self.append_result_text("⏳ 正在测试 MySQL 连接...", "info"))
            mysql_auth = None
            default_pwd = "cloudreve"
            try:
                test_cmd = [mysql_client, "-u", "root", f"-p{default_pwd}", "-e", "SELECT 1;"]
                result = subprocess.run(test_cmd, capture_output=True, timeout=10, creationflags=0x08000000)
                if result.returncode == 0:
                    mysql_auth = ["-u", "root", f"-p{default_pwd}"]
                    self.root.after(0, lambda: self.append_result_text("✅ 使用默认密码连接成功", "success"))
                else:
                    error_msg = result.stderr.decode('gbk', errors='ignore')
                    if "Access denied" in error_msg:
                        # 需要在主线程中弹窗输入密码
                        self.root.after(0, lambda: self._prompt_mysql_password(mysql_client, default_pwd))
                        return  # 密码输入会在回调中继续处理
                    else:
                        raise Exception(f"连接失败：{error_msg[:200]}")
            except Exception as e:
                self.root.after(0, lambda: self.append_result_text(f"❌ MySQL 连接测试失败：{e}", "danger"))
                self.root.after(0, lambda: messagebox.showerror("错误", f"无法连接 MySQL，请检查服务状态和 root 密码。\n错误：{e}"))
                return

            # 如果密码正确，继续执行配置
            if mysql_auth:
                self._continue_mysql_config(mysql_client, mysql_auth)

        except Exception as e:
            log.error(f"安装 MySQL 数据库失败: {e}")
            import traceback
            log.error(f"异常详情: {traceback.format_exc()}")
            self.root.after(0, lambda: self.append_result_text(f"❌ 安装 MySQL 数据库失败：{str(e)}", "danger"))
            self.root.after(0, lambda: messagebox.showerror("错误", f"安装 MySQL 数据库失败：{str(e)}"))
        finally:
            # 恢复按钮状态（如果中途失败，需要恢复）
            self.root.after(0, lambda: self.install_btn.config(state=NORMAL))
            self.root.after(0, lambda: self.uninstall_btn.config(state=NORMAL))
            self.root.after(0, lambda: self.open_btn.config(state=NORMAL))
            self.root.after(0, lambda: self.upgrade_btn.config(state=NORMAL))
            self.root.after(0, self.refresh_service_status)

    def _prompt_mysql_password(self, mysql_client, default_pwd):
        """主线程中弹出密码输入框，并继续配置"""
        root_pwd = simpledialog.askstring("MySQL root 密码",
                                           "默认密码 'cloudreve' 不正确，请输入正确的 MySQL root 密码：",
                                           show='*', parent=self.root)
        if root_pwd is None:
            # 用户取消，清理按钮状态
            self.install_btn.config(state=NORMAL)
            self.uninstall_btn.config(state=NORMAL)
            self.open_btn.config(state=NORMAL)
            self.upgrade_btn.config(state=NORMAL)
            return
        # 在新线程中继续测试密码
        threading.Thread(target=self._test_mysql_password, args=(mysql_client, root_pwd), daemon=True).start()

    def _test_mysql_password(self, mysql_client, root_pwd):
        """在后台测试输入的密码，成功后继续配置"""
        try:
            test_cmd = [mysql_client, "-u", "root", f"-p{root_pwd}", "-e", "SELECT 1;"]
            result = subprocess.run(test_cmd, capture_output=True, timeout=10, creationflags=0x08000000)
            if result.returncode == 0:
                mysql_auth = ["-u", "root", f"-p{root_pwd}"]
                self.root.after(0, lambda: self.append_result_text("✅ 手动输入密码连接成功", "success"))
                self._continue_mysql_config(mysql_client, mysql_auth)
            else:
                error_msg = result.stderr.decode('gbk', errors='ignore')
                raise Exception(f"连接失败：{error_msg[:200]}")
        except Exception as e:
            self.root.after(0, lambda: self.append_result_text(f"❌ MySQL 连接测试失败：{e}", "danger"))
            self.root.after(0, lambda: messagebox.showerror("错误", f"无法连接 MySQL，请检查 root 密码。\n错误：{e}"))
            # 恢复按钮状态
            self.root.after(0, lambda: self.install_btn.config(state=NORMAL))
            self.root.after(0, lambda: self.uninstall_btn.config(state=NORMAL))
            self.root.after(0, lambda: self.open_btn.config(state=NORMAL))
            self.root.after(0, lambda: self.upgrade_btn.config(state=NORMAL))
            self.root.after(0, self.refresh_service_status)

    # ---------- MySQL 凭据输入对话框 ----------
    def _ask_mysql_app_credentials(self, default_user="cloudreve", default_pass="your_password"):
        """在主线程中弹出对话框获取自定义用户名和密码，结果存入 self.mysql_app_user/pass 并设置事件"""
        try:
            # 获取用户名
            user = simpledialog.askstring(
                "设置 Cloudreve 数据库账户",
                f"请输入 Cloudreve 连接 MySQL 的用户名（长度 ≥4）\n默认：{default_user}",
                parent=self.root,
                initialvalue=default_user
            )
            if user is None:  # 用户取消
                user = default_user
                self._show_default_cred_warning()
            else:
                user = user.strip()
                while len(user) < 4:
                    messagebox.showwarning("用户名无效", "用户名长度至少为 4，请重新输入。")
                    user = simpledialog.askstring(
                        "设置 Cloudreve 数据库账户",
                        f"用户名长度至少 4 位，请重新输入\n默认：{default_user}",
                        parent=self.root,
                        initialvalue=user
                    )
                    if user is None:
                        user = default_user
                        self._show_default_cred_warning()
                        break
                    user = user.strip()

            # 获取密码
            pwd = simpledialog.askstring(
                "设置 Cloudreve 数据库密码",
                f"请输入 Cloudreve 连接 MySQL 的密码（长度 ≥4）\n默认：{default_pass}",
                parent=self.root,
                show='*',
                initialvalue=default_pass
            )
            if pwd is None:  # 用户取消
                pwd = default_pass
                self._show_default_cred_warning()
            else:
                pwd = pwd.strip()
                while len(pwd) < 4:
                    messagebox.showwarning("密码无效", "密码长度至少为 4，请重新输入。")
                    pwd = simpledialog.askstring(
                        "设置 Cloudreve 数据库密码",
                        f"密码长度至少 4 位，请重新输入\n默认：{default_pass}",
                        parent=self.root,
                        show='*',
                        initialvalue=pwd
                    )
                    if pwd is None:
                        pwd = default_pass
                        self._show_default_cred_warning()
                        break
                    pwd = pwd.strip()

            self.mysql_app_user = user
            self.mysql_app_pass = pwd
        finally:
            self.mysql_cred_event.set()

    def _show_default_cred_warning(self):
        """弹窗提示将使用默认凭据"""
        messagebox.showinfo("使用默认凭据", "您未输入自定义凭据，将使用默认用户名 'cloudreve' 和密码 'your_password'。")

    # ---------- MySQL 配置主流程（继续） ----------
    def _continue_mysql_config(self, mysql_client, mysql_auth):
        """在后台继续执行数据库和用户创建、配置修改、服务重启等操作"""
        try:
            # 设置进度起始点
            self.root.after(0, lambda: self.update_progress(0, "开始配置 MySQL..."))

            # 备份原有配置
            shutil.copy2(self.CONF_PATH, self.CONF_PATH + ".install_mysql_bak")
            self.root.after(0, lambda: self.append_result_text("📦 已备份原有配置文件", "info"))
            self.root.after(0, lambda: self.update_progress(10, "备份配置文件完成"))

            # ===== 获取用户自定义的 MySQL 用户名和密码 =====
            self.mysql_cred_event.clear()
            self.root.after(0, lambda: self._ask_mysql_app_credentials())
            waited = self.mysql_cred_event.wait(timeout=300)
            if not waited:
                self.root.after(0, lambda: self.append_result_text("⚠️ 等待用户输入超时，将使用默认凭据", "warning"))
                self.mysql_app_user = "cloudreve"
                self.mysql_app_pass = "your_password"
            else:
                self.root.after(0, lambda: self.append_result_text(f"✅ 使用自定义凭据：{self.mysql_app_user} / {self.mysql_app_pass}", "success"))
            # ============================================

            # 创建数据库和用户
            self.root.after(0, lambda: self.append_result_text("⏳ 正在创建数据库和用户...", "info"))
            self.root.after(0, lambda: self.update_progress(20, "创建数据库和用户..."))
            sql_commands = [
                "CREATE DATABASE IF NOT EXISTS cloudreve CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;",
                f"CREATE USER IF NOT EXISTS '{self.mysql_app_user}'@'localhost' IDENTIFIED BY '{self.mysql_app_pass}';",
                f"GRANT ALL PRIVILEGES ON cloudreve.* TO '{self.mysql_app_user}'@'localhost';",
                "FLUSH PRIVILEGES;"
            ]
            for cmd in sql_commands:
                try:
                    subprocess.run([mysql_client] + mysql_auth + ["-e", cmd], timeout=10, check=True, creationflags=0x08000000)
                except subprocess.CalledProcessError as e:
                    error_msg = e.stderr.decode('gbk', errors='ignore')
                    self.root.after(0, lambda: self.append_result_text(f"❌ 执行 SQL 失败：{error_msg[:200]}", "danger"))
                    self.root.after(0, lambda: messagebox.showerror("错误", f"创建数据库/用户失败：{error_msg[:200]}"))
                    return
            self.root.after(0, lambda: self.append_result_text("✅ 数据库和用户创建完成", "success"))
            self.root.after(0, lambda: self.update_progress(30, "数据库和用户创建完成"))

            # 修改 Cloudreve 配置文件
            self.root.after(0, lambda: self.append_result_text("⏳ 正在修改 Cloudreve 配置文件...", "info"))
            self.root.after(0, lambda: self.update_progress(40, "修改配置文件..."))
            config = configparser.ConfigParser()
            config.optionxform = str
            config.read(self.CONF_PATH, encoding='utf-8')
            if 'Database' not in config:
                config['Database'] = {}
            config['Database']['Type'] = 'mysql'
            config['Database']['Host'] = '127.0.0.1'
            config['Database']['Port'] = '3306'
            config['Database']['User'] = self.mysql_app_user
            config['Database']['Password'] = self.mysql_app_pass
            config['Database']['Name'] = 'cloudreve'
            config['Database']['Charset'] = 'utf8mb4'
            with open(self.CONF_PATH, 'w', encoding='utf-8') as f:
                config.write(f)
            self.root.after(0, lambda: self.append_result_text("✅ 配置文件已更新", "success"))
            self.root.after(0, lambda: self.update_progress(50, "配置文件已更新"))

            # 保存 MySQL 凭据到文件
            root_pwd = ""
            for arg in mysql_auth:
                if arg.startswith("-p"):
                    root_pwd = arg[2:]
                    break
            if not root_pwd:
                root_pwd = "unknown"
            cred_file = os.path.join(os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__), "mysql_credentials.txt")
            try:
                with open(cred_file, 'w', encoding='utf-8') as f:
                    f.write(f"MySQL root 密码：{root_pwd}\n")
                    f.write(f"Cloudreve 数据库用户名：{self.mysql_app_user}\n")
                    f.write(f"Cloudreve 数据库密码：{self.mysql_app_pass}\n")
                self.root.after(0, lambda: self.append_result_text(f"✅ MySQL 凭据已保存至：{cred_file}", "success"))
            except Exception as e:
                self.root.after(0, lambda: self.append_result_text(f"⚠️ 保存 MySQL 凭据失败：{e}", "warning"))

            # ===== 自动备份配置 =====
            self.root.after(0, lambda: self.append_result_text("📦 正在自动备份当前配置...", "info"))
            self.auto_backup_config()
            # =====================

            # 处理 Cloudreve 服务（启动/重启）
            self.root.after(0, lambda: self.append_result_text("⏳ 正在配置 Cloudreve 服务...", "info"))
            self.root.after(0, lambda: self.update_progress(55, "处理 Cloudreve 服务..."))
            service_status = get_service_status(self.SERVICE_NAME)
            process_exists = check_process_exists(self.PROCESS_NAME)
            service_running = (service_status == "RUNNING" and process_exists)

            if not service_running:
                self.root.after(0, lambda: self.append_result_text("ℹ️ 网盘服务未运行，正在启动...", "info"))
                start_service(self.SERVICE_NAME)
            else:
                self.root.after(0, lambda: self.append_result_text("ℹ️ 网盘服务已运行，正在重启以使配置生效...", "info"))
                stop_service(self.SERVICE_NAME)
                time.sleep(2)
                start_service(self.SERVICE_NAME)

            # 等待服务启动
            def refresh_cb():
                self.root.after(0, self.refresh_service_status)

            def progress_cb(r, t, e, total):
                progress = 55 + int(35 * r / t)
                self.root.after(0, self.update_progress, progress, f"等待服务启动... ({r}/{t})")

            service_started, status_msg, elapsed = wait_for_service_start(
                self.SERVICE_NAME,
                self.PROCESS_NAME,
                status_callback=refresh_cb,
                progress_callback=progress_cb
            )
            if not service_started:
                self.root.after(0, lambda: self.append_result_text(f"⚠️ 服务启动超时：{status_msg}", "warning"))
                self.root.after(0, lambda: messagebox.showwarning("启动失败", f"网盘服务启动失败：{status_msg}"))
                return
            else:
                self.root.after(0, lambda: self.append_result_text(f"✅ 服务启动成功（耗时{elapsed}秒）", "success"))

            # 等待端口开放
            current_port = get_port_from_conf(self.CONF_PATH, self.DEFAULT_PORT)
            try:
                wait_for_port_open('localhost', current_port, max_wait=30)
                self.root.after(0, lambda: self.append_result_text(f"✅ 端口{current_port}已开放", "success"))
            except TimeoutError:
                self.root.after(0, lambda: self.append_result_text(f"⚠️ 端口{current_port}未开放（服务可能仍在启动）", "warning"))

            self.root.after(0, lambda: self.update_progress(95, "准备打开网盘..."))
            url = f"http://localhost:{current_port}"

            def ask_open():
                result = messagebox.askokcancel(
                    "MySQL 配置完成",
                    f"MySQL 数据库已成功配置！\n网盘访问地址：{url}\n\n是否立即打开网盘页面？",
                    parent=self.root
                )
                if result:
                    webbrowser.open(url)
                    self.root.after(0, lambda: self.append_result_text(f"✅ 已打开网盘页面：{url}", "success"))
                else:
                    self.root.after(0, lambda: self.append_result_text("ℹ️ 用户取消打开网页", "info"))
                self.root.after(0, lambda: self.update_progress(100, "配置完成"))
                self.root.after(0, lambda: self.update_statusbar("MySQL 配置完成", bootstyle="success"))
            self.root.after(0, ask_open)

        except Exception as e:
            log.error(f"继续配置 MySQL 失败: {e}")
            import traceback
            log.error(f"异常详情: {traceback.format_exc()}")
            self.root.after(0, lambda: self.append_result_text(f"❌ 配置 MySQL 失败：{str(e)}", "danger"))
            self.root.after(0, lambda: messagebox.showerror("错误", f"配置 MySQL 失败：{str(e)}"))
            self.root.after(0, lambda: self.update_progress(100, "配置失败"))

    # ========== 打开网盘 ==========
    def start_open_cloudreve(self):
        self.clear_result_text()
        self.append_result_text("🌐 开始打开网盘...", "info")
        self.update_statusbar("正在打开网盘...")
        open_thread = threading.Thread(target=self.open_cloudreve_worker, daemon=True)
        open_thread.start()

    def open_cloudreve_worker(self):
        try:
            self.open_btn.config(state=DISABLED)
            self.install_btn.config(state=DISABLED)
            self.uninstall_btn.config(state=DISABLED)
            self.upgrade_btn.config(state=DISABLED)

            self.root.after(0, lambda: self.update_progress(10, "检查服务是否存在..."))
            if not check_service_exists(self.SERVICE_NAME):
                self.root.after(0, lambda: self.update_progress(100, "服务未安装"))
                self.root.after(0, lambda: self.append_result_text("❌ 网盘服务未安装，请先执行安装。", "danger"))
                self.root.after(0, lambda: messagebox.showwarning("无法打开", "网盘服务未安装，请先执行安装。"))
                self.open_btn.config(state=NORMAL)
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                self.upgrade_btn.config(state=NORMAL)
                return

            self.root.after(0, lambda: self.update_progress(20, "读取配置端口..."))
            current_port = get_port_from_conf(self.CONF_PATH, self.DEFAULT_PORT)
            self.root.after(0, lambda: self.append_result_text(f"ℹ️ 配置文件端口：{current_port}", "info"))

            service_status = get_service_status(self.SERVICE_NAME)
            if service_status != "RUNNING":
                self.root.after(0, lambda: self.update_progress(30, "服务未运行，正在启动..."))
                self.root.after(0, lambda: self.append_result_text("⏳ 服务未运行，正在启动...", "info"))
                start_service(self.SERVICE_NAME)
                self.root.after(0, lambda: self.append_result_text("✅ 已发送启动指令", "success"))

            self.root.after(0, lambda: self.update_progress(50, "等待服务启动..."))
            self.root.after(0, lambda: self.append_result_text("⏳ 正在检测服务状态...", "info"))
            service_started, status_msg, elapsed = wait_for_service_start(
                self.SERVICE_NAME,
                self.PROCESS_NAME,
                status_callback=lambda: self.root.after(0, self.refresh_service_status),
                progress_callback=lambda r, t, e, total: self.root.after(0, self.update_progress, 50 + int(30 * r / t), f"等待服务启动... ({r}/{t})")
            )
            if not service_started:
                self.root.after(0, lambda: self.update_progress(100, "服务启动失败"))
                self.root.after(0, lambda: self.append_result_text(f"⚠️ 服务启动超时：{status_msg}", "warning"))
                self.root.after(0, lambda: messagebox.showwarning("启动失败", f"网盘服务启动失败：{status_msg}"))
                self.open_btn.config(state=NORMAL)
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                self.upgrade_btn.config(state=NORMAL)
                return
            else:
                self.root.after(0, lambda: self.append_result_text(f"✅ 服务启动成功（耗时{elapsed}秒）", "success"))

            self.root.after(0, lambda: self.update_progress(80, "等待端口开放..."))
            self.root.after(0, lambda: self.append_result_text("⏳ 正在等待端口开放...", "info"))
            try:
                def port_progress(elapsed, max_wait, remaining):
                    self.root.after(0, self.update_progress, 80 + int(20 * (elapsed / max_wait)), f"等待端口开放... ({elapsed:.1f}/{max_wait}秒)")
                wait_for_port_open('localhost', current_port, max_wait=30, progress_callback=port_progress)
                self.root.after(0, lambda: self.append_result_text(f"✅ 端口{current_port}已开放", "success"))
            except TimeoutError:
                self.root.after(0, lambda: self.update_progress(100, "端口开放超时"))
                self.root.after(0, lambda: self.append_result_text(f"⚠️ 端口{current_port}未开放（服务可能仍在启动）", "warning"))
                self.root.after(0, lambda: messagebox.showwarning("超时", f"端口{current_port}未及时开放，请手动检查服务状态。"))
                self.open_btn.config(state=NORMAL)
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                self.upgrade_btn.config(state=NORMAL)
                return

            # 打开网盘页面（直接打开，无确认对话框）
            self.root.after(0, lambda: self.update_progress(95, "打开网盘页面..."))
            url = f"http://localhost:{current_port}"
            try:
                webbrowser.open(url)
                self.root.after(0, lambda: self.append_result_text(f"✅ 已打开网盘页面：{url}", "success"))
                log.info(f"打开网盘成功：{url}")
            except Exception as e:
                self.root.after(0, lambda: self.append_result_text(f"⚠️ 打开页面失败，请手动访问：{url}，错误：{e}", "warning"))
                self.root.after(0, lambda: messagebox.showwarning("打开失败", f"无法自动打开网页，请手动访问：{url}"))

            self.root.after(0, lambda: self.update_progress(100, "操作完成"))
            self.root.after(0, lambda: self.append_result_text("🎉 网盘打开流程完成", "success"))
            self.root.after(0, lambda: self.update_statusbar("网盘已打开"))

        except Exception as e:
            log.error(f"打开网盘异常: {e}")
            import traceback
            log.error(f"异常详情: {traceback.format_exc()}")
            self.root.after(0, lambda: self.append_result_text(f"❌ 打开网盘过程发生异常：{str(e)}", "danger"))
            self.root.after(0, lambda: self.update_progress(100, "异常"))
            self.root.after(0, lambda: messagebox.showerror("错误", f"打开网盘时发生错误：{str(e)}"))
            self.root.after(0, lambda: self.update_statusbar("打开失败"))
        finally:
            self.open_btn.config(state=NORMAL)
            self.install_btn.config(state=NORMAL)
            self.uninstall_btn.config(state=NORMAL)
            self.upgrade_btn.config(state=NORMAL)
            self.refresh_service_status()

    # ========== 端口检测 ==========
    def check_port_status(self):
        try:
            self.clear_result_text()
            self.append_result_text("📌 开始检测端口状态...", "info")
            self.update_statusbar("检测端口...")
            is_valid, port = self.validate_port()
            if not is_valid:
                messagebox.showerror("错误", port)
                self.append_result_text(f"❌ 端口验证失败：{port}", "danger")
                return
            self.update_progress(0, "正在检测端口状态...")
            conf_port = get_port_from_conf(self.CONF_PATH, self.DEFAULT_PORT)
            port_open = is_port_open('localhost', port)
            is_used = is_port_occupied(port)

            report = f"""📝 端口检测报告
━━━━━━━━━━━━━━━━━━━━
检测端口：{port} | 配置文件端口：{conf_port}
端口开放：{'✅' if port_open else '❌'} | 端口占用：{'⚠️' if is_used else '✅'}
━━━━━━━━━━━━━━━━━━━━"""

            can_install = False
            if is_used:
                proc = get_process_using_port(port)
                if proc == "cloudreve.exe":
                    can_install = True
                    report += f"""
ℹ️ 端口{port}已被 {proc} 占用（当前正在运行的网盘服务）"""
                else:
                    can_install = False
                    report += f"""
⚠️ 端口{port}已被占用（占用程序：{proc if proc else '未知'}）"""
            else:
                can_install = True
                if port != conf_port:
                    report += f"""
ℹ️ 端口不一致：当前{port} ≠ 配置{conf_port}
安装时将自动同步为：Listen = :{port}"""
                else:
                    report += f"""
✅ 端口{port}可用"""

            if can_install:
                report += """
✅ 可以安装网盘"""
                self.update_progress(100, "端口检测完成（可安装）")
                self.append_result_text(report, "success")
            else:
                report += """
❌ 不能安装网盘（请更换端口或关闭占用程序）"""
                self.update_progress(100, "端口检测完成（不可安装）")
                self.append_result_text(report, "danger")
            self.update_statusbar("端口检测完成")
        except Exception as e:
            log.error(f"检测端口异常: {e}")
            self.append_result_text(f"❌ 端口检测失败：{str(e)}", "danger")
            self.update_statusbar("检测失败")

    # ========== 安装模块 ==========
    def start_install(self):
        self.clear_result_text()
        self.append_result_text("🚀 开始执行安装流程...", "info")
        self.update_statusbar("正在安装...")
        install_thread = threading.Thread(target=self.install_worker, daemon=True)
        install_thread.start()

    def install_worker(self):
        try:
            self.install_btn.config(state=DISABLED)
            self.uninstall_btn.config(state=DISABLED)
            self.open_btn.config(state=DISABLED)
            self.upgrade_btn.config(state=DISABLED)
            
            self.root.after(0, lambda: self.update_progress(5, "验证端口输入..."))
            is_port_valid, current_port = self.validate_port()
            if not is_port_valid:
                self.root.after(0, lambda: self.update_progress(100, "端口配置错误"))
                self.root.after(0, lambda: self.append_result_text(f"""❌ 安装失败！
端口验证失败：{current_port}
请输入1-65535之间的有效端口""", "danger"))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                self.open_btn.config(state=NORMAL)
                self.upgrade_btn.config(state=NORMAL)
                return
            
            self.root.after(0, lambda: self.update_progress(8, "检查端口占用..."))
            if is_port_occupied(current_port):
                proc = get_process_using_port(current_port)
                if proc == "cloudreve.exe":
                    self.root.after(0, lambda: self.append_result_text(f"ℹ️ 端口{current_port}已被 {proc} 占用，将停止现有服务并重新安装", "info"))
                else:
                    self.root.after(0, lambda: self.update_progress(100, "端口被占用"))
                    self.root.after(0, lambda: self.append_result_text(f"""❌ 安装失败！
端口 {current_port} 已被占用（占用程序：{proc if proc else '未知'}）
请更换端口后重试""", "danger"))
                    self.install_btn.config(state=NORMAL)
                    self.uninstall_btn.config(state=NORMAL)
                    self.open_btn.config(state=NORMAL)
                    self.upgrade_btn.config(state=NORMAL)
                    return
            
            self.root.after(0, lambda: self.update_progress(10, "检查管理员权限..."))
            if not check_admin_permission():
                self.root.after(0, lambda: self.update_progress(100, "权限不足"))
                self.root.after(0, lambda: self.append_result_text("""❌ 安装失败！
需要管理员权限，请右键以管理员身份运行""", "danger"))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                self.open_btn.config(state=NORMAL)
                self.upgrade_btn.config(state=NORMAL)
                return
            
            self.root.after(0, lambda: self.update_progress(15, "检查必要文件..."))
            if not os.path.exists(self.WINSW_EXE):
                self.root.after(0, lambda: self.update_progress(100, "文件缺失"))
                self.root.after(0, lambda: self.append_result_text(f"""❌ 安装失败！
未找到 winsw.exe
路径：{self.WINSW_EXE}""", "danger"))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                self.open_btn.config(state=NORMAL)
                self.upgrade_btn.config(state=NORMAL)
                return
            if not os.path.exists(self.WINSW_XML):
                self.root.after(0, lambda: self.update_progress(100, "配置文件缺失"))
                self.root.after(0, lambda: self.append_result_text(f"""❌ 安装失败！
未找到 winsw.xml
路径：{self.WINSW_XML}""", "danger"))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                self.open_btn.config(state=NORMAL)
                self.upgrade_btn.config(state=NORMAL)
                return
            
            self.root.after(0, lambda: self.update_progress(20, f"配置防火墙（端口{current_port}）..."))
            firewall_success, firewall_error = add_firewall_rule(current_port)
            if not firewall_success:
                self.root.after(0, lambda: self.append_result_text(f"⚠️ 防火墙配置警告：{firewall_error}", "warning"))
            
            self.root.after(0, lambda: self.update_progress(25, "停止现有服务..."))
            if check_service_exists(self.SERVICE_NAME):
                stop_service(self.SERVICE_NAME)
                self.root.after(0, lambda: self.append_result_text("✅ 已停止现有服务", "success"))
            else:
                self.root.after(0, lambda: self.append_result_text("ℹ️ 未检测到现有服务", "info"))
            
            self.root.after(0, lambda: self.update_progress(35, "结束现有进程..."))
            if check_process_exists(self.PROCESS_NAME):
                kill_process(self.PROCESS_NAME)
                self.root.after(0, lambda: self.append_result_text("✅ 已结束现有进程", "success"))
            else:
                self.root.after(0, lambda: self.append_result_text("ℹ️ 未检测到现有进程", "info"))
            
            self.root.after(0, lambda: self.update_progress(40, "安装服务..."))
            run_cmd_with_retry([self.WINSW_EXE, "install"], timeout=30)
            self.root.after(0, lambda: self.append_result_text("✅ 服务安装完成", "success"))
            
            self.root.after(0, lambda: self.update_progress(45, "启动服务..."))
            start_service(self.SERVICE_NAME)
            self.root.after(0, lambda: self.append_result_text("✅ 尝试启动服务", "success"))
            
            self.root.after(0, lambda: self.update_progress(50, "等待服务启动..."))
            self.root.after(0, lambda: self.append_result_text("⏳ 正在检测服务状态...", "info"))
            service_started, status_msg, elapsed = wait_for_service_start(
                self.SERVICE_NAME,
                self.PROCESS_NAME,
                status_callback=lambda: self.root.after(0, self.refresh_service_status),
                progress_callback=lambda r, t, e, total: self.root.after(0, self.update_progress, 50 + int(30 * r / t), f"等待服务启动... ({r}/{t})")
            )
            if not service_started:
                self.root.after(0, lambda: self.update_progress(100, "服务启动失败"))
                self.root.after(0, lambda: self.append_result_text(f"""⚠️ 安装不完整！
服务启动超时（{elapsed}秒）
状态：{status_msg}
建议：手动启动服务或检查配置""", "warning"))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                self.open_btn.config(state=NORMAL)
                self.upgrade_btn.config(state=NORMAL)
                return
            else:
                self.root.after(0, lambda: self.append_result_text(f"✅ 服务启动成功（耗时{elapsed}秒）", "success"))
            
            self.root.after(0, lambda: self.update_progress(55, "等待配置文件..."))
            def conf_progress(elapsed, max_wait, remaining):
                self.root.after(0, self.update_progress, 55 + int(25 * (elapsed / max_wait)), f"等待配置文件生成... ({elapsed:.1f}/{max_wait}秒)")
            conf_generated = wait_for_conf_file(self.CONF_PATH, timeout=120, progress_callback=conf_progress)
            if not conf_generated:
                self.root.after(0, lambda: self.update_progress(100, "配置文件生成超时"))
                self.root.after(0, lambda: self.append_result_text(f"""⚠️ 配置文件生成失败！
路径：{self.CONF_PATH}
建议：手动创建配置文件或检查服务是否正常""", "warning"))
            else:
                self.root.after(0, lambda: self.update_progress(60, "同步端口配置（并备份原文件）..."))
                modify_success = modify_conf_port(self.CONF_PATH, current_port)
                if modify_success:
                    self.root.after(0, lambda: self.append_result_text(f"✅ 配置文件端口已同步为 {current_port}（原文件已备份为 .bak）", "success"))
                else:
                    self.root.after(0, lambda: self.append_result_text(f"⚠️ 配置文件端口同步失败（需手动修改）", "warning"))
                
                self.root.after(0, lambda: self.update_progress(70, "重启服务应用配置..."))
                stop_service(self.SERVICE_NAME)
                time.sleep(3)
                start_service(self.SERVICE_NAME)
                self.root.after(0, lambda: self.append_result_text("✅ 服务已重启，配置生效", "success"))
                
                self.root.after(0, lambda: self.update_progress(80, "等待端口开放..."))
                def port_progress(elapsed, max_wait, remaining):
                    self.root.after(0, self.update_progress, 80 + int(15 * (elapsed / max_wait)), f"等待端口开放... ({elapsed:.1f}/{max_wait}秒)")
                try:
                    wait_for_port_open('localhost', current_port, max_wait=30, progress_callback=port_progress)
                    self.root.after(0, lambda: self.append_result_text(f"✅ 端口{current_port}已开放", "success"))
                except TimeoutError:
                    self.root.after(0, lambda: self.append_result_text(f"⚠️ 端口{current_port}未开放（服务可能仍在启动）", "warning"))
                
                # 弹出确认对话框，用户选择是否打开网页
                self.root.after(0, lambda: self.update_progress(90, "准备打开网盘..."))
                url = f"http://localhost:{current_port}"
                def ask_open():
                    result = messagebox.askokcancel(
                        "安装成功",
                        f"网盘安装成功！\n访问地址：{url}\n\n是否立即打开网盘页面？",
                        parent=self.root
                    )
                    if result:
                        webbrowser.open(url)
                        self.root.after(0, lambda: self.append_result_text(f"✅ 已打开网盘页面：{url}", "success"))
                    else:
                        self.root.after(0, lambda: self.append_result_text("ℹ️ 用户取消打开网页", "info"))
                    self.root.after(0, lambda: self.update_statusbar("安装完成"))
                self.root.after(0, ask_open)
                
                self.root.after(0, lambda: self.update_progress(95, "启用离开模式..."))
                if enable_away_mode():
                    self.root.after(0, lambda: self.append_result_text("✅ 离开模式已启用（系统睡眠时保持网络连接）", "success"))
                else:
                    self.root.after(0, lambda: self.append_result_text("⚠️ 启用离开模式失败，请手动检查注册表权限", "warning"))
                
                self.root.after(0, lambda: self.append_result_text(f"""🎉 安装流程完成！
网盘访问地址：{url}
服务名称：{self.SERVICE_NAME}
进程名称：{self.PROCESS_NAME}""", "success"))
        except Exception as e:
            log.error(f"安装异常: {e}")
            import traceback
            log.error(f"异常详情: {traceback.format_exc()}")
            self.root.after(0, lambda: self.append_result_text(f"""❌ 安装过程发生异常：
{str(e)}
请查看日志文件获取详细信息""", "danger"))
            self.root.after(0, lambda: self.update_progress(100, "安装异常"))
            self.root.after(0, lambda: self.update_statusbar("安装失败"))
        finally:
            self.install_btn.config(state=NORMAL)
            self.uninstall_btn.config(state=NORMAL)
            self.open_btn.config(state=NORMAL)
            self.upgrade_btn.config(state=NORMAL)
            self.refresh_service_status()

    # ========== 卸载模块 ==========
    def start_uninstall(self):
        self.clear_result_text()
        self.append_result_text("🗑️ 开始执行卸载流程...", "info")
        self.update_statusbar("正在卸载...")
        uninstall_thread = threading.Thread(target=self.uninstall_worker, daemon=True)
        uninstall_thread.start()

    def uninstall_worker(self):
        try:
            self.install_btn.config(state=DISABLED)
            self.uninstall_btn.config(state=DISABLED)
            self.open_btn.config(state=DISABLED)
            self.upgrade_btn.config(state=DISABLED)
            
            self.root.after(0, lambda: self.update_progress(10, "检查管理员权限..."))
            if not check_admin_permission():
                self.root.after(0, lambda: self.update_progress(100, "权限不足"))
                self.root.after(0, lambda: self.append_result_text("""❌ 卸载失败！
需要管理员权限，请右键以管理员身份运行""", "danger"))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                self.open_btn.config(state=NORMAL)
                self.upgrade_btn.config(state=NORMAL)
                return
            
            self.root.after(0, lambda: self.update_progress(20, "停止服务..."))
            if check_service_exists(self.SERVICE_NAME):
                stop_service(self.SERVICE_NAME)
                self.root.after(0, lambda: self.append_result_text("✅ 已停止服务", "success"))
            else:
                self.root.after(0, lambda: self.append_result_text("ℹ️ 未检测到服务", "info"))
            
            self.root.after(0, lambda: self.update_progress(30, "结束进程..."))
            if check_process_exists(self.PROCESS_NAME):
                kill_process(self.PROCESS_NAME)
                self.root.after(0, lambda: self.append_result_text("✅ 已结束进程", "success"))
            else:
                self.root.after(0, lambda: self.append_result_text("ℹ️ 未检测到进程", "info"))
            
            self.root.after(0, lambda: self.update_progress(40, "卸载服务..."))
            if os.path.exists(self.WINSW_EXE):
                run_cmd_with_retry([self.WINSW_EXE, "uninstall"], timeout=30)
                self.root.after(0, lambda: self.append_result_text("✅ 服务已卸载", "success"))
            else:
                self.root.after(0, lambda: self.append_result_text("⚠️ winsw.exe不存在，跳过服务卸载", "warning"))
            
            self.root.after(0, lambda: self.update_progress(50, "清理防火墙规则..."))
            current_port = get_port_from_conf(self.CONF_PATH, self.DEFAULT_PORT)
            try:
                run_cmd_with_retry(['cmd', '/c', f'netsh advfirewall firewall delete rule name="Cloudreve_Port_Inbound"'], timeout=10)
                run_cmd_with_retry(['cmd', '/c', f'netsh advfirewall firewall delete rule name="Cloudreve_Port_Outbound"'], timeout=10)
                self.root.after(0, lambda: self.append_result_text("✅ 防火墙规则已清理", "success"))
            except Exception as e:
                self.root.after(0, lambda: self.append_result_text(f"⚠️ 防火墙清理失败：{e}，请手动检查", "warning"))
            
            self.root.after(0, lambda: self.update_progress(60, "禁用离开模式..."))
            if disable_away_mode():
                self.root.after(0, lambda: self.append_result_text("✅ 离开模式已禁用", "success"))
            else:
                self.root.after(0, lambda: self.append_result_text("⚠️ 禁用离开模式失败，请手动检查注册表", "warning"))
            
            self.root.after(0, lambda: self.update_progress(100, "卸载完成"))
            self.root.after(0, lambda: self.append_result_text("""🎉 卸载流程完成！
已执行操作：
1. 停止并卸载系统服务
2. 结束相关进程
3. 清理防火墙规则
4. 禁用离开模式（恢复系统正常睡眠）

注意：data目录（含配置/数据文件）未删除，如需彻底清理请手动删除。""", "success"))
            self.root.after(0, lambda: self.update_statusbar("卸载完成"))
        except Exception as e:
            log.error(f"卸载异常: {e}")
            import traceback
            log.error(f"异常详情: {traceback.format_exc()}")
            self.root.after(0, lambda: self.append_result_text(f"""❌ 卸载过程发生异常：
{str(e)}
请查看日志文件获取详细信息""", "danger"))
            self.root.after(0, lambda: self.update_progress(100, "卸载异常"))
            self.root.after(0, lambda: self.update_statusbar("卸载失败"))
        finally:
            self.install_btn.config(state=NORMAL)
            self.uninstall_btn.config(state=NORMAL)
            self.open_btn.config(state=NORMAL)
            self.upgrade_btn.config(state=NORMAL)
            self.refresh_service_status()

    # ========== 升级模块 ==========
    def start_upgrade(self):
        choice = messagebox.askquestion(
            "升级方式选择",
            "请选择升级方式：\n\n点击“是”进行自动升级（从GitHub下载最新版本）\n点击“否”手动选择文件升级",
            icon='question'
        )
        if choice == 'yes':
            self.clear_result_text()
            self.append_result_text("🚀 开始自动升级流程...", "info")
            self.update_statusbar("自动升级中...")
            upgrade_thread = threading.Thread(target=self.auto_upgrade_worker, daemon=True)
            upgrade_thread.start()
        else:
            self.clear_result_text()
            self.append_result_text("🔧 开始手动升级流程...", "info")
            self.update_statusbar("手动升级中...")
            upgrade_thread = threading.Thread(target=self.manual_upgrade_worker, daemon=True)
            upgrade_thread.start()

    def auto_upgrade_worker(self):
        try:
            self.install_btn.config(state=DISABLED)
            self.uninstall_btn.config(state=DISABLED)
            self.open_btn.config(state=DISABLED)
            self.upgrade_btn.config(state=DISABLED)

            self.root.after(0, lambda: self.update_progress(5, "检查当前版本..."))
            current_version = get_current_version(self.CLOUDREVE_EXE)
            if current_version:
                self.root.after(0, lambda: self.append_result_text(f"ℹ️ 当前版本：{current_version}", "info"))
            else:
                self.root.after(0, lambda: self.append_result_text("⚠️ 无法获取当前版本，将尝试下载最新版", "warning"))

            self.root.after(0, lambda: self.update_progress(10, "从GitHub获取最新版本..."))
            try:
                latest_version = get_latest_version_from_github()
                self.root.after(0, lambda: self.append_result_text(f"📦 GitHub最新版本：{latest_version}", "info"))
            except Exception as e:
                self.root.after(0, lambda: self.update_progress(100, "网络错误"))
                self.root.after(0, lambda: self.append_result_text(f"❌ 自动升级失败：{str(e)}", "danger"))
                self.root.after(0, lambda: messagebox.showerror("网络错误", f"无法连接到GitHub，请检查网络后重试。\n\n错误信息：{str(e)}"))
                self.upgrade_btn.config(state=NORMAL)
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                self.open_btn.config(state=NORMAL)
                return

            if current_version and not compare_versions(current_version, latest_version):
                self.root.after(0, lambda: self.update_progress(100, "已是最新版本"))
                self.root.after(0, lambda: self.append_result_text("✅ 当前已是最新版本，无需升级", "success"))
                self.root.after(0, lambda: messagebox.showinfo("无需升级", f"当前版本 {current_version} 已是最新版本。"))
                self.upgrade_btn.config(state=NORMAL)
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                self.open_btn.config(state=NORMAL)
                return

            confirm = messagebox.askyesno(
                "确认升级",
                f"发现新版本 {latest_version}，当前版本 {current_version or '未知'}。\n是否立即升级？"
            )
            if not confirm:
                self.root.after(0, lambda: self.update_progress(100, "用户取消升级"))
                self.root.after(0, lambda: self.append_result_text("⚠️ 升级已取消", "warning"))
                self.upgrade_btn.config(state=NORMAL)
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                self.open_btn.config(state=NORMAL)
                return

            self.root.after(0, lambda: self.update_progress(20, "获取下载链接..."))
            try:
                download_url = get_download_url_from_github()
                if not download_url:
                    raise RuntimeError("未找到适用于Windows的下载链接")
                self.root.after(0, lambda: self.append_result_text(f"🔗 下载链接已获取", "info"))
            except Exception as e:
                self.root.after(0, lambda: self.update_progress(100, "获取链接失败"))
                self.root.after(0, lambda: self.append_result_text(f"❌ 获取下载链接失败：{str(e)}", "danger"))
                self.root.after(0, lambda: messagebox.showerror("错误", f"无法获取下载链接，请稍后重试或使用手动升级。"))
                self.upgrade_btn.config(state=NORMAL)
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                self.open_btn.config(state=NORMAL)
                return

            self.root.after(0, lambda: self.update_progress(25, "开始下载新版本..."))
            os.makedirs(TEMP_DIR, exist_ok=True)
            zip_path = os.path.join(TEMP_DIR, "cloudreve_latest.zip")
            try:
                def download_progress(percent, downloaded, total):
                    self.root.after(0, self.update_progress, 25 + int(35 * percent / 100),
                                    f"下载中... {percent}% ({downloaded}/{total} 字节)")
                download_file(download_url, zip_path, download_progress)
                self.root.after(0, lambda: self.append_result_text("✅ 下载完成", "success"))
            except Exception as e:
                self.root.after(0, lambda: self.update_progress(100, "下载失败"))
                self.root.after(0, lambda: self.append_result_text(f"❌ 下载失败：{str(e)}", "danger"))
                self.root.after(0, lambda: messagebox.showerror("下载失败", f"文件下载失败，请检查网络后重试。\n\n错误信息：{str(e)}"))
                self.upgrade_btn.config(state=NORMAL)
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                self.open_btn.config(state=NORMAL)
                return

            self.root.after(0, lambda: self.update_progress(60, "解压升级包..."))
            extract_dir = os.path.join(TEMP_DIR, "extract")
            try:
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    zip_ref.extractall(extract_dir)
                self.root.after(0, lambda: self.append_result_text("✅ 解压完成", "success"))
            except Exception as e:
                self.root.after(0, lambda: self.update_progress(100, "解压失败"))
                self.root.after(0, lambda: self.append_result_text(f"❌ 解压失败：{str(e)}", "danger"))
                self.root.after(0, lambda: messagebox.showerror("解压失败", f"无法解压下载的ZIP文件，请检查文件完整性。"))
                self.upgrade_btn.config(state=NORMAL)
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                self.open_btn.config(state=NORMAL)
                return

            new_exe_path = None
            for root_dir, dirs, files in os.walk(extract_dir):
                for file in files:
                    if file.lower() == "cloudreve.exe":
                        new_exe_path = os.path.join(root_dir, file)
                        break
                if new_exe_path:
                    break
            if not new_exe_path:
                self.root.after(0, lambda: self.update_progress(100, "文件缺失"))
                self.root.after(0, lambda: self.append_result_text("❌ 解压包中未找到 cloudreve.exe", "danger"))
                self.root.after(0, lambda: messagebox.showerror("文件缺失", "下载的ZIP包中未找到 cloudreve.exe，请尝试手动升级。"))
                self.upgrade_btn.config(state=NORMAL)
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                self.open_btn.config(state=NORMAL)
                return

            service_exists_before = check_service_exists(self.SERVICE_NAME)
            self.root.after(0, lambda: self.append_result_text(f"ℹ️ 升级前服务状态：{'存在' if service_exists_before else '不存在'}", "info"))

            if service_exists_before:
                self.root.after(0, lambda: self.update_progress(65, "停止服务..."))
                if check_service_exists(self.SERVICE_NAME):
                    stop_service(self.SERVICE_NAME)
                    self.root.after(0, lambda: self.append_result_text("✅ 已停止服务", "success"))
                else:
                    self.root.after(0, lambda: self.append_result_text("ℹ️ 服务未运行", "info"))

                self.root.after(0, lambda: self.update_progress(70, "结束进程..."))
                if check_process_exists(self.PROCESS_NAME):
                    kill_process(self.PROCESS_NAME)
                    self.root.after(0, lambda: self.append_result_text("✅ 已结束进程", "success"))
                else:
                    self.root.after(0, lambda: self.append_result_text("ℹ️ 未检测到进程", "info"))
            else:
                self.root.after(0, lambda: self.update_progress(65, "检查残留进程..."))
                if check_process_exists(self.PROCESS_NAME):
                    kill_process(self.PROCESS_NAME)
                    self.root.after(0, lambda: self.append_result_text("✅ 已结束残留进程", "success"))
                else:
                    self.root.after(0, lambda: self.append_result_text("ℹ️ 未检测到进程", "info"))

            self.root.after(0, lambda: self.update_progress(75, "备份旧版本..."))
            backup_path = self.CLOUDREVE_EXE + ".bak"
            if os.path.exists(self.CLOUDREVE_EXE):
                try:
                    shutil.copy2(self.CLOUDREVE_EXE, backup_path)
                    self.root.after(0, lambda: self.append_result_text(f"✅ 已备份原文件至 {backup_path}", "success"))
                except Exception as e:
                    self.root.after(0, lambda: self.append_result_text(f"⚠️ 备份失败：{e}，继续升级", "warning"))

            self.root.after(0, lambda: self.update_progress(80, "替换新版本..."))
            try:
                shutil.copy2(new_exe_path, self.CLOUDREVE_EXE)
                self.root.after(0, lambda: self.append_result_text(f"✅ 新版本已复制到 {self.CLOUDREVE_EXE}", "success"))
                if os.path.exists(self.CLOUDREVE_EXE):
                    mod_time = os.path.getmtime(self.CLOUDREVE_EXE)
                    mod_time_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mod_time))
                    self.root.after(0, lambda: self.append_result_text(f"📅 新版本文件修改日期：{mod_time_str}", "info"))
            except Exception as e:
                self.root.after(0, lambda: self.update_progress(100, "替换失败"))
                self.root.after(0, lambda: self.append_result_text(f"❌ 替换文件失败：{e}", "danger"))
                self.root.after(0, lambda: messagebox.showerror("替换失败", f"无法替换文件，请检查权限或手动升级。"))
                self.upgrade_btn.config(state=NORMAL)
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                self.open_btn.config(state=NORMAL)
                return

            if service_exists_before:
                self.root.after(0, lambda: self.update_progress(85, "启动服务..."))
                start_service(self.SERVICE_NAME)
                self.root.after(0, lambda: self.append_result_text("✅ 已启动服务", "success"))

                self.root.after(0, lambda: self.update_progress(88, "等待服务启动..."))
                self.root.after(0, lambda: self.append_result_text("⏳ 正在检测服务状态...", "info"))
                service_started, status_msg, elapsed = wait_for_service_start(
                    self.SERVICE_NAME,
                    self.PROCESS_NAME,
                    status_callback=lambda: self.root.after(0, self.refresh_service_status),
                    progress_callback=lambda r, t, e, total: self.root.after(0, self.update_progress, 88 + int(5 * r / t), f"等待服务启动... ({r}/{t})")
                )
                if not service_started:
                    self.root.after(0, lambda: self.update_progress(100, "服务启动失败"))
                    self.root.after(0, lambda: self.append_result_text(f"⚠️ 服务启动超时：{status_msg}", "warning"))
                    self.root.after(0, lambda: messagebox.showwarning("启动失败", f"升级后服务启动失败，请手动检查服务状态。"))
                    self.upgrade_btn.config(state=NORMAL)
                    self.install_btn.config(state=NORMAL)
                    self.uninstall_btn.config(state=NORMAL)
                    self.open_btn.config(state=NORMAL)
                    return
                else:
                    self.root.after(0, lambda: self.append_result_text(f"✅ 服务启动成功（耗时{elapsed}秒）", "success"))

                self.root.after(0, lambda: self.update_progress(93, "等待端口开放..."))
                current_port = get_port_from_conf(self.CONF_PATH, self.DEFAULT_PORT)
                try:
                    wait_for_port_open('localhost', current_port, max_wait=30)
                    self.root.after(0, lambda: self.append_result_text(f"✅ 端口{current_port}已开放", "success"))
                except TimeoutError:
                    self.root.after(0, lambda: self.append_result_text(f"⚠️ 端口{current_port}未开放（服务可能仍在启动）", "warning"))

                self.root.after(0, lambda: self.update_progress(96, "打开网盘页面..."))
                url = f"http://localhost:{current_port}"
                try:
                    webbrowser.open(url)
                    self.root.after(0, lambda: self.append_result_text(f"✅ 已打开网盘页面：{url}", "success"))
                except Exception as e:
                    self.root.after(0, lambda: self.append_result_text(f"⚠️ 打开页面失败，请手动访问：{url}，错误：{e}", "warning"))

                self.root.after(0, lambda: self.update_progress(100, "升级完成"))
                if os.path.exists(self.CLOUDREVE_EXE):
                    mod_time = os.path.getmtime(self.CLOUDREVE_EXE)
                    mod_time_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mod_time))
                else:
                    mod_time_str = "未知"
                self.root.after(0, lambda: self.append_result_text(f"""🎉 自动升级成功！
当前版本已更新为 {latest_version}，访问地址：{url}
新版本修改日期：{mod_time_str}
如遇问题，可使用备份文件 {backup_path} 恢复。""", "success"))
                self.root.after(0, lambda: self.update_statusbar("自动升级完成"))
            else:
                self.root.after(0, lambda: self.update_progress(100, "升级完成"))
                if os.path.exists(self.CLOUDREVE_EXE):
                    mod_time = os.path.getmtime(self.CLOUDREVE_EXE)
                    mod_time_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mod_time))
                else:
                    mod_time_str = "未知"
                self.root.after(0, lambda: self.append_result_text(f"""🎉 自动升级成功！
新版本文件已替换，但网盘服务不存在，请手动安装后使用。
新版本修改日期：{mod_time_str}
备份文件：{backup_path}""", "success"))
                self.root.after(0, lambda: self.update_statusbar("自动升级完成（服务未安装）"))

            try:
                if os.path.exists(zip_path):
                    os.remove(zip_path)
                if os.path.exists(extract_dir):
                    shutil.rmtree(extract_dir, ignore_errors=True)
                os.rmdir(TEMP_DIR)
            except:
                pass

        except Exception as e:
            log.error(f"自动升级异常: {e}")
            import traceback
            log.error(f"异常详情: {traceback.format_exc()}")
            self.root.after(0, lambda: self.append_result_text(f"❌ 自动升级过程发生异常：{str(e)}", "danger"))
            self.root.after(0, lambda: self.update_progress(100, "升级异常"))
            self.root.after(0, lambda: messagebox.showerror("错误", f"自动升级失败：{str(e)}"))
            self.root.after(0, lambda: self.update_statusbar("自动升级失败"))
        finally:
            self.upgrade_btn.config(state=NORMAL)
            self.install_btn.config(state=NORMAL)
            self.uninstall_btn.config(state=NORMAL)
            self.open_btn.config(state=NORMAL)
            self.refresh_service_status()

    def manual_upgrade_worker(self):
        try:
            self.install_btn.config(state=DISABLED)
            self.uninstall_btn.config(state=DISABLED)
            self.open_btn.config(state=DISABLED)
            self.upgrade_btn.config(state=DISABLED)

            self.root.after(0, lambda: self.update_progress(10, "检查管理员权限..."))
            if not check_admin_permission():
                self.root.after(0, lambda: self.update_progress(100, "权限不足"))
                self.root.after(0, lambda: self.append_result_text("""❌ 升级失败！
需要管理员权限，请右键以管理员身份运行""", "danger"))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                self.open_btn.config(state=NORMAL)
                self.upgrade_btn.config(state=NORMAL)
                return

            self.root.after(0, lambda: self.update_progress(15, "检查服务状态..."))
            service_exists_before = check_service_exists(self.SERVICE_NAME)
            self.root.after(0, lambda: self.append_result_text(f"ℹ️ 升级前服务状态：{'存在' if service_exists_before else '不存在'}", "info"))
            if not service_exists_before:
                self.root.after(0, lambda: self.append_result_text("⚠️ 当前网盘服务不存在，升级后将只替换程序文件，不启动服务。", "warning"))

            self.root.after(0, lambda: self.update_progress(20, "请选择新版本 cloudreve.exe..."))
            new_exe_path = filedialog.askopenfilename(
                title="选择新版本 cloudreve.exe",
                initialfile="cloudreve.exe",
                filetypes=[("可执行文件", "*.exe"), ("所有文件", "*.*")]
            )
            if not new_exe_path:
                self.root.after(0, lambda: self.update_progress(100, "用户取消升级"))
                self.root.after(0, lambda: self.append_result_text("⚠️ 未选择文件，升级已取消", "warning"))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                self.open_btn.config(state=NORMAL)
                self.upgrade_btn.config(state=NORMAL)
                return

            selected_filename = os.path.basename(new_exe_path)
            if selected_filename.lower() != "cloudreve.exe":
                self.root.after(0, lambda: self.update_progress(100, "文件名错误"))
                self.root.after(0, lambda: self.append_result_text(f"❌ 升级失败！\n所选文件名为：{selected_filename}\n请选择文件名必须为 cloudreve.exe 的文件。", "danger"))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                self.open_btn.config(state=NORMAL)
                self.upgrade_btn.config(state=NORMAL)
                return

            if not os.path.isfile(new_exe_path):
                self.root.after(0, lambda: self.update_progress(100, "无效文件"))
                self.root.after(0, lambda: self.append_result_text("❌ 选择的文件无效，请选择有效的 cloudreve.exe 文件", "danger"))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                self.open_btn.config(state=NORMAL)
                self.upgrade_btn.config(state=NORMAL)
                return

            self.root.after(0, lambda: self.append_result_text(f"📁 已选择升级文件：{new_exe_path}", "info"))

            current_port = get_port_from_conf(self.CONF_PATH, self.DEFAULT_PORT)
            self.root.after(0, lambda: self.append_result_text(f"ℹ️ 当前配置文件端口：{current_port}", "info"))

            if service_exists_before:
                self.root.after(0, lambda: self.update_progress(30, "停止服务..."))
                if check_service_exists(self.SERVICE_NAME):
                    stop_service(self.SERVICE_NAME)
                    self.root.after(0, lambda: self.append_result_text("✅ 已停止服务", "success"))
                else:
                    self.root.after(0, lambda: self.append_result_text("ℹ️ 未检测到服务", "info"))

                self.root.after(0, lambda: self.update_progress(40, "结束进程..."))
                if check_process_exists(self.PROCESS_NAME):
                    kill_process(self.PROCESS_NAME)
                    self.root.after(0, lambda: self.append_result_text("✅ 已结束进程", "success"))
                else:
                    self.root.after(0, lambda: self.append_result_text("ℹ️ 未检测到进程", "info"))
            else:
                self.root.after(0, lambda: self.update_progress(30, "检查残留进程..."))
                if check_process_exists(self.PROCESS_NAME):
                    kill_process(self.PROCESS_NAME)
                    self.root.after(0, lambda: self.append_result_text("✅ 已结束残留进程", "success"))
                else:
                    self.root.after(0, lambda: self.append_result_text("ℹ️ 未检测到进程", "info"))

            self.root.after(0, lambda: self.update_progress(50, "备份旧版本..."))
            backup_path = self.CLOUDREVE_EXE + ".bak"
            if os.path.exists(self.CLOUDREVE_EXE):
                try:
                    shutil.copy2(self.CLOUDREVE_EXE, backup_path)
                    self.root.after(0, lambda: self.append_result_text(f"✅ 已备份原文件至 {backup_path}", "success"))
                except Exception as e:
                    self.root.after(0, lambda: self.append_result_text(f"⚠️ 备份失败：{e}，继续升级", "warning"))
            else:
                self.root.after(0, lambda: self.append_result_text("ℹ️ 未找到原 cloudreve.exe，跳过备份", "info"))

            self.root.after(0, lambda: self.update_progress(60, "替换新版本..."))
            try:
                shutil.copy2(new_exe_path, self.CLOUDREVE_EXE)
                self.root.after(0, lambda: self.append_result_text(f"✅ 新版本已复制到 {self.CLOUDREVE_EXE}", "success"))
                if os.path.exists(self.CLOUDREVE_EXE):
                    mod_time = os.path.getmtime(self.CLOUDREVE_EXE)
                    mod_time_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mod_time))
                    self.root.after(0, lambda: self.append_result_text(f"📅 新版本文件修改日期：{mod_time_str}", "info"))
            except Exception as e:
                self.root.after(0, lambda: self.update_progress(100, "文件替换失败"))
                self.root.after(0, lambda: self.append_result_text(f"❌ 替换文件失败：{e}", "danger"))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                self.open_btn.config(state=NORMAL)
                self.upgrade_btn.config(state=NORMAL)
                return

            if service_exists_before:
                self.root.after(0, lambda: self.update_progress(70, "启动服务..."))
                start_service(self.SERVICE_NAME)
                self.root.after(0, lambda: self.append_result_text("✅ 已启动服务", "success"))

                self.root.after(0, lambda: self.update_progress(80, "等待服务启动..."))
                self.root.after(0, lambda: self.append_result_text("⏳ 正在检测服务状态...", "info"))
                service_started, status_msg, elapsed = wait_for_service_start(
                    self.SERVICE_NAME,
                    self.PROCESS_NAME,
                    status_callback=lambda: self.root.after(0, self.refresh_service_status),
                    progress_callback=lambda r, t, e, total: self.root.after(0, self.update_progress, 80 + int(10 * r / t), f"等待服务启动... ({r}/{t})")
                )
                if not service_started:
                    self.root.after(0, lambda: self.update_progress(100, "服务启动失败"))
                    self.root.after(0, lambda: self.append_result_text(f"""⚠️ 升级不完整！
服务启动超时（{elapsed}秒）
状态：{status_msg}
建议：手动启动服务或检查配置""", "warning"))
                    self.install_btn.config(state=NORMAL)
                    self.uninstall_btn.config(state=NORMAL)
                    self.open_btn.config(state=NORMAL)
                    self.upgrade_btn.config(state=NORMAL)
                    return
                else:
                    self.root.after(0, lambda: self.append_result_text(f"✅ 服务启动成功（耗时{elapsed}秒）", "success"))

                self.root.after(0, lambda: self.update_progress(90, "等待端口开放..."))
                def port_progress(elapsed, max_wait, remaining):
                    self.root.after(0, self.update_progress, 90 + int(5 * (elapsed / max_wait)), f"等待端口开放... ({elapsed:.1f}/{max_wait}秒)")
                try:
                    wait_for_port_open('localhost', current_port, max_wait=30, progress_callback=port_progress)
                    self.root.after(0, lambda: self.append_result_text(f"✅ 端口{current_port}已开放", "success"))
                except TimeoutError:
                    self.root.after(0, lambda: self.append_result_text(f"⚠️ 端口{current_port}未开放（服务可能仍在启动）", "warning"))

                self.root.after(0, lambda: self.update_progress(95, "打开网盘页面..."))
                url = f"http://localhost:{current_port}"
                try:
                    webbrowser.open(url)
                    self.root.after(0, lambda: self.append_result_text(f"✅ 已打开网盘页面：{url}", "success"))
                except Exception as e:
                    self.root.after(0, lambda: self.append_result_text(f"⚠️ 打开页面失败，请手动访问：{url}，错误：{e}", "warning"))

                self.root.after(0, lambda: self.update_progress(100, "升级完成"))
                if os.path.exists(self.CLOUDREVE_EXE):
                    mod_time = os.path.getmtime(self.CLOUDREVE_EXE)
                    mod_time_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mod_time))
                else:
                    mod_time_str = "未知"
                self.root.after(0, lambda: self.append_result_text(f"""🎉 手动升级成功！
当前版本已更新为所选文件，访问地址：{url}
新版本修改日期：{mod_time_str}
如遇问题，可使用备份文件 {backup_path} 恢复。""", "success"))
                self.root.after(0, lambda: self.update_statusbar("手动升级完成"))
            else:
                self.root.after(0, lambda: self.update_progress(100, "升级完成"))
                if os.path.exists(self.CLOUDREVE_EXE):
                    mod_time = os.path.getmtime(self.CLOUDREVE_EXE)
                    mod_time_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mod_time))
                else:
                    mod_time_str = "未知"
                self.root.after(0, lambda: self.append_result_text(f"""🎉 手动升级成功！
新版本文件已替换，但网盘服务不存在，请手动安装后使用。
新版本修改日期：{mod_time_str}
备份文件：{backup_path}""", "success"))
                self.root.after(0, lambda: self.update_statusbar("手动升级完成（服务未安装）"))

        except Exception as e:
            log.error(f"手动升级异常: {e}")
            import traceback
            log.error(f"异常详情: {traceback.format_exc()}")
            self.root.after(0, lambda: self.append_result_text(f"""❌ 手动升级过程发生异常：
{str(e)}
请查看日志文件获取详细信息""", "danger"))
            self.root.after(0, lambda: self.update_progress(100, "升级异常"))
            self.root.after(0, lambda: self.update_statusbar("手动升级失败"))
        finally:
            self.install_btn.config(state=NORMAL)
            self.uninstall_btn.config(state=NORMAL)
            self.open_btn.config(state=NORMAL)
            self.upgrade_btn.config(state=NORMAL)
            self.refresh_service_status()

if __name__ == "__main__":
    run_as_admin()
    root = ttk.Window(themename="flatly")
    app = CloudreveManagerGUI(root)
    root.mainloop()