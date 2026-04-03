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
import locale

# ========== 语言管理器 ==========
class LanguageManager:
    def __init__(self, lang_dir="lang", default_lang="zh_CN"):
        self.lang_dir = lang_dir
        self.default_lang = default_lang
        self.current_lang = default_lang
        self.strings = {}
        self.load_language(default_lang)

    def load_language(self, lang_code):
        file_path = os.path.join(self.lang_dir, f"{lang_code}.json")
        if not os.path.exists(file_path):
            file_path = os.path.join(self.lang_dir, f"{self.default_lang}.json")
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                self.strings = json.load(f)
            self.current_lang = lang_code
        except Exception as e:
            log.error(f"加载语言文件失败: {e}")
            self.strings = {}

    def tr(self, key, *args):
        text = self.strings.get(key, key)
        if args:
            try:
                text = text.format(*args)
            except:
                pass
        return text

    def get_current_lang(self):
        return self.current_lang

    def set_language(self, lang_code):
        self.load_language(lang_code)

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

# ========== 日志配置 ==========
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

# ========== 带重试的子进程执行 ==========
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

# ========== 配置文件操作函数 ==========
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
    try:
        if os.path.exists(conf_path):
            backup_path = conf_path + ".bak"
            shutil.copy2(conf_path, backup_path)
            log.info(f"已备份原配置文件至 {backup_path}")

        config = configparser.ConfigParser()
        config.optionxform = str
        config.read(conf_path, encoding='utf-8')

        if 'System' not in config:
            config['System'] = {}

        keys_to_remove = []
        for key in config['System'].keys():
            if key.lower() == 'listen':
                keys_to_remove.append(key)
        for key in keys_to_remove:
            del config['System'][key]
            log.info(f"已删除旧的 Listen 配置项: {key}")

        config['System']['Listen'] = f":{port}"

        with open(conf_path, 'w', encoding='utf-8') as f:
            config.write(f)

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
    common_paths = [
        r"C:\Program Files\MySQL\MySQL Server 8.0\bin\mysqldump.exe",
        r"C:\Program Files\MySQL\MySQL Server 5.7\bin\mysqldump.exe",
        r"C:\Program Files (x86)\MySQL\MySQL Server 8.0\bin\mysqldump.exe",
        r"C:\Program Files (x86)\MySQL\MySQL Server 5.7\bin\mysqldump.exe",
    ]
    for path in common_paths:
        if os.path.isfile(path):
            return path
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

# ========== 语言管理器实例 ==========
lang = LanguageManager()

# ========== GUI类 ==========
class CloudreveManagerGUI:
    def __init__(self, root):
        log.info("初始化Cloudreve服务管理工具")
        self.root = root
        self.root.title(lang.tr("app_title"))
        
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
        
        # 获取本机IP地址（用于界面显示）
        self.local_ip = self.get_local_ip()
        
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
        
        # 保存菜单引用以便更新
        self.config_menu = None
        self.db_menu = None
        self.service_menu = None
        self.lang_menu = None
        self.help_menu = None
        self.title_frame = None   # 安装提醒框
        self.port_config_frame = None
        self.progress_frame = None
        self.result_frame = None

        # ========== 新增：检查必要文件 ==========
        self.check_required_files()
        
        self.create_optimized_ui()
        self.refresh_service_status()

    # ========== 获取本机IP地址 ==========
    def get_local_ip(self):
        """获取本机局域网IPv4地址"""
        try:
            # 创建一个UDP套接字连接到外部地址，不实际发送数据
            with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
                s.connect(('8.8.8.8', 80))
                ip = s.getsockname()[0]
            return ip
        except Exception:
            try:
                # 备用方法：获取主机名对应的IP
                hostname = socket.gethostname()
                ip = socket.gethostbyname(hostname)
                if ip and ip != '127.0.0.1':
                    return ip
            except:
                pass
            return "未知"

    # ========== 新增：检查必要文件的方法 ==========
    def check_required_files(self):
        missing = []
        if not os.path.exists(self.CLOUDREVE_EXE):
            missing.append("cloudreve.exe")
        if not os.path.exists(self.WINSW_EXE):
            missing.append("winsw.exe")
        if not os.path.exists(self.WINSW_XML):
            missing.append("winsw.xml")
        if missing:
            missing_str = "\n".join(missing)
            messagebox.showwarning(
                lang.tr("missing_files_title"),
                lang.tr("missing_files_msg", missing_str),
                parent=self.root
            )

    def create_optimized_ui(self):
        main_container = ttk.Frame(self.root, padding=15)
        main_container.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        # ========== 安装提醒框 ==========
        is_on_c_drive = self.APP_DIR.startswith("C:") or self.APP_DIR.startswith("c:")
        reminder_text = lang.tr("install_reminder_warning") if is_on_c_drive else lang.tr("install_reminder_success")
        self.title_frame = ttk.LabelFrame(main_container, text=lang.tr("install_reminder_title"))
        self.title_frame.pack(fill=X, pady=(0, 8))
        self.sub_title = ttk.Label(self.title_frame, text=reminder_text, font=("微软雅黑", 9), bootstyle=WARNING if is_on_c_drive else SUCCESS)
        self.sub_title.pack(pady=8, padx=10)
        
        # ========== 端口配置框（标题后显示本机IP） ==========
        port_title = lang.tr("port_config_with_ip", self.local_ip)
        self.port_config_frame = ttk.LabelFrame(main_container, text=port_title)
        self.port_config_frame.pack(fill=X, pady=(0, 8))
        
        port_frame = ttk.Frame(self.port_config_frame)
        port_frame.pack(fill=X, padx=10, pady=8)
        
        self.port_label = ttk.Label(port_frame, text=lang.tr("port_label"), font=("微软雅黑", 11))
        self.port_label.grid(row=0, column=0, padx=(0, 8), sticky="w")
        self.port_entry = ttk.Entry(port_frame, textvariable=self.custom_port, font=("微软雅黑", 11), width=10, justify=CENTER)
        self.port_entry.grid(row=0, column=1, padx=(0, 10), sticky="w")
        self.check_btn = ttk.Button(port_frame, text=lang.tr("check_port_btn"), command=self.check_port_status, bootstyle=INFO, width=10)
        self.check_btn.grid(row=0, column=2, sticky="w")
        
        # 进度框
        self.progress_frame = ttk.LabelFrame(main_container, text=lang.tr("progress_ready"))
        self.progress_frame.pack(fill=X, pady=(0, 8))
        progress_inner = ttk.Frame(self.progress_frame)
        progress_inner.pack(fill=X, padx=10, pady=8)
        self.progress_var = ttk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_inner, variable=self.progress_var, maximum=100, bootstyle="SUCCESS.PROGRESS")
        self.progress_bar.pack(fill=X, padx=5, pady=(5, 3))
        self.progress_label = ttk.Label(progress_inner, text=lang.tr("progress_ready"), font=("微软雅黑", 9))
        self.progress_label.pack(fill=X, padx=5, pady=(0, 5))
        
        # 结果框
        self.result_frame = ttk.LabelFrame(main_container, text=lang.tr("result_label"))
        self.result_frame.pack(fill=BOTH, expand=True, pady=(0, 10))
        result_inner = ttk.Frame(self.result_frame)
        result_inner.pack(fill=BOTH, expand=True, padx=10, pady=6)
        result_scroll = ttk.Scrollbar(result_inner)
        result_scroll.pack(side=RIGHT, fill=Y)
        self.result_text = ttk.Text(result_inner, font=("微软雅黑", 9), wrap=WORD, yscrollcommand=result_scroll.set, height=8)
        self.result_text.pack(fill=BOTH, expand=True, padx=(0, 5))
        result_scroll.config(command=self.result_text.yview)
        self.result_text.config(state=DISABLED)
        self.append_result_text(lang.tr("ready_text"), "info")
        
        # ========== 菜单栏 ==========
        menubar = ttk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # 配置菜单
        self.config_menu = ttk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=lang.tr("menu_config"), menu=self.config_menu)
        self.config_menu.add_command(label=lang.tr("menu_backup"), command=self.backup_data)
        self.config_menu.add_command(label=lang.tr("menu_restore"), command=self.restore_data)
        self.config_menu.add_separator()
        self.config_menu.add_command(label=lang.tr("menu_edit_config"), command=self.edit_config)
        self.config_menu.add_separator()
        self.config_menu.add_command(label=lang.tr("menu_exit"), command=self.root.quit)
        
        # 切换数据库菜单
        self.db_menu = ttk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=lang.tr("menu_switch_db"), menu=self.db_menu)
        self.db_menu.add_command(label=lang.tr("menu_install_mysql"), command=self.install_mysql_database)
        self.db_menu.add_command(label=lang.tr("menu_use_default_db"), command=self.use_default_database)
        
        # 服务菜单
        self.service_menu = ttk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=lang.tr("menu_service"), menu=self.service_menu)
        self.service_menu.add_command(label=lang.tr("menu_open_panel"), command=self.start_open_cloudreve)
        self.service_menu.add_command(label=lang.tr("menu_upgrade"), command=self.start_upgrade)
        self.service_menu.add_separator()
        self.service_menu.add_command(label=lang.tr("menu_start_service"), command=self.start_service_action)
        self.service_menu.add_command(label=lang.tr("menu_stop_service"), command=self.stop_service_action)
        
        # 语言菜单
        self.lang_menu = ttk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=lang.tr("menu_language"), menu=self.lang_menu)
        self.lang_menu.add_command(label=lang.tr("lang_zh"), command=lambda: self.change_language("zh_CN"))
        self.lang_menu.add_command(label=lang.tr("lang_en"), command=lambda: self.change_language("en_US"))
        
        # 帮助菜单
        self.help_menu = ttk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=lang.tr("menu_help"), menu=self.help_menu)
        self.help_menu.add_command(label=lang.tr("menu_about"), command=self.show_about)
        
        # ========== 底部按钮 ==========
        button_container = ttk.Frame(main_container)
        btn_grid = ttk.Frame(button_container)
        btn_grid.pack(fill=X, padx=20)
        btn_grid.grid_columnconfigure(0, weight=1)
        btn_grid.grid_columnconfigure(1, weight=1)
        
        self.install_btn = ttk.Button(btn_grid, text=lang.tr("install_btn"), command=self.start_install, bootstyle=SUCCESS, padding=(15, 8), width=12)
        self.install_btn.grid(row=0, column=0, padx=(0, 6), pady=8, sticky="ew")
        
        self.uninstall_btn = ttk.Button(btn_grid, text=lang.tr("uninstall_btn"), command=self.start_uninstall, bootstyle=DANGER, padding=(15, 8), width=12)
        self.uninstall_btn.grid(row=0, column=1, padx=(6, 0), pady=8, sticky="ew")
        
        button_container.pack(side=BOTTOM, fill=X, pady=(0, 5))
        
        # 底部状态栏
        self.status_bar = ttk.Label(self.root, text=lang.tr("status_bar_ready"), relief=SUNKEN, anchor=W)
        self.status_bar.pack(side=BOTTOM, fill=X)

    # ========== 语言切换 ==========
    def change_language(self, lang_code):
        lang.set_language(lang_code)
        self.refresh_ui_texts()
        self.save_language_setting(lang_code)

    def save_language_setting(self, lang_code):
        try:
            config_file = os.path.join(self.APP_DIR, "settings.ini")
            config = configparser.ConfigParser()
            config.read(config_file, encoding='utf-8')
            if not config.has_section("Language"):
                config.add_section("Language")
            config.set("Language", "current", lang_code)
            with open(config_file, 'w', encoding='utf-8') as f:
                config.write(f)
        except Exception as e:
            log.error(f"保存语言设置失败: {e}")

    def load_language_setting(self):
        try:
            config_file = os.path.join(self.APP_DIR, "settings.ini")
            config = configparser.ConfigParser()
            config.read(config_file, encoding='utf-8')
            if config.has_section("Language") and config.has_option("Language", "current"):
                return config.get("Language", "current")
        except:
            pass
        return None

    def refresh_ui_texts(self):
        # 窗口标题
        self.root.title(lang.tr("app_title"))
        # 安装提醒框
        is_on_c_drive = self.APP_DIR.startswith("C:") or self.APP_DIR.startswith("c:")
        reminder_text = lang.tr("install_reminder_warning") if is_on_c_drive else lang.tr("install_reminder_success")
        self.title_frame.config(text=lang.tr("install_reminder_title"))
        self.sub_title.config(text=reminder_text)

        # 端口配置框（标题后显示本机IP）
        port_title = lang.tr("port_config_with_ip", self.local_ip)
        self.port_config_frame.config(text=port_title)
        self.port_label.config(text=lang.tr("port_label"))
        self.check_btn.config(text=lang.tr("check_port_btn"))

        # 进度框
        self.progress_frame.config(text=lang.tr("progress_ready"))
        self.progress_label.config(text=lang.tr("progress_ready"))

        # 结果框
        self.result_frame.config(text=lang.tr("result_label"))

        # 菜单栏
        menubar = self.root.nametowidget(self.root['menu'])
        menubar.entryconfig(0, label=lang.tr("menu_config"))
        menubar.entryconfig(1, label=lang.tr("menu_switch_db"))
        menubar.entryconfig(2, label=lang.tr("menu_service"))
        menubar.entryconfig(3, label=lang.tr("menu_language"))
        menubar.entryconfig(4, label=lang.tr("menu_help"))

        # 配置菜单
        self.config_menu.entryconfig(0, label=lang.tr("menu_backup"))
        self.config_menu.entryconfig(1, label=lang.tr("menu_restore"))
        self.config_menu.entryconfig(3, label=lang.tr("menu_edit_config"))
        self.config_menu.entryconfig(5, label=lang.tr("menu_exit"))

        # 切换数据库菜单
        self.db_menu.entryconfig(0, label=lang.tr("menu_install_mysql"))
        self.db_menu.entryconfig(1, label=lang.tr("menu_use_default_db"))

        # 服务菜单
        self.service_menu.entryconfig(0, label=lang.tr("menu_open_panel"))
        self.service_menu.entryconfig(1, label=lang.tr("menu_upgrade"))
        self.service_menu.entryconfig(3, label=lang.tr("menu_start_service"))
        self.service_menu.entryconfig(4, label=lang.tr("menu_stop_service"))

        # 语言菜单
        self.lang_menu.entryconfig(0, label=lang.tr("lang_zh"))
        self.lang_menu.entryconfig(1, label=lang.tr("lang_en"))

        # 帮助菜单
        self.help_menu.entryconfig(0, label=lang.tr("menu_about"))

        # 按钮
        self.install_btn.config(text=lang.tr("install_btn"))
        self.uninstall_btn.config(text=lang.tr("uninstall_btn"))

        # 状态栏（如果当前显示的是“就绪”，则更新，否则保留）
        current_status = self.status_bar.cget("text")
        if current_status in (lang.tr("status_bar_ready"), "就绪", "Ready"):
            self.status_bar.config(text=lang.tr("status_bar_ready"))

        # 刷新服务状态
        self.refresh_service_status()

    # ========== 辅助方法 ==========
    def update_statusbar(self, text, bootstyle=None):
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

    # ========== 服务状态刷新 ==========
    def refresh_service_status(self):
        try:
            service_status = get_service_status(self.SERVICE_NAME)
            process_exists = check_process_exists(self.PROCESS_NAME)
            if service_status == "RUNNING" and process_exists:
                status_text = lang.tr("service_running")
                bootstyle = "success"
            else:
                status_text = lang.tr("service_stopped")
                bootstyle = "danger"

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
            status_msg = lang.tr("status_service_db", status_text, db_type, db_extra)
            self.update_statusbar(status_msg, bootstyle=db_color)

            log.info(f"刷新状态 - 网盘服务：{status_text}，数据库：{db_type}")
        except Exception as e:
            log.error(f"刷新状态异常: {e}")
            self.update_statusbar(lang.tr("status_failed"), bootstyle="danger")

    def validate_port(self):
        try:
            port_str = self.custom_port.get().strip()
            if not port_str.isdigit():
                return False, lang.tr("port_not_digit")
            port = int(port_str)
            if port < 1 or port > 65535:
                return False, lang.tr("port_out_of_range")
            return True, port
        except Exception as e:
            return False, str(e)

    # ========== 配置备份与恢复 ==========
    def backup_data(self):
        if not os.path.exists(self.DATA_DIR):
            messagebox.showwarning(lang.tr("backup_confirm_title"), lang.tr("backup_confirm_msg"))
            return
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        default_filename = f"cloudreve_config_backup_{timestamp}.zip"
        backup_path = filedialog.asksaveasfilename(
            title=lang.tr("backup_title"),
            defaultextension=".zip",
            filetypes=[("ZIP 压缩文件", "*.zip"), ("所有文件", "*.*")],
            initialdir=self.APP_DIR,
            initialfile=default_filename
        )
        if not backup_path:
            return
        try:
            temp_backup_dir = os.path.join(TEMP_DIR, "backup_temp")
            if os.path.exists(temp_backup_dir):
                shutil.rmtree(temp_backup_dir)
            os.makedirs(temp_backup_dir, exist_ok=True)

            dest_data = os.path.join(temp_backup_dir, "data")
            shutil.copytree(self.DATA_DIR, dest_data, ignore=shutil.ignore_patterns("uploads"))

            if os.path.exists(self.CONF_PATH):
                config = configparser.ConfigParser()
                config.read(self.CONF_PATH, encoding='utf-8')
                if 'Database' in config and config['Database'].get('Type', '').lower() == 'mysql':
                    self.append_result_text(lang.tr("backup_detected_mysql"), "info")
                    sql_temp = os.path.join(temp_backup_dir, f"mysql_dump_{config['Database'].get('Name', 'cloudreve')}.sql")
                    success, err = dump_mysql_database(self.CONF_PATH, sql_temp)
                    if success:
                        self.append_result_text(lang.tr("backup_mysql_success"), "success")
                    else:
                        self.append_result_text(lang.tr("backup_mysql_failed", err), "warning")
                else:
                    self.append_result_text(lang.tr("backup_sqlite"), "info")

            with zipfile.ZipFile(backup_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, dirs, files in os.walk(temp_backup_dir):
                    for file in files:
                        full_path = os.path.join(root, file)
                        arcname = os.path.relpath(full_path, start=temp_backup_dir)
                        zipf.write(full_path, arcname)

            self.append_result_text(lang.tr("backup_success", backup_path), "success")
            messagebox.showinfo(lang.tr("backup_success_title"), lang.tr("backup_success", backup_path), parent=self.root)
        except Exception as e:
            log.error(f"备份配置失败: {e}")
            messagebox.showerror(lang.tr("backup_failed_title"), lang.tr("backup_failed", str(e)))
            self.append_result_text(lang.tr("backup_failed", str(e)), "danger")
        finally:
            if os.path.exists(temp_backup_dir):
                shutil.rmtree(temp_backup_dir, ignore_errors=True)
            self.refresh_service_status()

    def restore_data(self):
        zip_path = filedialog.askopenfilename(
            title=lang.tr("restore_title"),
            filetypes=[("ZIP 压缩文件", "*.zip"), ("所有文件", "*.*")]
        )
        if not zip_path:
            return
        if not os.path.exists(self.DATA_DIR):
            if not messagebox.askyesno(lang.tr("restore_confirm_title"), lang.tr("restore_confirm_msg_no_data")):
                return
        else:
            uploads_path = os.path.join(self.DATA_DIR, "uploads")
            if os.path.exists(uploads_path):
                if not messagebox.askyesno(lang.tr("restore_confirm_title"), lang.tr("restore_confirm_msg_with_uploads")):
                    return
            else:
                if not messagebox.askyesno(lang.tr("restore_confirm_title"), lang.tr("restore_confirm_msg_no_uploads")):
                    return

        temp_uploads = None
        try:
            uploads_path = os.path.join(self.DATA_DIR, "uploads")
            if os.path.exists(uploads_path):
                temp_uploads = uploads_path + ".backup_temp"
                shutil.move(uploads_path, temp_uploads)

            temp_restore_dir = os.path.join(TEMP_DIR, "restore_temp")
            if os.path.exists(temp_restore_dir):
                shutil.rmtree(temp_restore_dir)
            os.makedirs(temp_restore_dir, exist_ok=True)
            with zipfile.ZipFile(zip_path, 'r') as zipf:
                zipf.extractall(temp_restore_dir)

            sql_files = [f for f in os.listdir(temp_restore_dir) if f.endswith('.sql')]
            if sql_files:
                self.append_result_text(lang.tr("restore_sql_found", ', '.join(sql_files)), "info")
                self.append_result_text(lang.tr("restore_sql_manual"), "warning")
                for sql in sql_files:
                    shutil.copy2(os.path.join(temp_restore_dir, sql), os.path.join(self.APP_DIR, sql))
                    self.append_result_text(lang.tr("restore_sql_copied", sql), "info")

            if os.path.exists(os.path.join(temp_restore_dir, "data")):
                for item in os.listdir(self.DATA_DIR):
                    if item != "uploads":
                        item_path = os.path.join(self.DATA_DIR, item)
                        if os.path.isfile(item_path):
                            os.remove(item_path)
                        elif os.path.isdir(item_path):
                            shutil.rmtree(item_path)
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

            self.append_result_text(lang.tr("restore_success", zip_path), "success")
            messagebox.showinfo(lang.tr("restore_success_title"), lang.tr("restore_success", zip_path), parent=self.root)
            self.refresh_service_status()
        except Exception as e:
            log.error(f"恢复配置失败: {e}")
            if temp_uploads and os.path.exists(temp_uploads):
                try:
                    if os.path.exists(uploads_path):
                        shutil.rmtree(uploads_path)
                    shutil.move(temp_uploads, uploads_path)
                except:
                    pass
            messagebox.showerror(lang.tr("restore_failed_title"), lang.tr("restore_failed", str(e)))
            self.append_result_text(lang.tr("restore_failed", str(e)), "danger")
        finally:
            if os.path.exists(temp_restore_dir):
                shutil.rmtree(temp_restore_dir, ignore_errors=True)

    def auto_backup_config(self):
        if not os.path.exists(self.DATA_DIR):
            log.warning("自动备份失败：data 目录不存在")
            return
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        backup_filename = f"cloudreve_config_backup_{timestamp}.zip"
        backup_path = os.path.join(self.APP_DIR, backup_filename)
        try:
            temp_backup_dir = os.path.join(TEMP_DIR, "auto_backup_temp")
            if os.path.exists(temp_backup_dir):
                shutil.rmtree(temp_backup_dir)
            os.makedirs(temp_backup_dir, exist_ok=True)

            dest_data = os.path.join(temp_backup_dir, "data")
            shutil.copytree(self.DATA_DIR, dest_data, ignore=shutil.ignore_patterns("uploads"))

            if os.path.exists(self.CONF_PATH):
                config = configparser.ConfigParser()
                config.read(self.CONF_PATH, encoding='utf-8')
                if 'Database' in config and config['Database'].get('Type', '').lower() == 'mysql':
                    self.append_result_text(lang.tr("auto_backup_mysql"), "info")
                    sql_temp = os.path.join(temp_backup_dir, f"mysql_dump_{config['Database'].get('Name', 'cloudreve')}.sql")
                    success, err = dump_mysql_database(self.CONF_PATH, sql_temp)
                    if success:
                        self.append_result_text(lang.tr("auto_backup_mysql_success"), "success")
                    else:
                        self.append_result_text(lang.tr("auto_backup_mysql_failed", err), "warning")

            with zipfile.ZipFile(backup_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, dirs, files in os.walk(temp_backup_dir):
                    for file in files:
                        full_path = os.path.join(root, file)
                        arcname = os.path.relpath(full_path, start=temp_backup_dir)
                        zipf.write(full_path, arcname)

            self.append_result_text(lang.tr("auto_backup_success", backup_path), "success")
            log.info(f"自动备份成功：{backup_path}")
        except Exception as e:
            log.error(f"自动备份失败: {e}")
            self.append_result_text(lang.tr("auto_backup_failed", str(e)), "warning")
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
                messagebox.showerror(lang.tr("error"), lang.tr("create_config_failed", str(e)))
                return
        try:
            os.startfile(self.CONF_PATH)
            self.append_result_text(lang.tr("edit_config_opened", self.CONF_PATH), "info")
        except Exception as e:
            log.error(f"打开配置文件失败: {e}")
            messagebox.showerror(lang.tr("error"), lang.tr("open_config_failed", str(e)))

    def show_about(self):
        messagebox.showinfo(lang.tr("about_title"), lang.tr("about_text"))

    # ========== 统一服务安装与启动 ==========
    def _ensure_service_installed_and_started(self, start_after_install=True):
        """确保服务已安装，并根据参数决定是否启动。返回 (是否成功, 启动耗时, 错误消息)"""
        try:
            # 检查服务是否存在
            if not check_service_exists(self.SERVICE_NAME):
                self.append_result_text(lang.tr("service_not_installed_installing"), "info")
                if not check_admin_permission():
                    return False, 0, lang.tr("admin_required")
                if not os.path.exists(self.WINSW_EXE):
                    return False, 0, lang.tr("winsw_missing")
                # 安装服务
                try:
                    run_cmd_with_retry([self.WINSW_EXE, "install"], timeout=30)
                    self.append_result_text(lang.tr("service_installed_ok"), "success")
                except Exception as e:
                    return False, 0, lang.tr("service_install_failed", str(e))
                # 安装后稍等片刻
                time.sleep(1)

            if not start_after_install:
                return True, 0, ""

            # 如果服务已存在，则启动它
            # 先尝试停止已有的服务（如果有的话）
            if check_service_exists(self.SERVICE_NAME):
                service_status = get_service_status(self.SERVICE_NAME)
                if service_status == "RUNNING":
                    self.append_result_text(lang.tr("service_already_running"), "info")
                    return True, 0, ""
                # 启动服务
                self.append_result_text(lang.tr("starting_service"), "info")
                start_service(self.SERVICE_NAME)

                # 等待服务启动
                service_started, status_msg, elapsed = wait_for_service_start(
                    self.SERVICE_NAME,
                    self.PROCESS_NAME,
                    status_callback=lambda: self.root.after(0, self.refresh_service_status),
                    progress_callback=None  # 进度由调用者处理
                )
                if service_started:
                    return True, elapsed, ""
                else:
                    return False, elapsed, status_msg
            else:
                return False, 0, lang.tr("service_not_found")
        except Exception as e:
            log.error(f"确保服务安装启动失败: {e}")
            return False, 0, str(e)

    # ========== 使用默认数据库（切换回 SQLite） ==========
    def use_default_database(self):
        if not os.path.exists(self.CONF_PATH):
            self.append_result_text(lang.tr("use_default_no_config"), "danger")
            messagebox.showerror(lang.tr("error"), lang.tr("use_default_no_config"))
            return
        if not messagebox.askyesno(lang.tr("switch_db_confirm_title"), lang.tr("switch_db_confirm_msg"), parent=self.root):
            return

        self.install_btn.config(state=DISABLED)
        self.uninstall_btn.config(state=DISABLED)
        threading.Thread(target=self._use_default_db_worker, daemon=True).start()

    def _use_default_db_worker(self):
        try:
            self.root.after(0, lambda: self.update_progress(0, lang.tr("switch_start")))
            backup_path = self.CONF_PATH + ".bak"
            shutil.copy2(self.CONF_PATH, backup_path)
            self.root.after(0, lambda: self.append_result_text(lang.tr("switch_backup", backup_path), "info"))
            self.root.after(0, lambda: self.update_progress(20, lang.tr("switch_backup_complete")))

            config = configparser.ConfigParser()
            config.optionxform = str
            config.read(self.CONF_PATH, encoding='utf-8')
            if 'Database' in config:
                del config['Database']
                self.root.after(0, lambda: self.append_result_text(lang.tr("switch_remove_db_section"), "success"))
            else:
                self.root.after(0, lambda: self.append_result_text(lang.tr("switch_no_db_section"), "info"))

            with open(self.CONF_PATH, 'w', encoding='utf-8') as f:
                config.write(f)
            self.root.after(0, lambda: self.update_progress(40, lang.tr("switch_config_updated")))

            # 确保服务已安装并启动
            ok, elapsed, err_msg = self._ensure_service_installed_and_started(start_after_install=True)
            if not ok:
                self.root.after(0, lambda: self.append_result_text(lang.tr("switch_service_failed", err_msg), "warning"))
                self.root.after(0, lambda: messagebox.showwarning(lang.tr("start_failed"), lang.tr("switch_service_failed", err_msg)))
                return
            else:
                self.root.after(0, lambda: self.append_result_text(lang.tr("switch_service_started", elapsed), "success"))

            current_port = get_port_from_conf(self.CONF_PATH, self.DEFAULT_PORT)
            try:
                wait_for_port_open('localhost', current_port, max_wait=30)
                self.root.after(0, lambda: self.append_result_text(lang.tr("switch_port_open", current_port), "success"))
            except TimeoutError:
                self.root.after(0, lambda: self.append_result_text(lang.tr("switch_port_timeout", current_port), "warning"))

            self.root.after(0, lambda: self.update_progress(100, lang.tr("switch_complete")))
            self.root.after(0, lambda: messagebox.showinfo(lang.tr("switch_success_title"), lang.tr("switch_db_success"), parent=self.root))

        except Exception as e:
            log.error(f"切换默认数据库失败: {e}")
            import traceback
            log.error(f"异常详情: {traceback.format_exc()}")
            self.root.after(0, lambda: self.append_result_text(lang.tr("switch_failed", str(e)), "danger"))
            self.root.after(0, lambda: messagebox.showerror(lang.tr("error"), lang.tr("switch_failed", str(e))))
            self.root.after(0, lambda: self.update_progress(100, lang.tr("switch_failed_short")))
        finally:
            self.root.after(0, lambda: self.install_btn.config(state=NORMAL))
            self.root.after(0, lambda: self.uninstall_btn.config(state=NORMAL))
            self.root.after(0, self.refresh_service_status)

    # ========== 安装 MySQL 数据库（后台线程版） ==========
    def install_mysql_database(self):
        self.install_btn.config(state=DISABLED)
        self.uninstall_btn.config(state=DISABLED)
        threading.Thread(target=self._install_mysql_worker, daemon=True).start()

    def _install_mysql_worker(self):
        try:
            # 增加管理员权限检查
            if not check_admin_permission():
                self.root.after(0, lambda: self.append_result_text(lang.tr("admin_required"), "danger"))
                return

            if not os.path.exists(self.CONF_PATH):
                self.root.after(0, lambda: self.append_result_text(lang.tr("install_mysql_no_config"), "danger"))
                self.root.after(0, lambda: messagebox.showerror(lang.tr("error"), lang.tr("install_mysql_no_config")))
                return

            self.root.after(0, lambda: self.append_result_text(lang.tr("install_mysql_detecting"), "info"))
            mysql_service_name = None
            for name in ["MySQL80", "MySQL"]:
                if check_service_exists(name):
                    mysql_service_name = name
                    break

            if not mysql_service_name:
                self.root.after(0, lambda: self.append_result_text(lang.tr("install_mysql_no_service"), "danger"))
                self.root.after(0, lambda: messagebox.showerror(lang.tr("error"), lang.tr("install_mysql_no_service")))
                return

            self.root.after(0, lambda: self.append_result_text(lang.tr("install_mysql_service_found", mysql_service_name), "success"))

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
                        self.root.after(0, lambda: self.append_result_text(lang.tr("install_mysql_client_found", mysql_client), "info"))
                    else:
                        raise Exception("无法解析服务二进制路径")
                else:
                    raise Exception("sc qc 命令失败")
            except Exception as e:
                self.root.after(0, lambda: self.append_result_text(lang.tr("install_mysql_client_error", str(e)), "danger"))
                self.root.after(0, lambda: messagebox.showerror(lang.tr("error"), lang.tr("install_mysql_client_error", str(e))))
                return

            self.root.after(0, lambda: self.append_result_text(lang.tr("install_mysql_testing_root"), "info"))
            mysql_auth = None
            default_pwd = "cloudreve"
            try:
                test_cmd = [mysql_client, "-u", "root", f"-p{default_pwd}", "-e", "SELECT 1;"]
                result = subprocess.run(test_cmd, capture_output=True, timeout=10, creationflags=0x08000000)
                if result.returncode == 0:
                    mysql_auth = ["-u", "root", f"-p{default_pwd}"]
                    self.root.after(0, lambda: self.append_result_text(lang.tr("install_mysql_default_password_ok"), "success"))
                else:
                    error_msg = result.stderr.decode('gbk', errors='ignore')
                    if "Access denied" in error_msg:
                        self.root.after(0, lambda: self._prompt_mysql_password(mysql_client, default_pwd))
                        return
                    else:
                        raise Exception(f"连接失败：{error_msg[:200]}")
            except Exception as e:
                self.root.after(0, lambda: self.append_result_text(lang.tr("install_mysql_test_failed", str(e)), "danger"))
                self.root.after(0, lambda: messagebox.showerror(lang.tr("error"), lang.tr("install_mysql_test_failed", str(e))))
                return

            if mysql_auth:
                self._continue_mysql_config(mysql_client, mysql_auth)

        except Exception as e:
            log.error(f"安装 MySQL 数据库失败: {e}")
            import traceback
            log.error(f"异常详情: {traceback.format_exc()}")
            self.root.after(0, lambda: self.append_result_text(lang.tr("install_mysql_failed", str(e)), "danger"))
            self.root.after(0, lambda: messagebox.showerror(lang.tr("error"), lang.tr("install_mysql_failed", str(e))))
        finally:
            self.root.after(0, lambda: self.install_btn.config(state=NORMAL))
            self.root.after(0, lambda: self.uninstall_btn.config(state=NORMAL))
            self.root.after(0, self.refresh_service_status)

    def _prompt_mysql_password(self, mysql_client, default_pwd):
        root_pwd = simpledialog.askstring(lang.tr("mysql_root_password_title"),
                                           lang.tr("mysql_root_password_prompt"),
                                           show='*', parent=self.root)
        if root_pwd is None:
            self.install_btn.config(state=NORMAL)
            self.uninstall_btn.config(state=NORMAL)
            return
        threading.Thread(target=self._test_mysql_password, args=(mysql_client, root_pwd), daemon=True).start()

    def _test_mysql_password(self, mysql_client, root_pwd):
        try:
            test_cmd = [mysql_client, "-u", "root", f"-p{root_pwd}", "-e", "SELECT 1;"]
            result = subprocess.run(test_cmd, capture_output=True, timeout=10, creationflags=0x08000000)
            if result.returncode == 0:
                mysql_auth = ["-u", "root", f"-p{root_pwd}"]
                self.root.after(0, lambda: self.append_result_text(lang.tr("mysql_manual_password_ok"), "success"))
                self._continue_mysql_config(mysql_client, mysql_auth)
            else:
                error_msg = result.stderr.decode('gbk', errors='ignore')
                raise Exception(f"连接失败：{error_msg[:200]}")
        except Exception as e:
            self.root.after(0, lambda: self.append_result_text(lang.tr("mysql_manual_password_failed", str(e)), "danger"))
            self.root.after(0, lambda: messagebox.showerror(lang.tr("error"), lang.tr("mysql_manual_password_failed", str(e))))
            self.root.after(0, lambda: self.install_btn.config(state=NORMAL))
            self.root.after(0, lambda: self.uninstall_btn.config(state=NORMAL))
            self.root.after(0, self.refresh_service_status)

    def _ask_mysql_app_credentials(self, default_user="cloudreve", default_pass="your_password"):
        try:
            user = simpledialog.askstring(
                lang.tr("mysql_app_user_title"),
                lang.tr("mysql_app_user_prompt", default_user),
                parent=self.root,
                initialvalue=default_user
            )
            if user is None:
                user = default_user
                self._show_default_cred_warning()
            else:
                user = user.strip()
                while len(user) < 4:
                    messagebox.showwarning(lang.tr("invalid_user_title"), lang.tr("invalid_user_msg"))
                    user = simpledialog.askstring(
                        lang.tr("mysql_app_user_title"),
                        lang.tr("mysql_app_user_retry", default_user),
                        parent=self.root,
                        initialvalue=user
                    )
                    if user is None:
                        user = default_user
                        self._show_default_cred_warning()
                        break
                    user = user.strip()

            pwd = simpledialog.askstring(
                lang.tr("mysql_app_pass_title"),
                lang.tr("mysql_app_pass_prompt", default_pass),
                parent=self.root,
                show='*',
                initialvalue=default_pass
            )
            if pwd is None:
                pwd = default_pass
                self._show_default_cred_warning()
            else:
                pwd = pwd.strip()
                while len(pwd) < 4:
                    messagebox.showwarning(lang.tr("invalid_pass_title"), lang.tr("invalid_pass_msg"))
                    pwd = simpledialog.askstring(
                        lang.tr("mysql_app_pass_title"),
                        lang.tr("mysql_app_pass_retry", default_pass),
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
        messagebox.showinfo(lang.tr("default_cred_title"), lang.tr("default_cred_msg"))

    def _continue_mysql_config(self, mysql_client, mysql_auth):
        try:
            self.root.after(0, lambda: self.update_progress(0, lang.tr("mysql_config_start")))

            shutil.copy2(self.CONF_PATH, self.CONF_PATH + ".install_mysql_bak")
            self.root.after(0, lambda: self.append_result_text(lang.tr("mysql_backup_ok"), "info"))
            self.root.after(0, lambda: self.update_progress(10, lang.tr("mysql_backup_complete")))

            self.mysql_cred_event.clear()
            self.root.after(0, lambda: self._ask_mysql_app_credentials())
            waited = self.mysql_cred_event.wait(timeout=300)
            if not waited:
                self.root.after(0, lambda: self.append_result_text(lang.tr("mysql_cred_timeout"), "warning"))
                self.mysql_app_user = "cloudreve"
                self.mysql_app_pass = "your_password"
            else:
                self.root.after(0, lambda: self.append_result_text(lang.tr("mysql_cred_ok", self.mysql_app_user, self.mysql_app_pass), "success"))

            self.root.after(0, lambda: self.append_result_text(lang.tr("mysql_creating_db_user"), "info"))
            self.root.after(0, lambda: self.update_progress(20, lang.tr("mysql_creating_db_user")))
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
                    self.root.after(0, lambda: self.append_result_text(lang.tr("mysql_sql_error", error_msg[:200]), "danger"))
                    self.root.after(0, lambda: messagebox.showerror(lang.tr("error"), lang.tr("mysql_sql_error", error_msg[:200])))
                    return
            self.root.after(0, lambda: self.append_result_text(lang.tr("mysql_db_user_created"), "success"))
            self.root.after(0, lambda: self.update_progress(30, lang.tr("mysql_db_user_created")))

            self.root.after(0, lambda: self.append_result_text(lang.tr("mysql_updating_config"), "info"))
            self.root.after(0, lambda: self.update_progress(40, lang.tr("mysql_updating_config")))
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
            self.root.after(0, lambda: self.append_result_text(lang.tr("mysql_config_updated"), "success"))
            self.root.after(0, lambda: self.update_progress(50, lang.tr("mysql_config_updated")))

            root_pwd = ""
            for arg in mysql_auth:
                if arg.startswith("-p"):
                    root_pwd = arg[2:]
                    break
            if not root_pwd:
                root_pwd = "unknown"
            cred_file = os.path.join(self.APP_DIR, "mysql_credentials.txt")
            try:
                with open(cred_file, 'w', encoding='utf-8') as f:
                    f.write(f"MySQL root 密码：{root_pwd}\n")
                    f.write(f"Cloudreve 数据库用户名：{self.mysql_app_user}\n")
                    f.write(f"Cloudreve 数据库密码：{self.mysql_app_pass}\n")
                self.root.after(0, lambda: self.append_result_text(lang.tr("mysql_cred_saved", cred_file), "success"))
            except Exception as e:
                self.root.after(0, lambda: self.append_result_text(lang.tr("mysql_cred_save_failed", str(e)), "warning"))

            self.root.after(0, lambda: self.append_result_text(lang.tr("mysql_auto_backup"), "info"))
            self.auto_backup_config()

            # ========== 服务启动部分修复 ==========
            self.root.after(0, lambda: self.append_result_text(lang.tr("mysql_config_service"), "info"))
            self.root.after(0, lambda: self.update_progress(55, lang.tr("mysql_config_service")))

            # 确保服务已安装并启动
            ok, elapsed, err_msg = self._ensure_service_installed_and_started(start_after_install=True)
            if not ok:
                self.root.after(0, lambda: self.append_result_text(lang.tr("mysql_service_failed", err_msg), "warning"))
                self.root.after(0, lambda: messagebox.showwarning(lang.tr("start_failed"), lang.tr("mysql_service_failed", err_msg)))
                return
            else:
                self.root.after(0, lambda: self.append_result_text(lang.tr("mysql_service_started", elapsed), "success"))

            current_port = get_port_from_conf(self.CONF_PATH, self.DEFAULT_PORT)
            try:
                wait_for_port_open('localhost', current_port, max_wait=30)
                self.root.after(0, lambda: self.append_result_text(lang.tr("mysql_port_open", current_port), "success"))
            except TimeoutError:
                self.root.after(0, lambda: self.append_result_text(lang.tr("mysql_port_timeout", current_port), "warning"))

            self.root.after(0, lambda: self.update_progress(95, lang.tr("mysql_ready_open")))
            url = f"http://localhost:{current_port}"

            def ask_open():
                result = messagebox.askokcancel(
                    lang.tr("mysql_config_success_title"),
                    lang.tr("mysql_config_success", url),
                    parent=self.root
                )
                if result:
                    webbrowser.open(url)
                    self.root.after(0, lambda: self.append_result_text(lang.tr("mysql_page_opened", url), "success"))
                else:
                    self.root.after(0, lambda: self.append_result_text(lang.tr("mysql_page_cancel"), "info"))
                self.root.after(0, lambda: self.update_progress(100, lang.tr("mysql_config_complete")))
                self.root.after(0, lambda: self.update_statusbar(lang.tr("mysql_config_done"), bootstyle="success"))
            self.root.after(0, ask_open)

        except Exception as e:
            log.error(f"继续配置 MySQL 失败: {e}")
            import traceback
            log.error(f"异常详情: {traceback.format_exc()}")
            self.root.after(0, lambda: self.append_result_text(lang.tr("mysql_config_failed", str(e)), "danger"))
            self.root.after(0, lambda: messagebox.showerror(lang.tr("error"), lang.tr("mysql_config_failed", str(e))))
            self.root.after(0, lambda: self.update_progress(100, lang.tr("mysql_config_failed_short")))

    # ========== 打开网盘 ==========
    def start_open_cloudreve(self):
        self.clear_result_text()
        self.append_result_text(lang.tr("open_start"), "info")
        open_thread = threading.Thread(target=self.open_cloudreve_worker, daemon=True)
        open_thread.start()

    def open_cloudreve_worker(self):
        try:
            self.install_btn.config(state=DISABLED)
            self.uninstall_btn.config(state=DISABLED)

            self.root.after(0, lambda: self.update_progress(10, lang.tr("open_check_service")))
            # 确保服务已安装并启动
            ok, elapsed, err_msg = self._ensure_service_installed_and_started(start_after_install=True)
            if not ok:
                self.root.after(0, lambda: self.update_progress(100, lang.tr("open_service_failed")))
                self.root.after(0, lambda: self.append_result_text(lang.tr("open_service_failed_msg", err_msg), "danger"))
                self.root.after(0, lambda: messagebox.showwarning(lang.tr("open_failed"), lang.tr("open_service_failed_msg", err_msg)))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                return
            else:
                self.root.after(0, lambda: self.append_result_text(lang.tr("open_service_started", elapsed), "success"))

            self.root.after(0, lambda: self.update_progress(80, lang.tr("open_wait_port")))
            current_port = get_port_from_conf(self.CONF_PATH, self.DEFAULT_PORT)
            self.root.after(0, lambda: self.append_result_text(lang.tr("open_wait_port_msg"), "info"))
            try:
                def port_progress(elapsed, max_wait, remaining):
                    self.root.after(0, self.update_progress, 80 + int(20 * (elapsed / max_wait)), lang.tr("open_port_progress", elapsed, max_wait))
                wait_for_port_open('localhost', current_port, max_wait=30, progress_callback=port_progress)
                self.root.after(0, lambda: self.append_result_text(lang.tr("open_port_ok", current_port), "success"))
            except TimeoutError:
                self.root.after(0, lambda: self.update_progress(100, lang.tr("open_port_timeout")))
                self.root.after(0, lambda: self.append_result_text(lang.tr("open_port_timeout_msg", current_port), "warning"))
                self.root.after(0, lambda: messagebox.showwarning(lang.tr("open_failed"), lang.tr("open_port_timeout_msg", current_port)))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                return

            self.root.after(0, lambda: self.update_progress(95, lang.tr("open_open_browser")))
            url = f"http://localhost:{current_port}"
            try:
                webbrowser.open(url)
                self.root.after(0, lambda: self.append_result_text(lang.tr("open_browser_ok", url), "success"))
                log.info(f"打开网盘成功：{url}")
            except Exception as e:
                self.root.after(0, lambda: self.append_result_text(lang.tr("open_browser_failed", url, str(e)), "warning"))
                self.root.after(0, lambda: messagebox.showwarning(lang.tr("open_failed"), lang.tr("open_browser_failed_msg", url)))

            self.root.after(0, lambda: self.update_progress(100, lang.tr("open_complete")))
            self.root.after(0, lambda: self.append_result_text(lang.tr("open_complete_msg"), "success"))

        except Exception as e:
            log.error(f"打开网盘异常: {e}")
            import traceback
            log.error(f"异常详情: {traceback.format_exc()}")
            self.root.after(0, lambda: self.append_result_text(lang.tr("open_exception", str(e)), "danger"))
            self.root.after(0, lambda: self.update_progress(100, lang.tr("open_exception_short")))
            self.root.after(0, lambda: messagebox.showerror(lang.tr("error"), lang.tr("open_exception", str(e))))
        finally:
            self.install_btn.config(state=NORMAL)
            self.uninstall_btn.config(state=NORMAL)
            self.refresh_service_status()

    # ========== 端口检测 ==========
    def check_port_status(self):
        try:
            self.clear_result_text()
            self.append_result_text(lang.tr("port_check_start"), "info")
            is_valid, port = self.validate_port()
            if not is_valid:
                messagebox.showerror(lang.tr("error"), port)
                self.append_result_text(lang.tr("port_check_invalid", port), "danger")
                return
            self.update_progress(0, lang.tr("port_check_progress"))
            conf_port = get_port_from_conf(self.CONF_PATH, self.DEFAULT_PORT)
            port_open = is_port_open('localhost', port)
            is_used = is_port_occupied(port)

            report = lang.tr("port_report", port, conf_port, '✅' if port_open else '❌', '⚠️' if is_used else '✅')

            can_install = False
            if is_used:
                proc = get_process_using_port(port)
                if proc == "cloudreve.exe":
                    can_install = True
                    report += lang.tr("port_occupied_by_self", port, proc)
                else:
                    can_install = False
                    report += lang.tr("port_occupied_other", port, proc if proc else '未知')
            else:
                can_install = True
                if port != conf_port:
                    report += lang.tr("port_mismatch", port, conf_port)
                else:
                    report += lang.tr("port_available", port)

            if can_install:
                report += lang.tr("can_install")
                self.update_progress(100, lang.tr("port_check_ok_install"))
                self.append_result_text(report, "success")
            else:
                report += lang.tr("cannot_install")
                self.update_progress(100, lang.tr("port_check_no_install"))
                self.append_result_text(report, "danger")
        except Exception as e:
            log.error(f"检测端口异常: {e}")
            self.append_result_text(lang.tr("port_check_exception", str(e)), "danger")

    # ========== 安装模块 ==========
    def start_install(self):
        self.clear_result_text()
        self.append_result_text(lang.tr("install_start"), "info")
        install_thread = threading.Thread(target=self.install_worker, daemon=True)
        install_thread.start()

    def install_worker(self):
        try:
            self.install_btn.config(state=DISABLED)
            self.uninstall_btn.config(state=DISABLED)
            
            self.root.after(0, lambda: self.update_progress(5, lang.tr("install_validate_port")))
            is_port_valid, current_port = self.validate_port()
            if not is_port_valid:
                self.root.after(0, lambda: self.update_progress(100, lang.tr("install_port_invalid")))
                self.root.after(0, lambda: self.append_result_text(lang.tr("install_port_invalid_msg", current_port), "danger"))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                return
            
            self.root.after(0, lambda: self.update_progress(8, lang.tr("install_check_port_occupied")))
            if is_port_occupied(current_port):
                proc = get_process_using_port(current_port)
                if proc == "cloudreve.exe":
                    self.root.after(0, lambda: self.append_result_text(lang.tr("install_port_occupied_self", current_port, proc), "info"))
                else:
                    self.root.after(0, lambda: self.update_progress(100, lang.tr("install_port_occupied_other")))
                    self.root.after(0, lambda: self.append_result_text(lang.tr("install_port_occupied_other_msg", current_port, proc if proc else '未知'), "danger"))
                    self.install_btn.config(state=NORMAL)
                    self.uninstall_btn.config(state=NORMAL)
                    return
            
            self.root.after(0, lambda: self.update_progress(10, lang.tr("install_check_admin")))
            if not check_admin_permission():
                self.root.after(0, lambda: self.update_progress(100, lang.tr("install_admin_failed")))
                self.root.after(0, lambda: self.append_result_text(lang.tr("install_admin_required"), "danger"))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                return
            
            self.root.after(0, lambda: self.update_progress(15, lang.tr("install_check_files")))
            if not os.path.exists(self.WINSW_EXE):
                self.root.after(0, lambda: self.update_progress(100, lang.tr("install_missing_winsw")))
                self.root.after(0, lambda: self.append_result_text(lang.tr("install_missing_winsw_msg", self.WINSW_EXE), "danger"))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                return
            if not os.path.exists(self.WINSW_XML):
                self.root.after(0, lambda: self.update_progress(100, lang.tr("install_missing_xml")))
                self.root.after(0, lambda: self.append_result_text(lang.tr("install_missing_xml_msg", self.WINSW_XML), "danger"))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                return
            
            self.root.after(0, lambda: self.update_progress(20, lang.tr("install_firewall", current_port)))
            firewall_success, firewall_error = add_firewall_rule(current_port)
            if not firewall_success:
                self.root.after(0, lambda: self.append_result_text(lang.tr("install_firewall_warning", firewall_error), "warning"))
            
            self.root.after(0, lambda: self.update_progress(25, lang.tr("install_stop_service")))
            if check_service_exists(self.SERVICE_NAME):
                stop_service(self.SERVICE_NAME)
                self.root.after(0, lambda: self.append_result_text(lang.tr("install_stopped_service"), "success"))
            else:
                self.root.after(0, lambda: self.append_result_text(lang.tr("install_no_service"), "info"))
            
            self.root.after(0, lambda: self.update_progress(35, lang.tr("install_kill_process")))
            if check_process_exists(self.PROCESS_NAME):
                kill_process(self.PROCESS_NAME)
                self.root.after(0, lambda: self.append_result_text(lang.tr("install_killed_process"), "success"))
            else:
                self.root.after(0, lambda: self.append_result_text(lang.tr("install_no_process"), "info"))
            
            self.root.after(0, lambda: self.update_progress(40, lang.tr("install_install_service")))
            run_cmd_with_retry([self.WINSW_EXE, "install"], timeout=30)
            self.root.after(0, lambda: self.append_result_text(lang.tr("install_service_installed"), "success"))
            
            self.root.after(0, lambda: self.update_progress(45, lang.tr("install_start_service")))
            start_service(self.SERVICE_NAME)
            self.root.after(0, lambda: self.append_result_text(lang.tr("install_start_command"), "success"))
            
            self.root.after(0, lambda: self.update_progress(50, lang.tr("install_wait_service")))
            self.root.after(0, lambda: self.append_result_text(lang.tr("install_wait_service_msg"), "info"))
            service_started, status_msg, elapsed = wait_for_service_start(
                self.SERVICE_NAME,
                self.PROCESS_NAME,
                status_callback=lambda: self.root.after(0, self.refresh_service_status),
                progress_callback=lambda r, t, e, total: self.root.after(0, self.update_progress, 50 + int(30 * r / t), lang.tr("install_wait_progress", r, t))
            )
            if not service_started:
                self.root.after(0, lambda: self.update_progress(100, lang.tr("install_service_failed")))
                self.root.after(0, lambda: self.append_result_text(lang.tr("install_service_failed_msg", elapsed, status_msg), "warning"))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                return
            else:
                self.root.after(0, lambda: self.append_result_text(lang.tr("install_service_ok", elapsed), "success"))
            
            self.root.after(0, lambda: self.update_progress(55, lang.tr("install_wait_config")))
            def conf_progress(elapsed, max_wait, remaining):
                self.root.after(0, self.update_progress, 55 + int(25 * (elapsed / max_wait)), lang.tr("install_config_progress", elapsed, max_wait))
            conf_generated = wait_for_conf_file(self.CONF_PATH, timeout=120, progress_callback=conf_progress)
            if not conf_generated:
                self.root.after(0, lambda: self.update_progress(100, lang.tr("install_config_timeout")))
                self.root.after(0, lambda: self.append_result_text(lang.tr("install_config_timeout_msg", self.CONF_PATH), "warning"))
            else:
                self.root.after(0, lambda: self.update_progress(60, lang.tr("install_sync_port")))
                modify_success = modify_conf_port(self.CONF_PATH, current_port)
                if modify_success:
                    self.root.after(0, lambda: self.append_result_text(lang.tr("install_port_synced", current_port), "success"))
                else:
                    self.root.after(0, lambda: self.append_result_text(lang.tr("install_port_sync_failed"), "warning"))
                
                self.root.after(0, lambda: self.update_progress(70, lang.tr("install_restart_service")))
                stop_service(self.SERVICE_NAME)
                time.sleep(3)
                start_service(self.SERVICE_NAME)
                self.root.after(0, lambda: self.append_result_text(lang.tr("install_restarted"), "success"))
                
                self.root.after(0, lambda: self.update_progress(80, lang.tr("install_wait_port")))
                def port_progress(elapsed, max_wait, remaining):
                    self.root.after(0, self.update_progress, 80 + int(15 * (elapsed / max_wait)), lang.tr("install_port_progress", elapsed, max_wait))
                try:
                    wait_for_port_open('localhost', current_port, max_wait=30, progress_callback=port_progress)
                    self.root.after(0, lambda: self.append_result_text(lang.tr("install_port_ok", current_port), "success"))
                except TimeoutError:
                    self.root.after(0, lambda: self.append_result_text(lang.tr("install_port_timeout", current_port), "warning"))
                
                self.root.after(0, lambda: self.update_progress(90, lang.tr("install_ready_open")))
                url = f"http://localhost:{current_port}"
                def ask_open():
                    result = messagebox.askokcancel(
                        lang.tr("install_success_title"),
                        lang.tr("install_success", url),
                        parent=self.root
                    )
                    if result:
                        webbrowser.open(url)
                        self.root.after(0, lambda: self.append_result_text(lang.tr("install_page_opened", url), "success"))
                    else:
                        self.root.after(0, lambda: self.append_result_text(lang.tr("install_page_cancel"), "info"))
                self.root.after(0, ask_open)
                
                self.root.after(0, lambda: self.update_progress(95, lang.tr("install_enable_away")))
                if enable_away_mode():
                    self.root.after(0, lambda: self.append_result_text(lang.tr("install_away_enabled"), "success"))
                else:
                    self.root.after(0, lambda: self.append_result_text(lang.tr("install_away_failed"), "warning"))
                
                self.root.after(0, lambda: self.append_result_text(lang.tr("install_complete", url, self.SERVICE_NAME, self.PROCESS_NAME), "success"))
        except Exception as e:
            log.error(f"安装异常: {e}")
            import traceback
            log.error(f"异常详情: {traceback.format_exc()}")
            self.root.after(0, lambda: self.append_result_text(lang.tr("install_exception", str(e)), "danger"))
            self.root.after(0, lambda: self.update_progress(100, lang.tr("install_exception_short")))
        finally:
            self.install_btn.config(state=NORMAL)
            self.uninstall_btn.config(state=NORMAL)
            self.refresh_service_status()

    # ========== 卸载模块 ==========
    def start_uninstall(self):
        self.clear_result_text()
        self.append_result_text(lang.tr("uninstall_start"), "info")
        uninstall_thread = threading.Thread(target=self.uninstall_worker, daemon=True)
        uninstall_thread.start()

    def uninstall_worker(self):
        try:
            self.install_btn.config(state=DISABLED)
            self.uninstall_btn.config(state=DISABLED)
            
            self.root.after(0, lambda: self.update_progress(10, lang.tr("uninstall_check_admin")))
            if not check_admin_permission():
                self.root.after(0, lambda: self.update_progress(100, lang.tr("uninstall_admin_failed")))
                self.root.after(0, lambda: self.append_result_text(lang.tr("uninstall_admin_required"), "danger"))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                return
            
            self.root.after(0, lambda: self.update_progress(20, lang.tr("uninstall_stop_service")))
            if check_service_exists(self.SERVICE_NAME):
                stop_service(self.SERVICE_NAME)
                self.root.after(0, lambda: self.append_result_text(lang.tr("uninstall_stopped_service"), "success"))
            else:
                self.root.after(0, lambda: self.append_result_text(lang.tr("uninstall_no_service"), "info"))
            
            self.root.after(0, lambda: self.update_progress(30, lang.tr("uninstall_kill_process")))
            if check_process_exists(self.PROCESS_NAME):
                kill_process(self.PROCESS_NAME)
                self.root.after(0, lambda: self.append_result_text(lang.tr("uninstall_killed_process"), "success"))
            else:
                self.root.after(0, lambda: self.append_result_text(lang.tr("uninstall_no_process"), "info"))
            
            self.root.after(0, lambda: self.update_progress(40, lang.tr("uninstall_uninstall_service")))
            if os.path.exists(self.WINSW_EXE):
                run_cmd_with_retry([self.WINSW_EXE, "uninstall"], timeout=30)
                self.root.after(0, lambda: self.append_result_text(lang.tr("uninstall_service_uninstalled"), "success"))
            else:
                self.root.after(0, lambda: self.append_result_text(lang.tr("uninstall_winsw_missing"), "warning"))
            
            self.root.after(0, lambda: self.update_progress(50, lang.tr("uninstall_clean_firewall")))
            current_port = get_port_from_conf(self.CONF_PATH, self.DEFAULT_PORT)
            try:
                run_cmd_with_retry(['cmd', '/c', f'netsh advfirewall firewall delete rule name="Cloudreve_Port_Inbound"'], timeout=10)
                run_cmd_with_retry(['cmd', '/c', f'netsh advfirewall firewall delete rule name="Cloudreve_Port_Outbound"'], timeout=10)
                self.root.after(0, lambda: self.append_result_text(lang.tr("uninstall_firewall_cleaned"), "success"))
            except Exception as e:
                self.root.after(0, lambda: self.append_result_text(lang.tr("uninstall_firewall_failed", str(e)), "warning"))
            
            self.root.after(0, lambda: self.update_progress(60, lang.tr("uninstall_disable_away")))
            if disable_away_mode():
                self.root.after(0, lambda: self.append_result_text(lang.tr("uninstall_away_disabled"), "success"))
            else:
                self.root.after(0, lambda: self.append_result_text(lang.tr("uninstall_away_failed"), "warning"))
            
            self.root.after(0, lambda: self.update_progress(100, lang.tr("uninstall_complete")))
            self.root.after(0, lambda: self.append_result_text(lang.tr("uninstall_complete_msg"), "success"))
            self.root.after(0, lambda: messagebox.showinfo(lang.tr("uninstall_success_title"), lang.tr("uninstall_success"), parent=self.root))
        except Exception as e:
            log.error(f"卸载异常: {e}")
            import traceback
            log.error(f"异常详情: {traceback.format_exc()}")
            self.root.after(0, lambda: self.append_result_text(lang.tr("uninstall_exception", str(e)), "danger"))
            self.root.after(0, lambda: self.update_progress(100, lang.tr("uninstall_exception_short")))
            self.root.after(0, lambda: messagebox.showerror(lang.tr("error"), lang.tr("uninstall_exception", str(e)), parent=self.root))
        finally:
            self.install_btn.config(state=NORMAL)
            self.uninstall_btn.config(state=NORMAL)
            self.refresh_service_status()

    # ========== 启动服务（菜单调用） ==========
    def start_service_action(self):
        self.append_result_text(lang.tr("start_service_start"), "info")
        threading.Thread(target=self._start_service_worker, daemon=True).start()

    def _start_service_worker(self):
        try:
            self.root.after(0, lambda: self.update_progress(5, lang.tr("start_service_check")))
            ok, elapsed, err_msg = self._ensure_service_installed_and_started(start_after_install=True)
            if not ok:
                self.root.after(0, lambda: self.update_progress(100, lang.tr("start_service_failed")))
                self.root.after(0, lambda: self.append_result_text(lang.tr("start_service_failed_msg", err_msg), "danger"))
                self.root.after(0, lambda: messagebox.showerror(lang.tr("start_failed"), lang.tr("start_service_failed_msg", err_msg)))
                return
            else:
                self.root.after(0, lambda: self.append_result_text(lang.tr("start_service_ok", elapsed), "success"))
                self.root.after(0, lambda: self.update_progress(100, lang.tr("start_service_complete")))
                self.root.after(0, lambda: messagebox.showinfo(lang.tr("start_success_title"), lang.tr("start_service_success", elapsed), parent=self.root))
        except Exception as e:
            log.error(f"启动服务失败: {e}")
            self.root.after(0, lambda: self.append_result_text(lang.tr("start_service_exception", str(e)), "danger"))
            self.root.after(0, lambda: messagebox.showerror(lang.tr("start_failed"), lang.tr("start_service_exception", str(e))))
        finally:
            self.root.after(0, self.refresh_service_status)

    def _install_service_only(self):
        """仅安装服务（不启动）"""
        try:
            if check_process_exists(self.PROCESS_NAME):
                kill_process(self.PROCESS_NAME)
                self.root.after(0, lambda: self.append_result_text(lang.tr("install_service_kill_old"), "success"))
            run_cmd_with_retry([self.WINSW_EXE, "install"], timeout=30)
            log.info("服务安装完成")
            self.root.after(0, lambda: self.append_result_text(lang.tr("install_service_installed_ok"), "success"))
        except Exception as e:
            log.error(f"安装服务失败: {e}")
            self.root.after(0, lambda: self.append_result_text(lang.tr("install_service_failed", str(e)), "danger"))
            raise

    # ========== 停止服务（菜单调用） ==========
    def stop_service_action(self):
        self.append_result_text(lang.tr("stop_service_start"), "info")
        threading.Thread(target=self._stop_service_worker, daemon=True).start()

    def _stop_service_worker(self):
        try:
            self.root.after(0, lambda: self.update_progress(20, lang.tr("stop_service_stopping")))
            success = stop_service(self.SERVICE_NAME)
            if not success:
                self.root.after(0, lambda: self.append_result_text(lang.tr("stop_service_command_failed"), "danger"))
                self.root.after(0, lambda: messagebox.showerror(lang.tr("stop_failed"), lang.tr("stop_service_command_failed")))
                return

            self.root.after(0, lambda: self.update_progress(40, lang.tr("stop_service_wait")))
            for i in range(10):
                if get_service_status(self.SERVICE_NAME) != "RUNNING":
                    break
                progress = 40 + int(50 * (i + 1) / 10)
                self.root.after(0, self.update_progress, progress, lang.tr("stop_service_wait_progress", i+1, 10))
                time.sleep(1)
            
            self.root.after(0, lambda: self.append_result_text(lang.tr("stop_service_ok"), "success"))
            self.root.after(0, lambda: self.update_progress(100, lang.tr("stop_service_complete")))
            self.root.after(0, lambda: messagebox.showinfo(lang.tr("stop_success_title"), lang.tr("stop_service_success"), parent=self.root))
        except Exception as e:
            log.error(f"停止服务失败: {e}")
            self.root.after(0, lambda: self.append_result_text(lang.tr("stop_service_exception", str(e)), "danger"))
            self.root.after(0, lambda: messagebox.showerror(lang.tr("stop_failed"), lang.tr("stop_service_exception", str(e))))
        finally:
            self.root.after(0, self.refresh_service_status)

    # ========== 升级模块 ==========
    def start_upgrade(self):
        choice = messagebox.askquestion(
            lang.tr("upgrade_choice_title"),
            lang.tr("upgrade_choice_msg"),
            icon='question'
        )
        if choice == 'yes':
            self.clear_result_text()
            self.append_result_text(lang.tr("auto_upgrade_start"), "info")
            upgrade_thread = threading.Thread(target=self.auto_upgrade_worker, daemon=True)
            upgrade_thread.start()
        else:
            self.clear_result_text()
            self.append_result_text(lang.tr("manual_upgrade_start"), "info")
            upgrade_thread = threading.Thread(target=self.manual_upgrade_worker, daemon=True)
            upgrade_thread.start()

    def auto_upgrade_worker(self):
        try:
            self.install_btn.config(state=DISABLED)
            self.uninstall_btn.config(state=DISABLED)

            self.root.after(0, lambda: self.update_progress(5, lang.tr("auto_upgrade_check_version")))
            current_version = get_current_version(self.CLOUDREVE_EXE)
            if current_version:
                self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_current", current_version), "info"))
            else:
                self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_no_version"), "warning"))

            self.root.after(0, lambda: self.update_progress(10, lang.tr("auto_upgrade_fetch_latest")))
            try:
                latest_version = get_latest_version_from_github()
                self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_latest", latest_version), "info"))
            except Exception as e:
                self.root.after(0, lambda: self.update_progress(100, lang.tr("auto_upgrade_network_error")))
                self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_fetch_failed", str(e)), "danger"))
                self.root.after(0, lambda: messagebox.showerror(lang.tr("network_error"), lang.tr("auto_upgrade_network_error_msg", str(e))))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                return

            if current_version and not compare_versions(current_version, latest_version):
                self.root.after(0, lambda: self.update_progress(100, lang.tr("auto_upgrade_already_latest")))
                self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_already_latest_msg"), "success"))
                self.root.after(0, lambda: messagebox.showinfo(lang.tr("no_upgrade"), lang.tr("auto_upgrade_already_latest_msg", current_version)))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                return

            confirm = messagebox.askyesno(
                lang.tr("upgrade_confirm_title"),
                lang.tr("upgrade_confirm_msg", latest_version, current_version or '未知')
            )
            if not confirm:
                self.root.after(0, lambda: self.update_progress(100, lang.tr("upgrade_cancelled")))
                self.root.after(0, lambda: self.append_result_text(lang.tr("upgrade_cancelled_msg"), "warning"))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                return

            self.root.after(0, lambda: self.update_progress(20, lang.tr("auto_upgrade_get_url")))
            try:
                download_url = get_download_url_from_github()
                if not download_url:
                    raise RuntimeError("未找到适用于Windows的下载链接")
                self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_url_ok"), "info"))
            except Exception as e:
                self.root.after(0, lambda: self.update_progress(100, lang.tr("auto_upgrade_url_failed")))
                self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_url_failed_msg", str(e)), "danger"))
                self.root.after(0, lambda: messagebox.showerror(lang.tr("error"), lang.tr("auto_upgrade_url_failed_msg", str(e))))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                return

            self.root.after(0, lambda: self.update_progress(25, lang.tr("auto_upgrade_download")))
            os.makedirs(TEMP_DIR, exist_ok=True)
            zip_path = os.path.join(TEMP_DIR, "cloudreve_latest.zip")
            try:
                def download_progress(percent, downloaded, total):
                    self.root.after(0, self.update_progress, 25 + int(35 * percent / 100),
                                    lang.tr("auto_upgrade_download_progress", percent, downloaded, total))
                download_file(download_url, zip_path, download_progress)
                self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_download_ok"), "success"))
            except Exception as e:
                self.root.after(0, lambda: self.update_progress(100, lang.tr("auto_upgrade_download_failed")))
                self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_download_failed_msg", str(e)), "danger"))
                self.root.after(0, lambda: messagebox.showerror(lang.tr("download_failed"), lang.tr("auto_upgrade_download_failed_msg", str(e))))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                return

            self.root.after(0, lambda: self.update_progress(60, lang.tr("auto_upgrade_extract")))
            extract_dir = os.path.join(TEMP_DIR, "extract")
            try:
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    zip_ref.extractall(extract_dir)
                self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_extract_ok"), "success"))
            except Exception as e:
                self.root.after(0, lambda: self.update_progress(100, lang.tr("auto_upgrade_extract_failed")))
                self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_extract_failed_msg", str(e)), "danger"))
                self.root.after(0, lambda: messagebox.showerror(lang.tr("extract_failed"), lang.tr("auto_upgrade_extract_failed_msg", str(e))))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
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
                self.root.after(0, lambda: self.update_progress(100, lang.tr("auto_upgrade_no_exe")))
                self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_no_exe_msg"), "danger"))
                self.root.after(0, lambda: messagebox.showerror(lang.tr("file_missing"), lang.tr("auto_upgrade_no_exe_msg")))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                return

            service_exists_before = check_service_exists(self.SERVICE_NAME)
            self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_service_before", "存在" if service_exists_before else "不存在"), "info"))

            if service_exists_before:
                self.root.after(0, lambda: self.update_progress(65, lang.tr("auto_upgrade_stop_service")))
                if check_service_exists(self.SERVICE_NAME):
                    stop_service(self.SERVICE_NAME)
                    self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_stopped"), "success"))
                else:
                    self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_not_running"), "info"))

                self.root.after(0, lambda: self.update_progress(70, lang.tr("auto_upgrade_kill_process")))
                if check_process_exists(self.PROCESS_NAME):
                    kill_process(self.PROCESS_NAME)
                    self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_killed"), "success"))
                else:
                    self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_no_process"), "info"))
            else:
                self.root.after(0, lambda: self.update_progress(65, lang.tr("auto_upgrade_check_residual")))
                if check_process_exists(self.PROCESS_NAME):
                    kill_process(self.PROCESS_NAME)
                    self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_killed_residual"), "success"))
                else:
                    self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_no_residual"), "info"))

            self.root.after(0, lambda: self.update_progress(75, lang.tr("auto_upgrade_backup")))
            backup_path = self.CLOUDREVE_EXE + ".bak"
            if os.path.exists(self.CLOUDREVE_EXE):
                try:
                    shutil.copy2(self.CLOUDREVE_EXE, backup_path)
                    self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_backup_ok", backup_path), "success"))
                except Exception as e:
                    self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_backup_failed", str(e)), "warning"))

            self.root.after(0, lambda: self.update_progress(80, lang.tr("auto_upgrade_replace")))
            try:
                shutil.copy2(new_exe_path, self.CLOUDREVE_EXE)
                self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_replace_ok", self.CLOUDREVE_EXE), "success"))
                if os.path.exists(self.CLOUDREVE_EXE):
                    mod_time = os.path.getmtime(self.CLOUDREVE_EXE)
                    mod_time_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mod_time))
                    self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_file_date", mod_time_str), "info"))
            except Exception as e:
                self.root.after(0, lambda: self.update_progress(100, lang.tr("auto_upgrade_replace_failed")))
                self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_replace_failed_msg", str(e)), "danger"))
                self.root.after(0, lambda: messagebox.showerror(lang.tr("replace_failed"), lang.tr("auto_upgrade_replace_failed_msg", str(e))))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                return

            if service_exists_before:
                self.root.after(0, lambda: self.update_progress(85, lang.tr("auto_upgrade_start_service")))
                # 使用统一方法确保服务启动
                ok, elapsed, err_msg = self._ensure_service_installed_and_started(start_after_install=True)
                if not ok:
                    self.root.after(0, lambda: self.update_progress(100, lang.tr("auto_upgrade_start_failed")))
                    self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_start_failed_msg", err_msg), "warning"))
                    self.root.after(0, lambda: messagebox.showwarning(lang.tr("start_failed"), lang.tr("auto_upgrade_start_failed_msg", err_msg)))
                    self.install_btn.config(state=NORMAL)
                    self.uninstall_btn.config(state=NORMAL)
                    return
                else:
                    self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_start_ok", elapsed), "success"))

                self.root.after(0, lambda: self.update_progress(93, lang.tr("auto_upgrade_wait_port")))
                current_port = get_port_from_conf(self.CONF_PATH, self.DEFAULT_PORT)
                try:
                    wait_for_port_open('localhost', current_port, max_wait=30)
                    self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_port_ok", current_port), "success"))
                except TimeoutError:
                    self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_port_timeout", current_port), "warning"))

                self.root.after(0, lambda: self.update_progress(96, lang.tr("auto_upgrade_open_browser")))
                url = f"http://localhost:{current_port}"
                try:
                    webbrowser.open(url)
                    self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_browser_ok", url), "success"))
                except Exception as e:
                    self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_browser_failed", url, str(e)), "warning"))

                self.root.after(0, lambda: self.update_progress(100, lang.tr("auto_upgrade_complete")))
                if os.path.exists(self.CLOUDREVE_EXE):
                    mod_time = os.path.getmtime(self.CLOUDREVE_EXE)
                    mod_time_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mod_time))
                else:
                    mod_time_str = "未知"
                self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_complete_msg", latest_version, url, mod_time_str, backup_path), "success"))
            else:
                self.root.after(0, lambda: self.update_progress(100, lang.tr("auto_upgrade_complete_no_service")))
                if os.path.exists(self.CLOUDREVE_EXE):
                    mod_time = os.path.getmtime(self.CLOUDREVE_EXE)
                    mod_time_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mod_time))
                else:
                    mod_time_str = "未知"
                self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_complete_no_service_msg", mod_time_str, backup_path), "success"))

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
            self.root.after(0, lambda: self.append_result_text(lang.tr("auto_upgrade_exception", str(e)), "danger"))
            self.root.after(0, lambda: self.update_progress(100, lang.tr("auto_upgrade_exception_short")))
            self.root.after(0, lambda: messagebox.showerror(lang.tr("error"), lang.tr("auto_upgrade_exception", str(e))))
        finally:
            self.install_btn.config(state=NORMAL)
            self.uninstall_btn.config(state=NORMAL)
            self.refresh_service_status()

    def manual_upgrade_worker(self):
        try:
            self.install_btn.config(state=DISABLED)
            self.uninstall_btn.config(state=DISABLED)

            self.root.after(0, lambda: self.update_progress(10, lang.tr("manual_upgrade_check_admin")))
            if not check_admin_permission():
                self.root.after(0, lambda: self.update_progress(100, lang.tr("manual_upgrade_admin_failed")))
                self.root.after(0, lambda: self.append_result_text(lang.tr("manual_upgrade_admin_required"), "danger"))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                return

            self.root.after(0, lambda: self.update_progress(15, lang.tr("manual_upgrade_check_service")))
            service_exists_before = check_service_exists(self.SERVICE_NAME)
            self.root.after(0, lambda: self.append_result_text(lang.tr("manual_upgrade_service_before", "存在" if service_exists_before else "不存在"), "info"))
            if not service_exists_before:
                self.root.after(0, lambda: self.append_result_text(lang.tr("manual_upgrade_no_service_warning"), "warning"))

            self.root.after(0, lambda: self.update_progress(20, lang.tr("manual_upgrade_select_file")))
            new_exe_path = filedialog.askopenfilename(
                title=lang.tr("manual_upgrade_select_title"),
                initialfile="cloudreve.exe",
                filetypes=[("可执行文件", "*.exe"), ("所有文件", "*.*")]
            )
            if not new_exe_path:
                self.root.after(0, lambda: self.update_progress(100, lang.tr("manual_upgrade_cancelled")))
                self.root.after(0, lambda: self.append_result_text(lang.tr("manual_upgrade_cancelled_msg"), "warning"))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                return

            selected_filename = os.path.basename(new_exe_path)
            if selected_filename.lower() != "cloudreve.exe":
                self.root.after(0, lambda: self.update_progress(100, lang.tr("manual_upgrade_wrong_name")))
                self.root.after(0, lambda: self.append_result_text(lang.tr("manual_upgrade_wrong_name_msg", selected_filename), "danger"))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                return

            if not os.path.isfile(new_exe_path):
                self.root.after(0, lambda: self.update_progress(100, lang.tr("manual_upgrade_invalid_file")))
                self.root.after(0, lambda: self.append_result_text(lang.tr("manual_upgrade_invalid_file_msg"), "danger"))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                return

            self.root.after(0, lambda: self.append_result_text(lang.tr("manual_upgrade_selected", new_exe_path), "info"))

            current_port = get_port_from_conf(self.CONF_PATH, self.DEFAULT_PORT)
            self.root.after(0, lambda: self.append_result_text(lang.tr("manual_upgrade_port", current_port), "info"))

            if service_exists_before:
                self.root.after(0, lambda: self.update_progress(30, lang.tr("manual_upgrade_stop_service")))
                if check_service_exists(self.SERVICE_NAME):
                    stop_service(self.SERVICE_NAME)
                    self.root.after(0, lambda: self.append_result_text(lang.tr("manual_upgrade_stopped"), "success"))
                else:
                    self.root.after(0, lambda: self.append_result_text(lang.tr("manual_upgrade_not_running"), "info"))

                self.root.after(0, lambda: self.update_progress(40, lang.tr("manual_upgrade_kill_process")))
                if check_process_exists(self.PROCESS_NAME):
                    kill_process(self.PROCESS_NAME)
                    self.root.after(0, lambda: self.append_result_text(lang.tr("manual_upgrade_killed"), "success"))
                else:
                    self.root.after(0, lambda: self.append_result_text(lang.tr("manual_upgrade_no_process"), "info"))
            else:
                self.root.after(0, lambda: self.update_progress(30, lang.tr("manual_upgrade_check_residual")))
                if check_process_exists(self.PROCESS_NAME):
                    kill_process(self.PROCESS_NAME)
                    self.root.after(0, lambda: self.append_result_text(lang.tr("manual_upgrade_killed_residual"), "success"))
                else:
                    self.root.after(0, lambda: self.append_result_text(lang.tr("manual_upgrade_no_residual"), "info"))

            self.root.after(0, lambda: self.update_progress(50, lang.tr("manual_upgrade_backup")))
            backup_path = self.CLOUDREVE_EXE + ".bak"
            if os.path.exists(self.CLOUDREVE_EXE):
                try:
                    shutil.copy2(self.CLOUDREVE_EXE, backup_path)
                    self.root.after(0, lambda: self.append_result_text(lang.tr("manual_upgrade_backup_ok", backup_path), "success"))
                except Exception as e:
                    self.root.after(0, lambda: self.append_result_text(lang.tr("manual_upgrade_backup_failed", str(e)), "warning"))
            else:
                self.root.after(0, lambda: self.append_result_text(lang.tr("manual_upgrade_no_backup"), "info"))

            self.root.after(0, lambda: self.update_progress(60, lang.tr("manual_upgrade_replace")))
            try:
                shutil.copy2(new_exe_path, self.CLOUDREVE_EXE)
                self.root.after(0, lambda: self.append_result_text(lang.tr("manual_upgrade_replace_ok", self.CLOUDREVE_EXE), "success"))
                if os.path.exists(self.CLOUDREVE_EXE):
                    mod_time = os.path.getmtime(self.CLOUDREVE_EXE)
                    mod_time_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mod_time))
                    self.root.after(0, lambda: self.append_result_text(lang.tr("manual_upgrade_file_date", mod_time_str), "info"))
            except Exception as e:
                self.root.after(0, lambda: self.update_progress(100, lang.tr("manual_upgrade_replace_failed")))
                self.root.after(0, lambda: self.append_result_text(lang.tr("manual_upgrade_replace_failed_msg", str(e)), "danger"))
                self.install_btn.config(state=NORMAL)
                self.uninstall_btn.config(state=NORMAL)
                return

            if service_exists_before:
                self.root.after(0, lambda: self.update_progress(70, lang.tr("manual_upgrade_start_service")))
                ok, elapsed, err_msg = self._ensure_service_installed_and_started(start_after_install=True)
                if not ok:
                    self.root.after(0, lambda: self.update_progress(100, lang.tr("manual_upgrade_start_failed")))
                    self.root.after(0, lambda: self.append_result_text(lang.tr("manual_upgrade_start_failed_msg", elapsed, err_msg), "warning"))
                    self.install_btn.config(state=NORMAL)
                    self.uninstall_btn.config(state=NORMAL)
                    return
                else:
                    self.root.after(0, lambda: self.append_result_text(lang.tr("manual_upgrade_start_ok", elapsed), "success"))

                self.root.after(0, lambda: self.update_progress(90, lang.tr("manual_upgrade_wait_port")))
                def port_progress(elapsed, max_wait, remaining):
                    self.root.after(0, self.update_progress, 90 + int(5 * (elapsed / max_wait)), lang.tr("manual_upgrade_port_progress", elapsed, max_wait))
                try:
                    wait_for_port_open('localhost', current_port, max_wait=30, progress_callback=port_progress)
                    self.root.after(0, lambda: self.append_result_text(lang.tr("manual_upgrade_port_ok", current_port), "success"))
                except TimeoutError:
                    self.root.after(0, lambda: self.append_result_text(lang.tr("manual_upgrade_port_timeout", current_port), "warning"))

                self.root.after(0, lambda: self.update_progress(95, lang.tr("manual_upgrade_open_browser")))
                url = f"http://localhost:{current_port}"
                try:
                    webbrowser.open(url)
                    self.root.after(0, lambda: self.append_result_text(lang.tr("manual_upgrade_browser_ok", url), "success"))
                except Exception as e:
                    self.root.after(0, lambda: self.append_result_text(lang.tr("manual_upgrade_browser_failed", url, str(e)), "warning"))

                self.root.after(0, lambda: self.update_progress(100, lang.tr("manual_upgrade_complete")))
                if os.path.exists(self.CLOUDREVE_EXE):
                    mod_time = os.path.getmtime(self.CLOUDREVE_EXE)
                    mod_time_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mod_time))
                else:
                    mod_time_str = "未知"
                self.root.after(0, lambda: self.append_result_text(lang.tr("manual_upgrade_complete_msg", url, mod_time_str, backup_path), "success"))
            else:
                self.root.after(0, lambda: self.update_progress(100, lang.tr("manual_upgrade_complete_no_service")))
                if os.path.exists(self.CLOUDREVE_EXE):
                    mod_time = os.path.getmtime(self.CLOUDREVE_EXE)
                    mod_time_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mod_time))
                else:
                    mod_time_str = "未知"
                self.root.after(0, lambda: self.append_result_text(lang.tr("manual_upgrade_complete_no_service_msg", mod_time_str, backup_path), "success"))

        except Exception as e:
            log.error(f"手动升级异常: {e}")
            import traceback
            log.error(f"异常详情: {traceback.format_exc()}")
            self.root.after(0, lambda: self.append_result_text(lang.tr("manual_upgrade_exception", str(e)), "danger"))
            self.root.after(0, lambda: self.update_progress(100, lang.tr("manual_upgrade_exception_short")))
            self.root.after(0, lambda: messagebox.showerror(lang.tr("error"), lang.tr("manual_upgrade_exception", str(e))))
        finally:
            self.install_btn.config(state=NORMAL)
            self.uninstall_btn.config(state=NORMAL)
            self.refresh_service_status()


if __name__ == "__main__":
    # 加载保存的语言设置
    saved_lang = None
    try:
        config_file = os.path.join(os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__), "settings.ini")
        config = configparser.ConfigParser()
        config.read(config_file, encoding='utf-8')
        if config.has_section("Language") and config.has_option("Language", "current"):
            saved_lang = config.get("Language", "current")
    except:
        pass

    if saved_lang:
        lang.set_language(saved_lang)
    else:
        # 没有保存的语言设置，尝试检测系统语言
        try:
            sys_lang, _ = locale.getdefaultlocale()
            if sys_lang:
                # 转换格式，如 zh_CN -> zh_CN, en_US -> en_US
                if sys_lang in ["zh_CN", "zh_TW", "zh_HK"]:
                    lang.set_language("zh_CN")
                elif sys_lang.startswith("en"):
                    lang.set_language("en_US")
                # 其他语言可继续添加
        except Exception as e:
            log.warning(f"检测系统语言失败: {e}")

    run_as_admin()
    root = ttk.Window(themename="flatly")
    app = CloudreveManagerGUI(root)
    root.mainloop()