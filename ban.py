import sys
import subprocess
import threading
import os

def install_package(package_name):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])

for package in ["pystray", "Pillow", "psutil", "datetime", "time", "winshell", "PyWin32"]:
    try:
        __import__(package)
    except ImportError:
        print(f"缺少包{package}，正在安装")
        install_package(package)

import psutil
from PIL import Image
from pystray import Icon as icon, Menu as menu, MenuItem as item
import datetime
import time
stop_event = threading.Event()
print("程序启动")

def set_autostart(autostart, script_path):
    startup_dir = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
    shortcut_path = os.path.join(startup_dir, 'BanScript.lnk')

    if autostart.lower() == 'yes':
        create_shortcut(sys.executable, shortcut_path)
    else:
        if os.path.exists(shortcut_path):
            os.remove(shortcut_path)

def create_shortcut(target_path, shortcut_path):
    vbscript = f"""
    Set oWS = WScript.CreateObject("WScript.Shell")
    sLinkFile = "{shortcut_path}"
    Set oLink = oWS.CreateShortcut(sLinkFile)
    oLink.TargetPath = "{target_path}"
    oLink.Save
    """

    vbs_path = "create_shortcut.vbs"
    with open(vbs_path, "w") as file:
        file.write(vbscript)

    os.system(f'cscript "{vbs_path}"')
    os.remove(vbs_path)

def read_config(filename):
    try:
        with open(filename, 'r', encoding='utf-8') as file:
            content = file.read()
    except:
        try:
            with open(filename, 'r', encoding='gbk') as file:
                content = file.read()
        except:
            print("config文件编码错误")
            sys.exit()

    content = [line.strip() for line in content.split('\n') if not line.strip().startswith('#')]
    autostart = content[0].strip().split('=')[1]
    processes = []
    time_spans = []
    section = None

    for line in content:
        if line == "{":
            section = []
        elif line == "}":
            if processes:
                time_spans = section
            else:
                processes = section
            section = None
        elif section is not None:
            section.append(line)

    valid_time_spans = []
    for span in time_spans:
        times = span.strip('[]').split('-')
        if times[0] != times[1]:
            valid_time_spans.append((times[0], times[1]))

    return processes, valid_time_spans, autostart


def is_time_in_span(current_time, time_span):
    start_time = datetime.datetime.strptime(time_span[0], '%H:%M:%S').time()
    end_time = datetime.datetime.strptime(time_span[1], '%H:%M:%S').time()

    if start_time == end_time:
        return False

    if start_time > end_time:
        return current_time >= start_time or current_time <= end_time
    else:
        return start_time <= current_time <= end_time


def terminate_processes(process_names):
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] in process_names:
            try:
                psutil.Process(proc.info['pid']).terminate()
            except psutil.NoSuchProcess:
                pass  

def check_processes(icon, processes, time_spans):
    while not stop_event.is_set():
        current_time = datetime.datetime.now().time()
        in_span = any(is_time_in_span(current_time, span) for span in time_spans)
        icon.title = f"当前时间: {current_time.strftime('%H:%M:%S')} \n封禁列表: {', '.join(processes)} \n封禁时间段: {', '.join(['-'.join(span) for span in time_spans])} \n{'正在封禁' if in_span else '非封禁时间'}"
        if in_span:
            terminate_processes(processes)
        time.sleep(1)

def create_icon(image_path):
    return Image.open(image_path)

def exit_action(icon, item):
    stop_event.set()
    icon.stop()

if __name__ == '__main__':
    processes, time_spans, autostart = read_config('banlist.config')
    set_autostart(autostart, os.path.abspath('ban.py'))
    icon_image = create_icon(os.path.join(sys.path[0], "favicon.png"))

    tray_icon = icon("Process Terminator", icon_image, menu=menu(
        item('Exit', lambda icon, item: exit_action(icon, item))
    ))

    process_thread = threading.Thread(target=check_processes, args=(tray_icon, processes, time_spans))
    process_thread.start()

    tray_icon.run()
