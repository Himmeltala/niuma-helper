import datetime
import psutil
import sys
import os

QUIT_TIME = "17:30"
TARGET_PROCESS_NAMES = ['WeChatAppEx.exe', 'Weixin.exe']

current_date = datetime.datetime.now().strftime("%Y-%m-%d")
log_root_path = "E:/项目/Python/日报周报自动化/logs"
log_dir_path = os.path.join(log_root_path, current_date)
os.makedirs(log_dir_path, exist_ok=True)
log_file_path = os.path.join(log_dir_path, "auto_close_wechat.log")

try:
    log_file = open(log_file_path, 'a', encoding='utf-8')
    sys.stdout = log_file
except FileNotFoundError:
    sys.exit(1)


def get_quit_datetime():
    now = datetime.datetime.now()
    quit_time = datetime.datetime.strptime(QUIT_TIME, "%H:%M").time()
    return datetime.datetime.combine(now.date(), quit_time)


def kill_wechat(now=''):
    print(f"[info] {now}: 尝试终止进程")
    for p in psutil.process_iter(['name']):
        if p.info['name'] and any(target in p.info['name'] for target in TARGET_PROCESS_NAMES):
            try:
                p.kill()
                print(f"[info] {now}: 终止进程 {p.info['name']}")
            except psutil.NoSuchProcess:
                print(f"[error] {now}: 进程 {p.info['name']} 在尝试终止时已不存在")
            except Exception as e:
                print(f"[error] {now}: 终止进程 {p.info['name']} 时出错: {e}")


now = datetime.datetime.now()
quit_datetime = get_quit_datetime()
if quit_datetime <= now:
    kill_wechat(now.strftime("%Y-%m-%d %H:%M:%S"))


sys.stdout = sys.__stdout__
log_file.close()
