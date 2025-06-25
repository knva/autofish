import time
import numpy as np
import sounddevice as sd
import win32gui
import win32con
import win32process
import psutil

# 音量阈值（可根据实际情况调整）
VOLUME_THRESHOLD = 0.08  # 响度阈值，0~1，建议先用 print 看实际值再调整
# 检查音量的采样间隔（秒）
CHECK_INTERVAL = 0.1
# 按键间隔（秒）
KEY_DELAY = 1.0
TIMEOUT = 30  # 超时时间，秒

# 虚拟键码
VK_X = 0x58  # 'X'
VK_Z = 0x5A  # 'Z'
WOW_EXE = 'wow.exe'

# 查找 wow.exe 主窗口句柄
def find_wow_hwnd():
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] and proc.info['name'].lower() == WOW_EXE:
            pid = proc.info['pid']
            def callback(hwnd, hwnds):
                _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
                if found_pid == pid and win32gui.IsWindowVisible(hwnd):
                    hwnds.append(hwnd)
                return True
            hwnds = []
            win32gui.EnumWindows(callback, hwnds)
            if hwnds:
                return hwnds[0]
    return None

# 发送按键消息到 wow.exe
def post_key(hwnd, vk_code):
    win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk_code, 0)
    time.sleep(0.05)
    win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk_code, 0)

last_trigger_time = time.time()

def audio_callback(indata, frames, time_info, status):
    global last_trigger_time
    volume_norm = np.linalg.norm(indata) / np.sqrt(len(indata))
    print(volume_norm)
    if volume_norm > VOLUME_THRESHOLD:
        hwnd = find_wow_hwnd()
        if hwnd:
            print(f"大音量触发: {volume_norm:.3f}")
            post_key(hwnd, VK_X)
            time.sleep(KEY_DELAY)
            post_key(hwnd, VK_Z)
            time.sleep(1)
            last_trigger_time = time.time()
        else:
            print("未找到 wow.exe 窗口")
            last_trigger_time = time.time()

if __name__ == "__main__":
    print("开始监听麦克风音量，超过阈值将依次向 wow.exe 发送 X 和 Z 键... 超过30秒未触发则自动发送Z")
    last_trigger_time = time.time()
    with sd.InputStream(callback=audio_callback, channels=1, samplerate=16000, blocksize=int(16000*CHECK_INTERVAL)):
        while True:
            if time.time() - last_trigger_time > TIMEOUT:
                hwnd = find_wow_hwnd()
                if hwnd:
                    print(f"超时{TIMEOUT}秒未检测到大音量，自动发送Z")
                    post_key(hwnd, VK_Z)
                    last_trigger_time = time.time()
                else:
                    print("超时未检测到大音量，且未找到 wow.exe 窗口")
                    last_trigger_time = time.time()
            time.sleep(1) 