import time
import numpy as np
import sounddevice as sd
import win32gui
import win32con
import win32process
import psutil
import tkinter as tk
from tkinter import ttk, scrolledtext
import threading
import queue
import random
import json
import os

# --- Constants and Configs ---
WOW_EXE = 'wow.exe'
CONFIG_FILE = 'audio_bot_config.json'
VK_CODE_MAP = {
    'a': 0x41, 'b': 0x42, 'c': 0x43, 'd': 0x44, 'e': 0x45, 'f': 0x46,
    'g': 0x47, 'h': 0x48, 'i': 0x49, 'j': 0x4A, 'k': 0x4B, 'l': 0x4C,
    'm': 0x4D, 'n': 0x4E, 'o': 0x4F, 'p': 0x50, 'q': 0x51, 'r': 0x52,
    's': 0x53, 't': 0x54, 'u': 0x55, 'v': 0x56, 'w': 0x57, 'x': 0x58,
    'y': 0x59, 'z': 0x5A,
    '0': 0x30, '1': 0x31, '2': 0x32, '3': 0x33, '4': 0x34,
    '5': 0x35, '6': 0x36, '7': 0x37, '8': 0x38, '9': 0x39,
}

# --- Helper Functions ---
def find_wow_hwnd():
    for proc in psutil.process_iter(['pid', 'name']):
        try:
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
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return None

def post_key(hwnd, vk_code):
    win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk_code, 0)
    time.sleep(0.05)
    win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk_code, 0)

# --- Main Application Class ---
class AudioBotGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("音频按键助手 Pro")
        self.master.geometry("550x620")
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

        # State variables
        self.is_running = False
        self.is_acting = False
        self.worker_thread = None
        self.log_queue = queue.Queue()
        self.last_trigger_time = time.time()
        self.audio_stream = None
        
        # UI Variables
        self.threshold_var = tk.DoubleVar(value=0.08)
        self.reel_in_key_var = tk.StringVar(value='x')
        self.cast_rod_key_var = tk.StringVar(value='z')
        self.timeout_var = tk.IntVar(value=30)
        self.delay_min_var = tk.DoubleVar(value=1.5)
        self.delay_max_var = tk.DoubleVar(value=2.5)
        self.device_var = tk.StringVar()
        self.status_var = tk.StringVar(value="状态: 已停止")

        # Load existing config
        self.load_config()
        
        # UI Setup
        self.create_widgets()
        self.process_log_queue()
        self.refresh_devices()

    def create_widgets(self):
        main_frame = ttk.Frame(self.master, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Settings Frame ---
        settings_frame = ttk.LabelFrame(main_frame, text="配置参数", padding="10")
        settings_frame.pack(fill=tk.X, pady=(0, 10))
        settings_frame.grid_columnconfigure(1, weight=1)

        # Device Selection
        ttk.Label(settings_frame, text="声卡设备:").grid(row=0, column=0, sticky=tk.W, pady=4)
        self.device_combo = ttk.Combobox(settings_frame, textvariable=self.device_var, state="readonly")
        self.device_combo.grid(row=0, column=1, sticky=tk.EW, pady=4)
        ttk.Button(settings_frame, text="刷新", width=5, command=self.refresh_devices).grid(row=0, column=2, padx=5)

        # Threshold
        ttk.Label(settings_frame, text="响度阈值:").grid(row=1, column=0, sticky=tk.W, pady=4)
        thresh_frame = ttk.Frame(settings_frame)
        thresh_frame.grid(row=1, column=1, columnspan=2, sticky=tk.EW)
        ttk.Scale(thresh_frame, from_=0.01, to=0.5, orient=tk.HORIZONTAL, variable=self.threshold_var).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(thresh_frame, textvariable=self.threshold_var, width=5).pack(side=tk.RIGHT)

        # Keys
        ttk.Label(settings_frame, text="收杆按键:").grid(row=2, column=0, sticky=tk.W, pady=4)
        ttk.Entry(settings_frame, textvariable=self.reel_in_key_var).grid(row=2, column=1, columnspan=2, sticky=tk.EW)

        ttk.Label(settings_frame, text="抛竿按键:").grid(row=3, column=0, sticky=tk.W, pady=4)
        ttk.Entry(settings_frame, textvariable=self.cast_rod_key_var).grid(row=3, column=1, columnspan=2, sticky=tk.EW)
        
        ttk.Label(settings_frame, text="超时重抛(秒):").grid(row=4, column=0, sticky=tk.W, pady=4)
        ttk.Entry(settings_frame, textvariable=self.timeout_var).grid(row=4, column=1, columnspan=2, sticky=tk.EW)

        # Delay Settings
        ttk.Label(settings_frame, text="收杆后等待(秒):").grid(row=5, column=0, sticky=tk.W, pady=4)
        delay_frame = ttk.Frame(settings_frame)
        delay_frame.grid(row=5, column=1, columnspan=2, sticky=tk.EW)
        ttk.Label(delay_frame, text="最小:").pack(side=tk.LEFT)
        ttk.Entry(delay_frame, textvariable=self.delay_min_var, width=8).pack(side=tk.LEFT, padx=5)
        ttk.Label(delay_frame, text="最大:").pack(side=tk.LEFT, padx=(10, 0))
        ttk.Entry(delay_frame, textvariable=self.delay_max_var, width=8).pack(side=tk.LEFT, padx=5)

        # --- Actions Frame ---
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X, pady=5)
        
        self.start_button = ttk.Button(action_frame, text="▶ 启动监听", command=self.start_worker)
        self.start_button.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        
        self.stop_button = ttk.Button(action_frame, text="■ 停止", command=self.stop_worker, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill=tk.X, pady=5)

        # --- Log Frame ---
        log_frame = ttk.LabelFrame(main_frame, text="运行日志", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=8, font=('Consolas', 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.configure(state='disabled')

    def refresh_devices(self):
        devices = sd.query_devices()
        input_devices = []
        for i, d in enumerate(devices):
            if d['max_input_channels'] > 0:
                input_devices.append(f"{i}: {d['name']}")
        
        self.device_combo['values'] = input_devices
        if input_devices:
            if self.device_var.get() not in input_devices:
                # Default to system default input
                default_idx = sd.default.device[0]
                for d_str in input_devices:
                    if d_str.startswith(f"{default_idx}:"):
                        self.device_var.set(d_str)
                        break
                else:
                    self.device_var.set(input_devices[0])

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.threshold_var.set(config.get('threshold', 0.08))
                    self.reel_in_key_var.set(config.get('reel_key', 'x'))
                    self.cast_rod_key_var.set(config.get('cast_key', 'z'))
                    self.timeout_var.set(config.get('timeout', 30))
                    self.delay_min_var.set(config.get('delay_min', 1.5))
                    self.delay_max_var.set(config.get('delay_max', 2.5))
                    self.device_var.set(config.get('device', ''))
            except Exception as e:
                print(f"加载配置失败: {e}")

    def save_config(self):
        config = {
            'threshold': round(self.threshold_var.get(), 3),
            'reel_key': self.reel_in_key_var.get(),
            'cast_key': self.cast_rod_key_var.get(),
            'timeout': self.timeout_var.get(),
            'delay_min': self.delay_min_var.get(),
            'delay_max': self.delay_max_var.get(),
            'device': self.device_var.get()
        }
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
        except Exception as e:
            self.log_message(f"保存配置失败: {e}")

    def log_message(self, message):
        self.log_queue.put(message)

    def process_log_queue(self):
        try:
            while True:
                message = self.log_queue.get_nowait()
                self.log_text.configure(state='normal')
                self.log_text.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {message}\n")
                self.log_text.configure(state='disabled')
                self.log_text.see(tk.END)
        except queue.Empty:
            pass
        finally:
            self.master.after(100, self.process_log_queue)

    def get_vk_code(self, key_str):
        return VK_CODE_MAP.get(key_str.lower().strip())

    def start_worker(self):
        if self.is_running: return
        
        # Validation
        if not self.get_vk_code(self.reel_in_key_var.get()) or not self.get_vk_code(self.cast_rod_key_var.get()):
            self.log_message("错误：按键配置无效，请输入 a-z 或 0-9")
            return
            
        self.save_config()
        self.is_running = True
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.status_var.set("状态: 正在监听音频...")
        
        self.worker_thread = threading.Thread(target=self._monitor_loop, daemon=True)
        self.worker_thread.start()

        # Initial Cast
        threading.Thread(target=self._perform_initial_cast, daemon=True).start()

    def stop_worker(self):
        if not self.is_running: return
        self.is_running = False
        if self.audio_stream:
            try:
                self.audio_stream.stop()
                self.audio_stream.close()
            except:
                pass
            self.audio_stream = None
            
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.status_var.set("状态: 已停止")
        self.log_message("停止监听服务。")

    def _perform_initial_cast(self):
        time.sleep(1.0) # 给用户一点切窗口的时间
        hwnd = find_wow_hwnd()
        if hwnd:
            self.log_message("执行启动抛竿...")
            post_key(hwnd, self.get_vk_code(self.cast_rod_key_var.get()))
            self.last_trigger_time = time.time()
        else:
            self.log_message("警告：未发现游戏窗口，请确保游戏已运行。")

    def _audio_callback(self, indata, frames, time_info, status):
        if not self.is_running or self.is_acting:
            return
        
        # Calculate RMS volume
        volume_norm = np.linalg.norm(indata) / np.sqrt(len(indata))
        if volume_norm > self.threshold_var.get():
            self.is_acting = True
            threading.Thread(target=self._perform_reel_and_cast, daemon=True).start()

    def _perform_reel_and_cast(self):
        try:
            if not self.is_running: return
            
            hwnd = find_wow_hwnd()
            if not hwnd:
                self.log_message("未找到游戏窗口，跳过本次触发。")
                return

            self.log_message(f"触发！音量满足阈值。正在收杆...")
            post_key(hwnd, self.get_vk_code(self.reel_in_key_var.get()))
            
            # 使用用户设置的随机延迟范围
            d_min = self.delay_min_var.get()
            d_max = self.delay_max_var.get()
            delay = random.uniform(min(d_min, d_max), max(d_min, d_max))
            
            self.log_message(f"等待 {delay:.2f}s 后重新抛竿...")
            time.sleep(delay)
            
            if self.is_running:
                post_key(hwnd, self.get_vk_code(self.cast_rod_key_var.get()))
                self.last_trigger_time = time.time()
                self.log_message("抛竿完成。")

        except Exception as e:
            self.log_message(f"执行动作异常: {e}")
        finally:
            self.is_acting = False

    def _monitor_loop(self):
        try:
            # Parse device index
            device_str = self.device_var.get()
            device_id = int(device_str.split(':')[0]) if device_str else None
            
            self.audio_stream = sd.InputStream(
                device=device_id,
                channels=1,
                samplerate=16000,
                blocksize=1600,
                callback=self._audio_callback
            )
            
            with self.audio_stream:
                self.log_message(f"音频流已建立。设备: {device_str}")
                while self.is_running:
                    # Timeout check
                    elapsed = time.time() - self.last_trigger_time
                    if elapsed > self.timeout_var.get():
                        self.log_message(f"超时重试：距离上次动作已过 {int(elapsed)}s")
                        hwnd = find_wow_hwnd()
                        if hwnd:
                            post_key(hwnd, self.get_vk_code(self.cast_rod_key_var.get()))
                        self.last_trigger_time = time.time()
                    
                    time.sleep(0.5)
                    
        except Exception as e:
            self.log_message(f"监听循环崩溃: {e}")
            self.master.after(0, self.stop_worker)

    def on_closing(self):
        self.save_config()
        self.stop_worker()
        self.master.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    # Simple style enhancement
    style = ttk.Style()
    style.theme_use('clam')
    app = AudioBotGUI(root)
    root.mainloop()
