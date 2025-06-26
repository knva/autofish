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

# --- Constants and Configs ---
WOW_EXE = 'wow.exe'
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

def post_key(hwnd, vk_code):
    win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk_code, 0)
    time.sleep(0.05)
    win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk_code, 0)

# --- Main Application Class ---
class AudioBotGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("音频按键助手")
        self.master.geometry("500x450")
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

        # State variables
        self.is_running = False
        self.is_acting = False  # Action lock
        self.worker_thread = None
        self.log_queue = queue.Queue()
        self.last_trigger_time = time.time()
        
        # UI Setup
        self.create_widgets()
        self.process_log_queue()

    def create_widgets(self):
        main_frame = ttk.Frame(self.master, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Controls Frame ---
        controls_frame = ttk.LabelFrame(main_frame, text="控制面板", padding="10")
        controls_frame.pack(fill=tk.X)
        controls_frame.grid_columnconfigure(0, weight=1)
        controls_frame.grid_columnconfigure(1, weight=1)

        # Settings
        ttk.Label(controls_frame, text="响度阈值:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.threshold_var = tk.DoubleVar(value=0.08)
        self.threshold_slider = ttk.Scale(controls_frame, from_=0.01, to=1.0, orient=tk.HORIZONTAL, variable=self.threshold_var)
        self.threshold_slider.grid(row=0, column=1, sticky=tk.EW)
        
        ttk.Label(controls_frame, text="收杆按键 (检测到声音):").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.reel_in_key_var = tk.StringVar(value='x')
        ttk.Entry(controls_frame, textvariable=self.reel_in_key_var).grid(row=1, column=1, sticky=tk.EW)

        ttk.Label(controls_frame, text="抛竿按键 (超时):").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.cast_rod_key_var = tk.StringVar(value='z')
        ttk.Entry(controls_frame, textvariable=self.cast_rod_key_var).grid(row=2, column=1, sticky=tk.EW)
        
        ttk.Label(controls_frame, text="超时时间 (秒):").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.timeout_var = tk.IntVar(value=30)
        ttk.Entry(controls_frame, textvariable=self.timeout_var).grid(row=3, column=1, sticky=tk.EW)

        # Actions and Status
        action_frame = ttk.Frame(controls_frame)
        action_frame.grid(row=4, column=0, columnspan=2, pady=10)
        
        self.start_button = ttk.Button(action_frame, text="启动", command=self.start_worker)
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.stop_button = ttk.Button(action_frame, text="停止", command=self.stop_worker, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        self.status_var = tk.StringVar(value="状态: 已停止")
        ttk.Label(action_frame, textvariable=self.status_var).pack(side=tk.LEFT, padx=10)

        # --- Log Frame ---
        log_frame = ttk.LabelFrame(main_frame, text="日志", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.configure(state='disabled')

    def log_message(self, message):
        self.log_queue.put(message)

    def process_log_queue(self):
        try:
            while True:
                message = self.log_queue.get_nowait()
                self.log_text.configure(state='normal')
                self.log_text.insert(tk.END, f"{time.strftime('%H:%M:%S')} - {message}\n")
                self.log_text.configure(state='disabled')
                self.log_text.see(tk.END)
        except queue.Empty:
            pass
        finally:
            self.master.after(100, self.process_log_queue)

    def start_worker(self):
        self.is_running = True
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.status_var.set("状态: 运行中...")
        self.log_message("启动监听...")
        
        self.worker_thread = threading.Thread(target=self._monitor_loop, daemon=True)
        self.worker_thread.start()

    def stop_worker(self):
        self.is_running = False
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.status_var.set("状态: 已停止")
        self.log_message("停止监听。")

    def get_vk_code(self, key_str):
        key_str = key_str.lower()
        return VK_CODE_MAP.get(key_str)

    def _audio_callback(self, indata, frames, time_info, status):
        if not self.is_running or self.is_acting:
            return
            
        volume_norm = np.linalg.norm(indata) / np.sqrt(len(indata))
        if volume_norm > self.threshold_var.get():
            self.is_acting = True
            # Run the action sequence in a separate thread to avoid blocking the audio callback
            action_thread = threading.Thread(target=self._perform_reel_and_cast, daemon=True)
            action_thread.start()

    def _perform_reel_and_cast(self):
        try:
            if not self.is_running:
                return

            self.log_message(f"检测到大音量，执行收杆后抛竿...")
            hwnd = find_wow_hwnd()
            if not hwnd:
                self.log_message("未找到 wow.exe 窗口，取消操作。")
                return

            # 1. Reel in
            reel_in_vk = self.get_vk_code(self.reel_in_key_var.get())
            if not reel_in_vk:
                self.log_message(f"错误：无效的收杆按键 '{self.reel_in_key_var.get()}'")
                return
            post_key(hwnd, reel_in_vk)
            
            # 2. Wait for 1-2 seconds
            delay = random.uniform(1, 2)
            self.log_message(f"收杆后等待 {delay:.1f} 秒...")
            time.sleep(delay)
            
            if not self.is_running:
                return

            # 3. Cast rod
            self.log_message("执行抛竿。")
            cast_rod_vk = self.get_vk_code(self.cast_rod_key_var.get())
            if not cast_rod_vk:
                self.log_message(f"错误：无效的抛竿按键 '{self.cast_rod_key_var.get()}'")
                return
            post_key(hwnd, cast_rod_vk)
            
            # 4. Reset timeout timer
            self.last_trigger_time = time.time()
            self.log_message("动作完成，超时计时器已重置。")

        finally:
            self.is_acting = False # Release lock

    def _monitor_loop(self):
        self.last_trigger_time = time.time()
        
        try:
            stream = sd.InputStream(
                callback=self._audio_callback,
                channels=1,
                samplerate=16000,
                blocksize=1600
            )
            stream.start()
            self.log_message(f"监听已开始，阈值: {self.threshold_var.get():.2f}")
            
            while self.is_running:
                timeout_seconds = self.timeout_var.get()
                if time.time() - self.last_trigger_time > timeout_seconds:
                    self.log_message(f"超时 {timeout_seconds} 秒，执行抛竿。")
                    hwnd = find_wow_hwnd()
                    if hwnd:
                        cast_rod_vk = self.get_vk_code(self.cast_rod_key_var.get())
                        if cast_rod_vk:
                            post_key(hwnd, cast_rod_vk)
                            self.last_trigger_time = time.time()
                        else:
                            self.log_message(f"错误：无效的抛竿按键 '{self.cast_rod_key_var.get()}'")
                    else:
                        self.log_message("未找到 wow.exe 窗口，但已重置超时计时器。")
                        self.last_trigger_time = time.time()
                time.sleep(1)
                
        except Exception as e:
            self.log_message(f"发生错误: {e}")
        finally:
            if 'stream' in locals() and stream.active:
                stream.stop()
                stream.close()
            # Ensure the GUI state is updated correctly upon thread exit
            self.master.after(0, self.stop_worker)
            
    def on_closing(self):
        if self.is_running:
            self.stop_worker()
        self.master.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = AudioBotGUI(root)
    root.mainloop() 