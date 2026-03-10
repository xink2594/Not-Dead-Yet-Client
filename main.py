import tkinter as tk
from tkinter import messagebox, ttk
import threading
import time
import json
import random
import sys
import os
import winreg
import paho.mqtt.client as mqtt
import win32api
import win32gui
import win32process
import psutil

# ================= 核心监控逻辑 =================

def get_idle_time():
    """获取键鼠空闲时间（秒）"""
    return (win32api.GetTickCount() - win32api.GetLastInputInfo()) / 1000.0

def get_active_window_info():
    """获取当前前台激活的进程名和窗口标题"""
    try:
        hwnd = win32gui.GetForegroundWindow()
        title = win32gui.GetWindowText(hwnd)
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        process = psutil.Process(pid)
        app_name = process.name()
        return app_name, title
    except Exception:
        return "Unknown", "Unknown"

# ================= 开机自启逻辑 =================

def set_autostart(enable=True):
    """设置或取消 Windows 开机自启"""
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    app_name = "NotDeadYetAgent"
    exe_path = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__)
    
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_ALL_ACCESS)
        if enable:
            winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, exe_path)
        else:
            winreg.DeleteValue(key, app_name)
        winreg.CloseKey(key)
        return True
    except Exception as e:
        print(f"设置开机自启失败: {e}")
        return False

# ================= UI 与主程序 =================

class AgentApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Not-Dead-Yet Agent")
        self.root.geometry("380x350")
        self.root.resizable(False, False)

        self.running = False
        self.mqtt_client = None
        self.monitor_thread = None

        self.setup_ui()

    def setup_ui(self):
        padding = {'padx': 10, 'pady': 5}

        # 1. Broker 地址 (此处前端填 wss，后端探针这里我们直接用 TCP 域名)
        tk.Label(self.root, text="MQTT Broker (域名):").grid(row=0, column=0, sticky="w", **padding)
        self.broker_var = tk.StringVar(value="broker.emqx.io") 
        tk.Entry(self.root, textvariable=self.broker_var, width=30).grid(row=1, column=0, columnspan=2, **padding)

        # 2. Topic 设置
        tk.Label(self.root, text="你的专属 Topic (房间/用户名):").grid(row=2, column=0, sticky="w", **padding)
        self.topic_var = tk.StringVar(value="NotDeadYet/TyTyLoo/Sleepy-Potato-9527")
        tk.Entry(self.root, textvariable=self.topic_var, width=30).grid(row=3, column=0, **padding)
        tk.Button(self.root, text="🎲", command=self.generate_topic).grid(row=3, column=1, padx=5)

        # 3. 发送频率
        tk.Label(self.root, text="发送频率 (秒):").grid(row=4, column=0, sticky="w", **padding)
        self.interval_var = tk.IntVar(value=5)
        ttk.Combobox(self.root, textvariable=self.interval_var, values=(3, 5, 10, 30), width=10).grid(row=4, column=1, sticky="w")

        # 4. 开机自启
        self.autostart_var = tk.BooleanVar(value=False)
        tk.Checkbutton(self.root, text="开启开机自启", variable=self.autostart_var, command=self.toggle_autostart).grid(row=5, column=0, columnspan=2, sticky="w", **padding)

        # 5. 控制按钮
        self.start_btn = tk.Button(self.root, text="▶ 启动探针", bg="#10B981", fg="white", font=("Arial", 12, "bold"), command=self.toggle_running)
        self.start_btn.grid(row=6, column=0, columnspan=2, pady=20, ipadx=50, ipady=5)

        # 6. 状态栏
        self.status_label = tk.Label(self.root, text="状态: 准备就绪", fg="gray")
        self.status_label.grid(row=7, column=0, columnspan=2)

    def generate_topic(self):
        adjs = ["Sleepy", "Crazy", "Lazy", "Flying", "Cyber", "Ninja"]
        nouns = ["Potato", "Cat", "Fox", "Panda", "Coder", "Doggo"]
        random_name = f"{random.choice(adjs)}-{random.choice(nouns)}-{random.randint(1000,9999)}"
        self.topic_var.set(f"NotDeadYet/TyTyLoo/{random_name}")

    def toggle_autostart(self):
        set_autostart(self.autostart_var.get())

    def toggle_running(self):
        if not self.running:
            self.start_agent()
        else:
            self.stop_agent()

    def start_agent(self):
        self.running = True
        self.start_btn.config(text="⏹ 停止探针", bg="#EF4444")
        self.status_label.config(text="状态: 正在连接...", fg="blue")
        
        self.monitor_thread = threading.Thread(target=self.agent_loop, daemon=True)
        self.monitor_thread.start()

    def stop_agent(self):
        self.running = False
        self.start_btn.config(text="▶ 启动探针", bg="#10B981")
        self.status_label.config(text="状态: 已停止", fg="gray")
        if self.mqtt_client:
            self.mqtt_client.disconnect()

    def agent_loop(self):
        broker = self.broker_var.get()
        topic = self.topic_var.get()
        interval = self.interval_var.get()
        
        user_id = topic.split('/')[-1]

        try:
            # 修复 1：适配 paho-mqtt 2.0 版本 API
            self.mqtt_client = mqtt.Client(mqtt.CallbackAPIVersion.VERSION2)
            
            # 设置遗嘱 (LWT)
            lwt_payload = json.dumps({"user": user_id, "status": "offline"})
            self.mqtt_client.will_set(topic, payload=lwt_payload, retain=True)

            # 修复 2：放弃复杂的 websockets，直接走最底层、最快的 1883 TCP 端口
            self.mqtt_client.connect(broker, 1883, 60)
            self.mqtt_client.loop_start()
            self.root.after(0, lambda: self.status_label.config(text="状态: 运行中 🟢 (TCP协议)", fg="green"))
        except Exception as e:
            self.root.after(0, lambda: self.status_label.config(text=f"连接失败: {e}", fg="red"))
            self.stop_agent()
            return

        last_state = None

        while self.running:
            idle_time = get_idle_time()
            app_name, win_title = get_active_window_info()
            
            # 空闲时间大于 5 分钟 (300秒) 算作 AFK
            if idle_time > 300: 
                current_status = "afk"
            else:
                current_status = "online"

            current_state = {
                "user": user_id,
                "status": current_status,
                "active_app": app_name if current_status == "online" else "",
                "window_title": win_title if current_status == "online" else ""
            }

            if current_state != last_state:
                try:
                    self.mqtt_client.publish(topic, json.dumps(current_state), retain=True)
                    last_state = current_state
                    print(f"[{time.strftime('%H:%M:%S')}] 数据已发送: {current_state}")
                except Exception as e:
                    print(f"发送失败: {e}")

            time.sleep(interval)

if __name__ == "__main__":
    root = tk.Tk()
    app = AgentApp(root)
    root.mainloop()