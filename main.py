import tkinter as tk
import customtkinter as ctk
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

# ================= 现代化 UI 全局设置 =================
ctk.set_appearance_mode("System")  # 跟随系统主题 (深色/浅色)
ctk.set_default_color_theme("blue") # 默认强调色

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

# ================= 现代化 UI 与主程序 =================

class AgentApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Not-Dead-Yet Agent")
        self.root.geometry("450x420")
        self.root.resizable(False, False)

        self.running = False
        self.mqtt_client = None
        self.monitor_thread = None

        self.setup_ui()

    def setup_ui(self):
        # 顶部标题栏
        self.title_label = ctk.CTkLabel(self.root, text="👀 Not-Dead-Yet 监控探针", font=ctk.CTkFont(size=20, weight="bold"))
        self.title_label.grid(row=0, column=0, columnspan=2, pady=(20, 20))

        # 1. Broker 地址
        ctk.CTkLabel(self.root, text="MQTT Broker (域名):", font=ctk.CTkFont(size=13)).grid(row=1, column=0, sticky="w", padx=30, pady=(5, 0))
        self.broker_var = tk.StringVar(value="broker.emqx.io") 
        self.broker_entry = ctk.CTkEntry(self.root, textvariable=self.broker_var, width=350, height=35)
        self.broker_entry.grid(row=2, column=0, columnspan=2, padx=30, pady=(0, 15))

        # 2. Topic 设置
        ctk.CTkLabel(self.root, text="你的专属 Topic (房间/用户名):", font=ctk.CTkFont(size=13)).grid(row=3, column=0, sticky="w", padx=30, pady=(5, 0))
        self.topic_var = tk.StringVar(value="NotDeadYet/TyTyLoo/CloydLarkin")
        
        # 将输入框和随机按钮放在一个框架里，使其对齐更好看
        self.topic_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        self.topic_frame.grid(row=4, column=0, columnspan=2, padx=30, pady=(0, 15), sticky="ew")
        self.topic_entry = ctk.CTkEntry(self.topic_frame, textvariable=self.topic_var, width=280, height=35)
        self.topic_entry.pack(side="left", padx=(0, 10))
        self.random_btn = ctk.CTkButton(self.topic_frame, text="🎲 随机", width=60, height=35, command=self.generate_topic)
        self.random_btn.pack(side="left")

        # 3. 发送频率 & 开机自启 (放在同一行)
        self.settings_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        self.settings_frame.grid(row=5, column=0, columnspan=2, padx=30, pady=(5, 15), sticky="ew")
        
        ctk.CTkLabel(self.settings_frame, text="发送频率:", font=ctk.CTkFont(size=13)).pack(side="left", padx=(0, 10))
        self.interval_var = ctk.StringVar(value="5")
        self.interval_combo = ctk.CTkComboBox(self.settings_frame, variable=self.interval_var, values=["3", "5", "10", "30"], width=80)
        self.interval_combo.pack(side="left", padx=(0, 30))

        self.autostart_var = ctk.BooleanVar(value=False)
        self.autostart_check = ctk.CTkCheckBox(self.settings_frame, text="开启开机自启", variable=self.autostart_var, command=self.toggle_autostart)
        self.autostart_check.pack(side="left")

        # 4. 控制大按钮
        self.start_btn = ctk.CTkButton(self.root, text="▶ 启动探针", height=45, font=ctk.CTkFont(size=15, weight="bold"),
                                       fg_color="#10B981", hover_color="#059669", command=self.toggle_running)
        self.start_btn.grid(row=6, column=0, columnspan=2, padx=30, pady=(15, 10), sticky="ew")

        # 5. 状态栏
        self.status_label = ctk.CTkLabel(self.root, text="状态: 准备就绪", text_color="gray", font=ctk.CTkFont(size=12))
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
        self.start_btn.configure(text="⏹ 停止探针", fg_color="#EF4444", hover_color="#B91C1C")
        self.status_label.configure(text="状态: 正在连接...", text_color="#3B82F6")
        
        # 禁用输入框，防止运行时修改
        self.broker_entry.configure(state="disabled")
        self.topic_entry.configure(state="disabled")
        self.interval_combo.configure(state="disabled")

        self.monitor_thread = threading.Thread(target=self.agent_loop, daemon=True)
        self.monitor_thread.start()

    def stop_agent(self):
        self.running = False
        self.start_btn.configure(text="▶ 启动探针", fg_color="#10B981", hover_color="#059669")
        self.status_label.configure(text="状态: 已停止", text_color="gray")
        
        # 恢复输入框
        self.broker_entry.configure(state="normal")
        self.topic_entry.configure(state="normal")
        self.interval_combo.configure(state="normal")

        if self.mqtt_client:
            self.mqtt_client.disconnect()

    def agent_loop(self):
        broker = self.broker_var.get()
        topic = self.topic_var.get()
        interval = int(self.interval_var.get())
        
        user_id = topic.split('/')[-1]

        try:
            self.mqtt_client = mqtt.Client(mqtt.CallbackAPIVersion.VERSION2)
            
            # 设置遗嘱 (LWT)
            lwt_payload = json.dumps({"user": user_id, "status": "offline"})
            self.mqtt_client.will_set(topic, payload=lwt_payload, retain=True)

            self.mqtt_client.connect(broker, 1883, 60)
            self.mqtt_client.loop_start()
            self.root.after(0, lambda: self.status_label.configure(text="状态: 运行中 🟢 (TCP协议)", text_color="#10B981"))
        except Exception as e:
            self.root.after(0, lambda: self.status_label.configure(text=f"连接失败: {e}", text_color="#EF4444"))
            self.root.after(0, self.stop_agent)
            return

        last_state = None

        while self.running:
            idle_time = get_idle_time()
            app_name, win_title = get_active_window_info()
            
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
    # 使用 CTk 而不是 tk.Tk()
    root = ctk.CTk()
    app = AgentApp(root)
    root.mainloop()