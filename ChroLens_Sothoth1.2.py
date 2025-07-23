# row0: 上方操作區（圖庫按鈕、Script按鈕、存檔按鈕、新增動作按鈕）
# row1: 腳本選單（腳本下拉選單、修改腳本名稱輸入框、修改按鈕）
# row2: 主區塊（動作表格/Treeview、日誌顯示區）
# row3: 下方執行區（重複次數、重複時間、回放速度、執行/停止按鈕、狀態顯示）
# pyinstaller --onedir --noconsole ChroLens_Sothoth1.2.py
# 抓取定點的圖片07/23


import ttkbootstrap as tb
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter import ttk, filedialog, simpledialog, messagebox
import pyautogui
import cv2
import numpy as np
import time
import threading
import os
import shutil
import json
import atexit
import keyboard
import mouse
import datetime
import copy
import ctypes
import sys
import pywinauto
import requests
import tempfile

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

IMAGE_DIR = os.path.join(BASE_DIR, "images")
SCRIPTS_DIR = os.path.join(BASE_DIR, "scripts")
HOTKEY_CONFIG_PATH = os.path.join(BASE_DIR, "hotkey_config.json")
LAST_SESSION_FILE = os.path.join(BASE_DIR, "last_session.json")
if not os.path.exists(IMAGE_DIR):
    os.makedirs(IMAGE_DIR)
if not os.path.exists(SCRIPTS_DIR):
    os.makedirs(SCRIPTS_DIR)

class Action:
    def __init__(self, pic_key, img_path, action, delay=1.1, detect_wait=0, stop_on_fail=False, loop_detect=False):
        self.pic_key = pic_key
        self.img_path = img_path
        self.action = action
        self.delay = delay
        self.detect_wait = detect_wait
        self.stop_on_fail = stop_on_fail
        self.loop_detect = loop_detect

def locate_image_on_screen(template_path, confidence=0.8):
    import imutils
    screenshot = pyautogui.screenshot()
    screenshot_rgb = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
    try:
        file_bytes = np.fromfile(template_path, dtype=np.uint8)
        template = cv2.imdecode(file_bytes, cv2.IMREAD_COLOR)
    except Exception as e:
        from tkinter import messagebox
        messagebox.showerror("錯誤", f"無法讀取圖片: {template_path}\n{e}")
        return None
    if template is None:
        from tkinter import messagebox
        messagebox.showerror("錯誤", f"無法讀取圖片: {template_path}")
        return None

    # 1. 彩色比對
    res = cv2.matchTemplate(screenshot_rgb, template, cv2.TM_CCOEFF_NORMED)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
    if max_val >= confidence:
        h, w = template.shape[:2]
        center = (max_loc[0] + w // 2, max_loc[1] + h // 2)
        return center

    # 2. 多尺度比對
    for scale in np.linspace(0.7, 1.3, 13)[::-1]:
        resized = imutils.resize(template, width=int(template.shape[1] * scale))
        if resized.shape[0] > screenshot_rgb.shape[0] or resized.shape[1] > screenshot_rgb.shape[1]:
            continue
        res = cv2.matchTemplate(screenshot_rgb, resized, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
        if max_val >= confidence:
            h, w = resized.shape[:2]
            center = (max_loc[0] + w // 2, max_loc[1] + h // 2)
            return center

    # 3. ORB 特徵點比對
    try:
        orb = cv2.ORB_create()
        kp1, des1 = orb.detectAndCompute(template, None)
        kp2, des2 = orb.detectAndCompute(screenshot_rgb, None)
        if des1 is not None and des2 is not None:
            bf = cv2.BFMatcher(cv2.NORM_HAMMING, crossCheck=True)
            matches = bf.match(des1, des2)
            matches = sorted(matches, key=lambda x: x.distance)
            if len(matches) > 10:
                pts = [kp2[m.trainIdx].pt for m in matches[:10]]
                x = int(np.mean([p[0] for p in pts]))
                y = int(np.mean([p[1] for p in pts]))
                return (x, y)
    except Exception:
        pass

    return None

def move_mouse_abs(x, y):
    ctypes.windll.user32.SetCursorPos(int(x), int(y))

def mouse_event_win(event, x=0, y=0, button='left', delta=0):
    user32 = ctypes.windll.user32
    if not button:
        button = 'left'
    if event == 'down' or event == 'up':
        flags = {'left': (0x0002, 0x0004), 'right': (0x0008, 0x0010), 'middle': (0x0020, 0x0040)}
        flag = flags.get(button, (0x0002, 0x0004))[0 if event == 'down' else 1]
        class MOUSEINPUT(ctypes.Structure):
            _fields_ = [("dx", ctypes.c_long),
                        ("dy", ctypes.c_long),
                        ("mouseData", ctypes.c_ulong),
                        ("dwFlags", ctypes.c_ulong),
                        ("time", ctypes.c_ulong),
                        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong))]
        class INPUT(ctypes.Structure):
            _fields_ = [("type", ctypes.c_ulong),
                        ("mi", MOUSEINPUT)]
        inp = INPUT()
        inp.type = 0
        inp.mi = MOUSEINPUT(0, 0, 0, flag, 0, None)
        user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(inp))
    elif event == 'wheel':
        class MOUSEINPUT(ctypes.Structure):
            _fields_ = [("dx", ctypes.c_long),
                        ("dy", ctypes.c_long),
                        ("mouseData", ctypes.c_ulong),
                        ("dwFlags", ctypes.c_ulong),
                        ("time", ctypes.c_ulong),
                        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong))]
        class INPUT(ctypes.Structure):
            _fields_ = [("type", ctypes.c_ulong),
                        ("mi", MOUSEINPUT)]
        inp = INPUT()
        inp.type = 0
        inp.mi = MOUSEINPUT(0, 0, int(delta * 120), 0x0800, 0, None)
        user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(inp))

class ChroLens_SothothApp(tb.Window):
    def __init__(self):
        super().__init__(themename="superhero")
        self.title("ChroLens_Sothoth1.2")
        self.resizable(False, False)
        self.actions = []
        self.pic_map = {}
        self.drag_data = {"item": None, "index": None}
        self._stop_flag = threading.Event()



        # ====== 上方操作區 ======
        frm_top = tb.Frame(self, padding=(10, 10, 10, 5))
        frm_top.pack(fill="x")
        self.btn_gallery = tb.Button(frm_top, text="圖庫", width=8, bootstyle=SECONDARY, command=self.open_gallery)
        self.btn_gallery.grid(row=0, column=0, padx=4)
        self.btn_script = tb.Button(frm_top, text="Script", width=8, bootstyle=INFO, command=self.open_script_merge)
        self.btn_script.grid(row=0, column=1, padx=4)
        self.btn_scheme = tb.Button(frm_top, text="方案", width=8, bootstyle=SUCCESS, command=self.open_scheme_folder)
        self.btn_scheme.grid(row=0, column=2, padx=4)
        self.btn_add = tb.Button(frm_top, text="新增動作", width=16, bootstyle=PRIMARY, command=self.add_action)
        self.btn_add.grid(row=0, column=3, padx=4)

        # ====== 新增：快捷鍵設定按鈕（加在最右邊） ======
        self.btn_hotkey = tb.Button(frm_top, text="快捷鍵", width=8, bootstyle=SECONDARY, command=self.open_hotkey_settings)
        self.btn_hotkey.grid(row=0, column=10, padx=4, sticky="e")

        # ====== row1: 方案選單區 ======
        frm_scheme = tb.Frame(self, padding=(10, 0, 10, 5))
        frm_scheme.pack(fill="x")

        tb.Label(frm_scheme, text="方案：", font=("Microsoft JhengHei", 11)).pack(side="left", padx=(0, 4))
        self.script_var = tk.StringVar()
        self.script_menu = tb.Combobox(frm_scheme, textvariable=self.script_var, width=24, state="readonly")
        self.script_menu.pack(side="left", padx=(0, 4))
        self.refresh_script_menu()
        self.script_menu.bind("<<ComboboxSelected>>", self.on_script_select)

        # 新增：儲存按鈕（在下拉選單右邊）
        btn_scheme_save = tb.Button(frm_scheme, text="儲存", width=8, bootstyle=SUCCESS, command=self.save_actions)
        btn_scheme_save.pack(side="left", padx=(0, 8))

        tb.Label(frm_scheme, text="修改名稱：", font=("Microsoft JhengHei", 11)).pack(side="left", padx=(0, 4))
        self.rename_var = tk.StringVar()
        tb.Entry(frm_scheme, textvariable=self.rename_var, width=18).pack(side="left", padx=(0, 4))
        tb.Button(frm_scheme, text="修改", width=8, bootstyle=WARNING, command=self.rename_script).pack(side="left", padx=(0, 4))

        # ====== 主區塊：動作表格與日誌 ======
        frm_main = tb.Frame(self, padding=(10, 5, 10, 5))
        frm_main.pack(fill="both", expand=True)

        frm_action = tb.Frame(frm_main)
        frm_action.pack(side="left", fill="y", expand=False)
        columns = ("#","pic_key","action","delay")
        self.tree = ttk.Treeview(frm_action, columns=columns, show="headings", height=18, selectmode="extended")
        self.tree.heading("#", text="序")
        self.tree.heading("pic_key", text="圖片名稱")
        self.tree.heading("action", text="動作")
        self.tree.heading("delay", text="延遲(秒)")
        self.tree.column("#", width=40, anchor="center")
        self.tree.column("pic_key", width=180, anchor="w")
        self.tree.column("action", width=100, anchor="center")
        self.tree.column("delay", width=80, anchor="center")
        self.tree.pack(fill="y", expand=False, pady=4)
        self.tree.bind("<Double-1>", self.on_tree_edit)
        self.tree.bind("<Button-1>", self.on_tree_click)
        self.tree.bind("<Delete>", self.on_tree_delete)
        self.tree.bind("<Button-3>", self.on_tree_right_click)
        self.tree.bind("<ButtonPress-1>", self.on_drag_start)
        self.tree.bind("<B1-Motion>", self.on_drag_motion)
        self.tree.bind("<ButtonRelease-1>", self.on_drag_drop)

        frm_log = tb.Frame(frm_main)
        frm_log.pack(side="left", fill="both", expand=True, padx=(10,0))
        self.log_text = tb.Text(frm_log, height=18, width=50, state="disabled", font=("Microsoft JhengHei", 9))
        self.log_text.pack(fill="both", expand=True, pady=(4,0))

        # ====== 下方執行區 ======
        frm_bottom = tb.Frame(self, padding=(10, 0, 10, 10))
        frm_bottom.pack(fill="x")
        tb.Label(frm_bottom, text="重複次數:", style="My.TLabel").pack(side="left", padx=(0,2))
        self.repeat_var = tk.StringVar(value="1")
        tb.Entry(frm_bottom, textvariable=self.repeat_var, width=6).pack(side="left", padx=(0,8))
        tb.Label(frm_bottom, text="重複時間:", style="My.TLabel").pack(side="left", padx=(0,2))
        self.repeat_time_var = tk.StringVar(value="00:00:00")
        repeat_time_entry = tb.Entry(frm_bottom, textvariable=self.repeat_time_var, width=10, foreground="#888888")
        repeat_time_entry.pack(side="left", padx=(0,8))
        def on_focus_in(event):
            if self.repeat_time_var.get() == "00:00:00":
                repeat_time_entry.delete(0, tk.END)
                repeat_time_entry.config(foreground="#000000")
        def on_focus_out(event):
            if not repeat_time_entry.get():
                self.repeat_time_var.set("00:00:00")
                repeat_time_entry.config(foreground="#888888")
        repeat_time_entry.bind("<FocusIn>", on_focus_in)
        repeat_time_entry.bind("<FocusOut>", on_focus_out)
        tb.Label(frm_bottom, text="回放速度(1x=100):", style="My.TLabel").pack(side="left", padx=(0,2))
        self.speed_var = tk.StringVar(value="100")
        tb.Entry(frm_bottom, textvariable=self.speed_var, width=6).pack(side="left", padx=(0,8))
        self.btn_run = tb.Button(frm_bottom, text="執行(F6)", width=8, bootstyle=DANGER, command=self.run_actions)
        self.btn_run.pack(side="right", padx=2)
        self.btn_stop = tb.Button(frm_bottom, text="停止(F7)", width=8, bootstyle=WARNING, command=self.stop_actions)
        self.btn_stop.pack(side="right", padx=2)
        self.status_label = tb.Label(frm_bottom, text="狀態：待命", style="My.TLabel")
        self.status_label.pack(side="left", padx=4)

        # 註冊全域快捷鍵（F6: 執行, F7: 停止）
        try:
            keyboard.add_hotkey('F6', lambda: self.run_actions(), suppress=False)
            keyboard.add_hotkey('F7', lambda: self.stop_actions(), suppress=False)
        except Exception as e:
            print(f"全域快捷鍵註冊失敗: {e}")

        # 在 __init__ 結尾，註冊錄製/停止錄製快捷鍵
        try:
            keyboard.add_hotkey('F10', lambda: self.start_record_script(self.script_var), suppress=False)
            keyboard.add_hotkey('F9', lambda: self.stop_record_script(), suppress=False)
        except Exception as e:
            print(f"錄製快捷鍵註冊失敗: {e}")

        self.update_tree()
        self.update_idletasks()
        self.minsize(self.winfo_reqwidth(), self.winfo_reqheight())
        self.load_last_session()
        atexit.register(self.save_last_session)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    # ====== 功能方法區 ======
    def open_gallery(self):
        folder = os.path.abspath(IMAGE_DIR)
        os.startfile(folder)
        self.log("開啟圖庫資料夾")

    # 新增：開啟方案資料夾
    def open_scheme_folder(self):
        folder = os.path.abspath(SCRIPTS_DIR)
        os.startfile(folder)
        self.log("開啟方案資料夾")

    # 方案選單只顯示方案檔案
    # 只排除 script_ 開頭的 json
    def refresh_script_menu(self):
        scripts = [os.path.splitext(f)[0] for f in os.listdir(SCRIPTS_DIR)
                   if f.endswith(".json") and f.startswith("p_")]
        self.script_menu["values"] = scripts
        if scripts:
            self.script_var.set(scripts[0])
        else:
            self.script_var.set("")

    def on_script_select(self, event=None):
        name = self.script_var.get()
        if name:
            self.load_script(name)

    # 只允許載入 p_ 開頭
    def load_script(self, script_name):
        if not script_name.startswith("p_"):
            messagebox.showerror("錯誤", "只能載入以 p_ 開頭的方案檔案")
            return
        path = os.path.join(SCRIPTS_DIR, script_name + ".json")
        if not os.path.exists(path):
            messagebox.showerror("錯誤", f"找不到方案檔案：{script_name}")
            return
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.actions.clear()
        self.tree.delete(*self.tree.get_children())
        for item in data:
            act = Action(
                item.get("pic_key", ""),
                item.get("img_path", ""),
                item.get("action", ""),
                item.get("delay", 0),
                item.get("detect_wait", 0),
                item.get("stop_on_fail", False),
                item.get("loop_detect", False)
            )
            self.actions.append(act)
            self.tree.insert("", "end", values=(len(self.actions), act.pic_key, "Script" if act.action.startswith("[SCRIPT]") else act.action, f"{act.delay:.1f}"))
        self.script_var.set(script_name)
        self.log(f"載入方案：{script_name}")

    def update_tree(self):
        self.tree.delete(*self.tree.get_children())
        for idx, act in enumerate(self.actions, 1):
            # 顯示時去除 [SCRIPT]
            action_display = act.action.replace("[SCRIPT]", "") if act.action.startswith("[SCRIPT]") else act.action
            self.tree.insert("", "end", values=(idx, act.pic_key, action_display, f"{act.delay:.1f}"))

    def log(self, msg):
        self.log_text.config(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def on_tree_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell":
            rowid = self.tree.identify_row(event.y)
            col = self.tree.identify_column(event.x)
            if not rowid:
                self.add_action_direct()
            else:
                self.on_tree_edit(event)

    def on_tree_edit(self, event):
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not item:
            return
        idx = int(self.tree.item(item, "values")[0]) - 1
        act = self.actions[idx]
        col_idx = int(col.replace("#", "")) - 1
        if col_idx == 1:
            self.edit_pic(act, idx)
        elif col_idx == 2:
            self.edit_action(act, idx)
        elif col_idx == 3:
            self.edit_delay_tree(act, idx)

    def edit_pic(self, act, idx):
        img_path = filedialog.askopenfilename(
            title="選擇圖片",
            filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp")]
        )
        if not img_path:
            return
        pic_num = idx + 1
        base_name = os.path.basename(img_path)
        name8 = os.path.splitext(base_name)[0][:8]
        ext = os.path.splitext(base_name)[1]
        pic_key = f"pic{pic_num}_{name8}"
        save_name = f"{pic_key}{ext}"
        save_path = os.path.join(IMAGE_DIR, save_name)
        # 無論是否存在都覆蓋
        try:
            shutil.copy(img_path, save_path)
        except Exception as e:
            messagebox.showerror("錯誤", f"圖片複製失敗: {e}")
            return
        act.pic_key = pic_key
        act.img_path = save_path
        self.update_tree()
        self.log(f"編輯圖片：{pic_key}")

    def edit_action(self, act, idx):
        win = tk.Toplevel(self)
        win.title("動作編輯視窗")
        win.geometry("500x400")
        win.resizable(False, False)

        # 1. 按鍵（只捕捉一個動作）
        frm_key = tb.Frame(win)
        frm_key.pack(fill="x", padx=10, pady=(20, 0), anchor="w")
        tb.Label(frm_key, text="按鍵", width=6, anchor="w").pack(side="left", anchor="w")
        if act.action.startswith("[SCRIPT]"):
            key_init = ""
            script_init = act.action.replace("[SCRIPT]", "")
        else:
            key_init = act.action
            script_init = ""
        key_var = tk.StringVar(value=key_init)
        key_entry = tb.Entry(frm_key, textvariable=key_var, width=20, font=("Microsoft JhengHei", 12), state="readonly")
        key_entry.pack(side="left", fill="x", expand=True, anchor="w")

        import keyboard as kb
        import threading

        def on_focus_in(event):
            def listen_key():
                try:
                    event_kb = kb.read_event(suppress=True)
                    if event_kb.event_type == "down":
                        name = event_kb.name
                        try:
                            if key_entry.winfo_exists():
                                key_entry.config(state="normal")
                                key_var.set(name)
                                key_entry.config(state="readonly")
                        except Exception:
                            pass
                except Exception:
                    try:
                        if key_entry.winfo_exists():
                            key_entry.config(state="normal")
                            key_var.set("")
                            key_entry.config(state="readonly")
                    except Exception:
                        pass
            key_entry.config(state="normal")
            key_var.set("")
            key_entry.config(state="readonly")
            threading.Thread(target=listen_key, daemon=True).start()

        key_entry.bind("<FocusIn>", on_focus_in)
        key_entry.bind("<KeyPress>", lambda e: "break")

        def on_mouse(event):
            key_entry.focus_set()
            key_entry.config(state="normal")
            if event.num == 1:
                key_var.set("左鍵點擊")
            elif event.num == 3:
                key_var.set("右鍵點擊")
            elif event.num == 2:
                key_var.set("中鍵點擊")
            key_entry.config(state="readonly")
            return "break"

        def on_mouse_wheel(event):
            key_entry.focus_set()
            key_entry.config(state="normal")
            if event.delta > 0:
                key_var.set("滾輪上")
            else:
                key_var.set("滾輪下")
            key_entry.config(state="readonly")
            return "break"

        key_entry.bind("<Button-1>", on_mouse)
        key_entry.bind("<Button-2>", on_mouse)
        key_entry.bind("<Button-3>", on_mouse)
        key_entry.bind("<MouseWheel>", on_mouse_wheel)

        # 2. Script
        frm_script = tb.Frame(win)
        frm_script.pack(fill="x", padx=10, pady=(30, 0), anchor="w")
        tb.Label(frm_script, text="Script", width=6, anchor="w").pack(side="left", anchor="w")
        script_var = tk.StringVar(value=script_init)
        script_files = [os.path.splitext(f)[0] for f in os.listdir(SCRIPTS_DIR)
                        if f.endswith(".json") and f.startswith("s_")]
        script_combo = tb.Combobox(frm_script, textvariable=script_var, values=script_files, width=20, state="readonly")
        script_combo.pack(side="left", fill="x", expand=True, anchor="w")

        # === 新增：腳本名稱修改框與按鈕 ===
        frm_rename = tb.Frame(win)
        frm_rename.pack(fill="x", padx=10, pady=(10, 0), anchor="w")
        rename_var = tk.StringVar()
        entry_rename = tb.Entry(frm_rename, textvariable=rename_var, width=20)
        entry_rename.pack(side="left", padx=4, anchor="w")
        def do_rename():
            old_name = script_var.get()
            new_name = rename_var.get().strip()
            if not old_name or not new_name:
                messagebox.showinfo("提示", "請選擇腳本並輸入新名稱。")
                return
            # 強制補 s_ 前綴
            if not new_name.startswith("s_"):
                new_name = "s_" + new_name
            if not new_name.endswith('.json'):
                new_name += '.json'
            old_path = os.path.join(SCRIPTS_DIR, old_name + ".json") if not old_name.endswith('.json') else os.path.join(SCRIPTS_DIR, old_name)
            new_path = os.path.join(SCRIPTS_DIR, new_name)
            if os.path.exists(new_path):
                messagebox.showerror("錯誤", "檔案已存在，請換個名稱。")
                return
            try:
                os.rename(old_path, new_path)
                self.log(f"腳本已更名為：{new_name}")
                # 重新整理下拉選單
                script_files = [os.path.splitext(f)[0] for f in os.listdir(SCRIPTS_DIR) if f.endsWith(".json") and f.startswith("s_")]
                script_combo["values"] = script_files
                script_var.set(os.path.splitext(new_name)[0])
            except Exception as e:
                messagebox.showerror("錯誤", f"更名失敗: {e}")
            rename_var.set("")
        btn_rename = tb.Button(frm_rename, text="修改腳本名稱", command=do_rename, bootstyle=WARNING, width=12)
        btn_rename.pack(side="left", padx=4, anchor="w")

        # 3. 錄製快捷鍵
        frm_record = tb.Frame(win)
        frm_record.pack(fill="x", padx=10, pady=(30, 0), anchor="w")
        import keyboard

        config_path = "hotkey_config.json"
        if os.path.exists(config_path):
            with open(config_path, "r", encoding="utf-8") as f:
                hotkey_config = json.load(f)
        else:
            hotkey_config = {"record": "F10", "stop_record": "F9"}

        record_hotkey_str = hotkey_config.get("record", "F10")
        stop_hotkey_str = hotkey_config.get("stop_record", "F9")

        def trigger_record():
            self.start_record_script(script_var)
        def trigger_stop_record():
            self.stop_record_script()

        btn_record = tb.Button(frm_record, text=f"錄製({record_hotkey_str})", width=12, bootstyle=SUCCESS, command=trigger_record)
        btn_record.pack(side="left", padx=4, anchor="w")
        btn_stop_record = tb.Button(frm_record, text=f"停止錄製({stop_hotkey_str})", width=12, bootstyle=WARNING, command=trigger_stop_record)
        btn_stop_record.pack(side="left", padx=4, anchor="w")

        # 只要視窗存在就註冊快捷鍵，關閉時自動移除
        record_hotkey = None
        stop_hotkey = None
        try:
            record_hotkey = keyboard.add_hotkey(record_hotkey_str, trigger_record, suppress=False)
        except Exception as e:
            messagebox.showerror("快捷鍵錯誤", f"錄製快捷鍵設定錯誤: {record_hotkey_str}\n{e}")
        try:
            stop_hotkey = keyboard.add_hotkey(stop_hotkey_str, trigger_stop_record, suppress=False)
        except Exception as e:
            messagebox.showerror("快捷鍵錯誤", f"停止錄製快捷鍵設定錯誤: {stop_hotkey_str}\n{e}")

        def on_close():
            if record_hotkey is not None:
                try:
                    keyboard.remove_hotkey(record_hotkey)
                except Exception:
                    pass
            if stop_hotkey is not None:
                try:
                    keyboard.remove_hotkey(stop_hotkey)
                except Exception:
                    pass
            win.destroy()

        win.protocol("WM_DELETE_WINDOW", on_close)
        win.grab_set()

        def on_ok():
            key_action = key_var.get().strip()
            script_name = script_var.get().strip()
            if key_action:
                act.action = key_action
            elif script_name:
                # 強制腳本名稱為 s_xxx 格式
                if not script_name.startswith("s_"):
                    script_name = "s_" + script_name
                act.action = f"[SCRIPT]{script_name}"
            else:
                act.action = ""
            win.destroy()
            # 修正日誌顯示
            if act.action.startswith("[SCRIPT]"):
                log_name = act.action.replace("[SCRIPT]", "")
            else:
                log_name = act.action
            self.update_tree()
            self.log(f"編輯動作：{log_name} 延遲{act.delay}秒")

        tb.Button(win, text="確定", bootstyle=SUCCESS, width=12, command=on_ok).pack(pady=30, anchor="center")

    def edit_delay_tree(self, act, idx):
        win = tk.Toplevel(self)
        win.title("延遲設定")
        win.geometry("340x340")  # 高度+50
        win.resizable(False, False)

        # 動作區塊
        frm_action = tb.Frame(win)
        frm_action.pack(fill="x", pady=(18, 0))
        tb.Label(frm_action, text="動作：延遲秒數", font=("Microsoft JhengHei", 12)).pack(side="left", padx=(8, 4))
        delay_var = tk.StringVar(value=str(int(act.delay) if act.delay == int(act.delay) else act.delay))
        entry_delay = tk.Entry(frm_action, textvariable=delay_var, width=10, font=("Microsoft JhengHei", 12))
        entry_delay.pack(side="left", fill="x")

        # 分隔線
        tb.Separator(win, orient="horizontal").pack(fill="x", pady=10)

        # 圖片偵測區塊
        frm_img = tb.Frame(win)
        frm_img.pack(fill="x")
        tb.Label(frm_img, text="圖片：偵測秒數", font=("Microsoft JhengHei", 12)).pack(side="left", padx=(8, 4))
        detect_wait_var = tk.StringVar(value=str(int(getattr(act, "detect_wait", 0)) if getattr(act, "detect_wait", 0) == int(getattr(act, "detect_wait", 0)) else getattr(act, "detect_wait", 0)))
        entry_detect = tk.Entry(frm_img, textvariable=detect_wait_var, width=10, font=("Microsoft JhengHei", 12))
        entry_detect.pack(side="left", fill="x")

        # 單選框控制變數
        detect_mode_var = tk.StringVar(value="auto_stop")

        # 偵測失敗自動停止（預設勾選）
        rb0 = tk.Radiobutton(
            win,
            text="偵測失敗自動停止",
            variable=detect_mode_var,
            value="auto_stop",
            font=("Microsoft JhengHei", 12)
        )
        rb0.pack(anchor="w", padx=30, pady=(10, 0))

        # 偵測失敗自動繼續
        rb1 = tk.Radiobutton(
            win,
            text="偵測失敗自動繼續",
            variable=detect_mode_var,
            value="continue",
            font=("Microsoft JhengHei", 12)
        )
        rb1.pack(anchor="w", padx=30, pady=(2, 0))

        # 循環偵測
        rb2 = tk.Radiobutton(
            win,
            text="循環偵測",
            variable=detect_mode_var,
            value="loop",
            font=("Microsoft JhengHei", 12)
        )
        rb2.pack(anchor="w", padx=30, pady=(2, 0))

        def on_ok():
            # 延遲秒數
            try:
                delay = float(delay_var.get())
                act.delay = delay
            except Exception:
                act.delay = 0
            # 偵測等待秒數
            try:
                detect_wait = float(detect_wait_var.get())
                act.detect_wait = detect_wait
            except Exception:
                act.detect_wait = 0
            # 根據單選框設定動作屬性
            mode = detect_mode_var.get()
            act.stop_on_fail = (mode == "continue")
            act.loop_detect = (mode == "loop")
            # auto_stop 不需特別寫入，因為預設就是這個行為
            self.update_tree()
            self.log(f"編輯延遲：{act.pic_key} - {act.action} - {act.delay}秒")
            win.destroy()

        btn = tb.Button(win, text="確定", bootstyle=SUCCESS, width=10, command=on_ok)
        btn.pack(pady=18)
        win.grab_set()


    def add_action_direct(self):
        pass

    def add_action(self):
        # 直接新增一個空白動作
        act = Action("", "", "", 0)
        self.actions.append(act)
        self.update_tree()
        self.log("已新增空白動作，請於列表中編輯內容")

    def run_actions(self):
        if not self.actions:
            messagebox.showinfo("提示", "請先新增動作")
            return
        try:
            repeat = int(self.repeat_var.get())
        except Exception:
            repeat = 1
        try:
            t = self.repeat_time_var.get()
            if ":" in t:
                parts = [int(x) for x in t.split(":")]
                while len(parts) < 3:
                    parts.insert(0, 0)
                repeat_time = parts[0]*3600 + parts[1]*60 + parts[2]
            else:
                repeat_time = int(t)
        except Exception:
            repeat_time = 0
        try:
            speed = int(self.speed_var.get())
            if speed < 1 or speed > 1000:
                speed = 100
        except Exception:
            speed = 100

        # 只要重複時間有設定（非0），就以時間為主，忽略次數
        if repeat_time > 0:
            repeat = 0  # 0 代表無限循環，直到時間到
        self._stop_flag.clear()
        self.btn_run.config(state=tk.DISABLED)
        self.status_label.config(text="狀態：執行中", foreground="#DB0E59")
        def on_finish():
            self.btn_run.config(state=tk.NORMAL)
            self.status_label.config(text="狀態：待命", foreground="#15D3BD")
            self.log("動作執行完畢")
        threading.Thread(
            target=self.perform_actions,
            args=(self.actions, on_finish, repeat, repeat_time, speed),
            daemon=True
        ).start()
        if repeat_time > 0:
            self.log(f"開始執行動作（重複時間{repeat_time}秒，速度{speed}）")
        elif repeat == 0:
            self.log(f"開始執行動作（無限循環，速度{speed}）")
        else:
            self.log(f"開始執行動作（重複{repeat}次，速度{speed}）")

    def perform_actions(self, actions, on_finish, repeat=1, repeat_time=0, speed=100):
        import pyautogui, time, keyboard, mouse, json, os
        speed_ratio = speed / 100.0
        start_time = time.time()
        count = 0
        while (repeat == 0 or count < repeat):
            for act in actions:
                if self._stop_flag.is_set():
                    on_finish()
                    return
                # 如果是錄製的 macro（json list）
                if act.action and act.action.strip().startswith("[") and act.action.strip().endswith("]"):
                    try:
                        macro = json.loads(act.action)
                    except Exception as e:
                        self.log(f"載入錄製動作失敗: {e}")
                        continue
                    if not macro:
                        continue
                    base_time = macro[0]["time"]
                    play_start = time.time()
                    for idx, e in enumerate(macro):
                        if self._stop_flag.is_set():
                            on_finish()
                            return
                        event_offset = (e['time'] - base_time) / speed_ratio
                        target_time = play_start + event_offset
                        while True:
                            now = time.time()
                            if self._stop_flag.is_set():
                                on_finish()
                                return
                            if now >= target_time:
                                break
                            time.sleep(min(0.01, target_time - now))
                        # 執行事件
                        if e["type"] == "keyboard":
                            if e["event"] == "down":
                                keyboard.press(e["name"])
                            elif e["event"] == "up":
                                keyboard.release(e["name"])
                        elif e["type"] == "mouse":
                            if e.get("event") == "move":
                                move_mouse_abs(e["x"], e["y"])
                            elif e.get("event") == "down":
                                mouse_event_win('down', button=e.get('button', 'left'))
                            elif e.get("event") == "up":
                                mouse_event_win('up', button=e.get('button', 'left'))
                            elif e.get("event") == "wheel":
                                mouse_event_win('wheel', delta=e.get('delta', 0))
                else:
                    # 單一動作（支援組合鍵）
                    if act.action.startswith("[SCRIPT]"):
                        # 執行子腳本
                        script_name = act.action.replace("[SCRIPT]", "").strip()
                        script_path = os.path.join(SCRIPTS_DIR, script_name + ".json")
                        if os.path.exists(script_path):
                            with open(script_path, "r", encoding="utf-8") as f:
                                content = f.read()
                                try:
                                    macro = json.loads(content)
                                except json.JSONDecodeError:
                                    # 只取第一個合法陣列
                                    import re
                                    match = re.search(r'\[[\s\S]*?\]', content)
                                    if match:
                                        try:
                                            macro = json.loads(match.group(0))
                                        except Exception as e:
                                            self.log(f"腳本格式錯誤：{script_name} ({e})")
                                            continue
                                    else:
                                        self.log(f"腳本格式錯誤：{script_name}")
                                        continue
                            base_time = macro[0]["time"] if macro else 0
                            t0 = time.time()
                            for e in macro:
                                if self._stop_flag.is_set():
                                    on_finish()
                                    return
                                wait = (e["time"] - base_time) / speed_ratio
                                while time.time() - t0 < wait:
                                    time.sleep(0.005)
                                # 執行事件
                                if e["type"] == "keyboard":
                                    if e["event"] == "down":
                                        keyboard.press(e["name"])
                                    elif e["event"] == "up":
                                        keyboard.release(e["name"])
                                elif e["type"] == "mouse":
                                    if e.get("event") == "move":
                                        move_mouse_abs(e["x"], e["y"])
                                    elif e.get("event") == "down":
                                        mouse_event_win('down', button=e.get('button', 'left'))
                                    elif e.get("event") == "up":
                                        mouse_event_win('up', button=e.get('button', 'left'))
                                    elif e.get("event") == "wheel":
                                        mouse_event_win('wheel', delta=e.get('delta', 0))
                        else:
                            self.log(f"找不到子腳本：{script_name}")
                    elif "+" in act.action:
                        keys = act.action.split("+")
                        for k in keys:
                            keyboard.press(k)
                        for k in reversed(keys):
                            keyboard.release(k)
                    elif act.action in ("左鍵點擊", "右鍵點擊"):
                        # --- 新增：如果有圖片，先找圖片並移動 ---
                        if act.img_path and os.path.exists(act.img_path):
                            found = False
                            detect_wait = getattr(act, "detect_wait", 0)
                            stop_on_fail = getattr(act, "stop_on_fail", False)
                            loop_detect = getattr(act, "loop_detect", False)
                            start_time = time.time()
                            while True:
                                # 新增：優先用 UI Automation
                                pos = self.locate_element_or_image(
                                    act.img_path,
                                    confidence=0.8,
                                    search_text=act.pic_key  # 假設 pic_key 是 UI 元件名稱
                                )
                                if pos:
                                    found = True
                                    break
                                if detect_wait > 0 and (time.time() - start_time) >= detect_wait:
                                    break
                                if not loop_detect:
                                    time.sleep(0.5 / speed_ratio)
                                    break
                                time.sleep(0.5 / speed_ratio)
                            if not found:
                                if stop_on_fail:
                                    messagebox.showwarning("警告", f"找不到圖片: {os.path.basename(act.img_path)}，已停止所有動作")
                                    on_finish()
                                    return
                                else:
                                    self.log(f"找不到圖片: {os.path.basename(act.img_path)}，略過此動作")
                                    continue
                            pyautogui.moveTo(pos)
                        # --- 再執行點擊 ---
                        if act.action == "左鍵點擊":
                            pyautogui.click()
                        elif act.action == "右鍵點擊":
                            pyautogui.click(button='right')
                    elif act.action:
                        keyboard.press_and_release(act.action)
                    time.sleep(act.delay / speed_ratio)
            count += 1
            # 只要有設定重複時間，時間到就結束
            if repeat_time > 0:
                elapsed = time.time() - start_time
                if elapsed >= repeat_time:
                    break
        on_finish()

    def stop_record_script(self):
        """停止錄製腳本"""
        self._recording_script = False
        self.log("已停止錄製")

    def open_hotkey_settings(self):
        win = tk.Toplevel(self)
        win.title("快捷鍵設定")
        win.geometry("340x300")
        win.resizable(False, False)

        # 你可以根據實際支援的功能調整 labels
        labels = {
            "run": "執行(F6)",
            "stop": "停止(F7)",
            "record": "錄製(F10)",
            "stop_record": "停止錄製F9"
        }
        default_hotkeys = {
            "run": "F6",
            "stop": "F7",
            "record": "F10",
            "stop_record": "F9"
        }
        # 讀取現有設定
        config_path = "hotkey_config.json"
        if os.path.exists(config_path):
            with open(config_path, "r", encoding="utf-8") as f:
                config = json.load(f)
        else:
            config = default_hotkeys.copy()
        vars = {}
        row = 0

        def on_entry_key(event, key, var):
            keys = []
            if event.state & 0x0001: keys.append("shift")
            if event.state & 0x0004: keys.append("ctrl")
            if event.state & 0x0008: keys.append("alt")
            key_name = event.keysym.lower()
            if key_name not in ("shift_l", "shift_r", "control_l", "control_r", "alt_l", "alt_r"):
                keys.append(key_name)
            var.set("+".join(keys))
            return "break"

        def on_entry_focus_in(event, var):
            var.set("請輸入按鍵")

        def on_entry_focus_out(event, key, var):
            if var.get() == "請輸入按鍵" or not var.get():
                var.set(config.get(key, default_hotkeys[key]))

        for key, label in labels.items():
            tb.Label(win, text=label, font=("Microsoft JhengHei", 11)).grid(row=row, column=0, padx=10, pady=8, sticky="w")
            var = tk.StringVar(value=config.get(key, default_hotkeys[key]))
            entry = tb.Entry(win, textvariable=var, width=16, font=("Consolas", 11), state="normal")
            entry.grid(row=row, column=1, padx=10)
            vars[key] = var
            entry.bind("<KeyRelease>", lambda e, k=key, v=var: on_entry_key(e, k, v))
            entry.bind("<FocusIn>", lambda e, v=var: on_entry_focus_in(e, v))
            entry.bind("<FocusOut>", lambda e, k=key, v=var: on_entry_focus_out(e, k, v))
            row += 1

        def save_and_apply():
            # 先移除舊的錄製快捷鍵
            try:
                keyboard.remove_hotkey(hotkey_config.get("record", "F10"))
            except Exception:
                pass
            try:
                keyboard.remove_hotkey(hotkey_config.get("stop_record", "F9"))
            except Exception:
                pass
            for key in default_hotkeys:
                val = vars[key].get()
                if val and val != "請輸入按鍵":
                    config[key] = val.upper()
            with open(HOTKEY_CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("完成", "快捷鍵設定已儲存")
            win.destroy()

        tb.Button(win, text="儲存", command=save_and_apply, width=10, bootstyle=SUCCESS).grid(row=row, column=0, columnspan=2, pady=16)
        win.grab_set()

    def open_script_merge(self):
        win = tk.Toplevel(self)
        win.title("腳本合併工具")
        win.geometry("900x550")
        win.resizable(False, False)

        left_frame = tb.Frame(win)
        left_frame.pack(side="left", fill="y", padx=10, pady=10)
        tb.Label(left_frame, text="所有腳本", style="My.TLabel").pack()
        script_files = [f for f in os.listdir(SCRIPTS_DIR) if f.endswith(".json") and f.startswith("s_")]
        script_names = [os.path.splitext(f)[0] for f in script_files]
        listbox_all = tk.Listbox(left_frame, selectmode="extended", width=28)
        for name in script_names:
            listbox_all.insert(tk.END, name)
        listbox_all.pack(fill="y", expand=True)

        btn_frame = tb.Frame(win)
        btn_frame.pack(side="left", fill="y", padx=5)
        btn_add = tb.Button(btn_frame, text="加入 →")
        btn_add.pack(pady=10)
        btn_remove = tb.Button(btn_frame, text="← 移除")
        btn_remove.pack(pady=10)
        btn_delete = tb.Button(btn_frame, text="刪除腳本", bootstyle="danger", width=10)
        btn_delete.pack(pady=20)
        btn_merge = tb.Button(btn_frame, text="合併", bootstyle="success", width=10)
        btn_merge.pack(pady=10)

        right_frame = tb.Frame(win)
        right_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        tb.Label(right_frame, text="合併清單", style="My.TLabel").pack()
        listbox_merge = tk.Listbox(right_frame, selectmode="extended", width=28)
        listbox_merge.pack(fill="both", expand=True)

        # 新增：合併命名區塊
        frm_merge_bottom = tb.Frame(right_frame)
        frm_merge_bottom.pack(fill="x", pady=(10, 0))
        tb.Label(frm_merge_bottom, text="合併名稱").pack(side="left", padx=(0, 4))
        entry_name = tb.Entry(frm_merge_bottom, width=18)
        entry_name.pack(side="left", padx=(0, 8))
        btn_merge_save = tb.Button(frm_merge_bottom, text="合併並儲存", width=12)
        btn_merge_save.pack(side="left")

        merge_items = []

        def on_delete():
            selected = listbox_all.curselection()
            if not selected:
                messagebox.showwarning("提示", "請先選擇要刪除的腳本")
                return
            names = [listbox_all.get(i) for i in selected]
            confirm = messagebox.askyesno("確認", f"確定要刪除這些腳本嗎？\n{', '.join(names)}")
            if not confirm:
                return
            for name in names:
                path = os.path.join(SCRIPTS_DIR, name + ".json")
                try:
                    os.remove(path)
                except Exception as e:
                    messagebox.showerror("錯誤", f"刪除失敗：{name}\n{e}")
            refresh_all_list()
            self.refresh_script_menu()
            self.log(f"已刪除腳本：{', '.join(names)}")

        def refresh_merge_list():
            listbox_merge.delete(0, tk.END)
            for name in merge_items:
                listbox_merge.insert(tk.END, name)

        def refresh_all_list():
            listbox_all.delete(0, tk.END)
            script_files = [f for f in os.listdir(SCRIPTS_DIR) if f.endswith(".json") and f.startswith("s_")]
            script_names = [os.path.splitext(f)[0] for f in script_files]
            for name in script_names:
                listbox_all.insert(tk.END, name)

        def on_add():
            selected = listbox_all.curselection()
            for i in selected:
                name = listbox_all.get(i)
                if name not in merge_items:
                    merge_items.append(name)
            refresh_merge_list()

        def on_remove():
            selected = list(listbox_merge.curselection())
            for i in reversed(selected):
                merge_items.pop(i)
            refresh_merge_list()

        def do_merge():
            if not merge_items:
                messagebox.showwarning("提示", "請先將要合併的腳本加入右側清單")
                return
            new_name = entry_name.get().strip()
            if not new_name:
                messagebox.showerror("錯誤", "請輸入新腳本名稱")
                return
            if not new_name.startswith("s_"):
                new_name = "s_" + new_name
            merged = []
            for name in merge_items:
                path = os.path.join(SCRIPTS_DIR, name + ".json")
                with open(path, "r", encoding="utf-8") as f:
                    merged += json.load(f)
            save_path = os.path.join(SCRIPTS_DIR, new_name + ".json")
            if os.path.exists(save_path):
                messagebox.showerror("錯誤", "檔案已存在，請換個新名稱")
                return
            with open(save_path, "w", encoding="utf-8") as f:
                json.dump(merged, f, ensure_ascii=False, indent=2)
            self.refresh_script_menu()
            self.script_var.set(new_name)
            self.log(f"合併並儲存腳本：{new_name}")
            win.destroy()

        btn_add.config(command=on_add)
        btn_remove.config(command=on_remove)
        btn_merge_save.config(command=do_merge)
        btn_delete.config(command=on_delete)
        btn_merge.config(command=do_merge)
        win.grab_set()

    def save_actions(self):
        if not self.actions:
            messagebox.showinfo("提示", "目前沒有動作可存檔")
            return
        current_name = self.script_var.get()
        name = simpledialog.askstring("存檔", "請輸入方案名稱（不需加p_）：", initialvalue=current_name.replace("p_", ""))
        if not name:
            return
        if not name.startswith("p_"):
            name = "p_" + name
        save_path = os.path.join(SCRIPTS_DIR, name + ".json")
        data = [
            {
                "pic_key": act.pic_key,
                "img_path": act.img_path,
                "action": act.action,
                "delay": act.delay,
                "detect_wait": getattr(act, "detect_wait", 0),
                "stop_on_fail": getattr(act, "stop_on_fail", False),
                "loop_detect": getattr(act, "loop_detect", False)
            }
            for act in self.actions
        ]
        with open(save_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        self.refresh_script_menu()
        self.script_var.set(name)
        self.log(f"儲存方案：{name}")

    def on_tree_delete(self, event):
        # 取得所有選取的項目
        selected = self.tree.selection()
        if not selected:
            return
        # 取得要刪除的索引（反向排序，避免刪除時索引錯亂）
        idxs = sorted([int(self.tree.item(item, "values")[0]) - 1 for item in selected], reverse=True)
        for idx in idxs:
            if 0 <= idx < len(self.actions):
                del self.actions[idx]
        self.update_tree()
        self.log(f"已刪除 {len(idxs)} 個動作")

    def on_tree_right_click(self, event):
        item = self.tree.identify_row(event.y)
        if not item:
            return
        idx = int(self.tree.item(item, "values")[0]) - 1
        if 0 <= idx < len(self.actions):
            del self.actions[idx]
            self.update_tree()
            self.log(f"已刪除第{idx+1}個動作（右鍵）")

    def on_close(self):
        self.save_last_session()
        self.destroy()

    def save_last_session(self):
        try:
            data = {
                "actions": [
                    {
                        "pic_key": act.pic_key,
                        "img_path": act.img_path,
                        "action": act.action,
                        "delay": act.delay
                    }
                    for act in self.actions
                ],
                "pic_map": self.pic_map,
                "script_var": self.script_var.get() if hasattr(self, "script_var") else "",
            }
            with open(LAST_SESSION_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"自動儲存失敗: {e}")

    def load_last_session(self):
        if os.path.exists(LAST_SESSION_FILE):
            try:
                with open(LAST_SESSION_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self.actions.clear()
                for item in data.get("actions", []):
                    act = Action(item["pic_key"], item["img_path"], item["action"], item["delay"])
                    self.actions.append(act)
                self.pic_map = data.get("pic_map", {})
                if hasattr(self, "script_var") and data.get("script_var"):
                    self.script_var.set(data["script_var"])
                self.update_tree()
            except Exception as e:
                print(f"自動載入失敗: {e}")

    def start_record_script(self, script_var):
        import keyboard
        from pynput.mouse import Controller, Listener
        import pynput.mouse
        import datetime

        self._recording_script = True
        self._recorded_events = []
        self._mouse_events = []
        self._record_start_time = time.time()
        mouse_ctrl = Controller()
        last_pos = mouse_ctrl.position

        def now_abs():
            return time.time()

        def on_click(x, y, button, pressed):
            if self._recording_script:
                self._mouse_events.append({
                    'type': 'mouse',
                    'event': 'down' if pressed else 'up',
                    'button': str(button).replace('Button.', ''),
                    'x': x,
                    'y': y,
                    'time': now_abs()
                })
        def on_scroll(x, y, dx, dy):
            if self._recording_script:
                self._mouse_events.append({
                    'type': 'mouse',
                    'event': 'wheel',
                    'delta': dy,
                    'x': x,
                    'y': y,
                    'time': now_abs()
                })
        mouse_listener = pynput.mouse.Listener(
            on_click=on_click,
            on_scroll=on_scroll
        )
        mouse_listener.start()
        self._mouse_events.append({
            'type': 'mouse',
            'event': 'move',
            'x': last_pos[0],
            'y': last_pos[1],
            'time': now_abs()
        })

        def do_record():
            nonlocal last_pos
            try:
                keyboard.start_recording()
                while self._recording_script:
                    pos = mouse_ctrl.position
                    now = now_abs()
                    if pos != last_pos:
                        self._mouse_events.append({
                            'type': 'mouse',
                            'event': 'move',
                            'x': pos[0],
                            'y': pos[1],
                            'time': now
                        })
                        last_pos = pos
                    time.sleep(0.01)
                mouse_listener.stop()
                try:
                    k_events = keyboard.stop_recording()
                except KeyError:
                    k_events = []
                filtered_k_events = k_events
                events = [
                    {'type': 'keyboard', 'event': e.event_type, 'name': e.name, 'time': e.time}
                    for e in filtered_k_events
                ] + self._mouse_events
                all_events = sorted(events, key=lambda e: e['time'])
                # 存檔
                name = script_var.get().strip()
                if not name or not name.startswith("s_"):
                    # 若未輸入名稱則自動產生
                    ts = datetime.datetime.now().strftime("%Y_%m%d_%H%M_%S")
                    name = f"s_{ts}"
                filename = name + ".json"
                save_path = os.path.join(SCRIPTS_DIR, filename)
                with open(save_path, "w", encoding="utf-8") as f:
                    json.dump(all_events, f, ensure_ascii=False, indent=2)
                # 更新下拉選單
                script_files = [os.path.splitext(f)[0] for f in os.listdir(SCRIPTS_DIR) if f.endswith(".json") and f.startswith("s_")]
                script_var.set(os.path.splitext(filename)[0])
                self.refresh_script_menu()
                self.log(f"錄製Script已儲存：{filename}")
            except Exception as e:
                self.log(f"錄製腳本時發生錯誤: {e}")
        threading.Thread(target=do_record, daemon=True).start()

    def on_drag_start(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.drag_data = {"item": item, "index": self.tree.index(item)}
        else:
            self.drag_data = {"item": None, "index": None}

    def on_drag_motion(self, event):
        pass  # 可選：可視化拖曳效果

    def on_drag_drop(self, event):
        if not self.drag_data.get("item"):
            return
        target_item = self.tree.identify_row(event.y)
        if not target_item or target_item == self.drag_data["item"]:
            return
        from_idx = self.drag_data["index"]
        to_idx = self.tree.index(target_item)
        # 交換 self.actions 順序
        act = self.actions.pop(from_idx)
        self.actions.insert(to_idx, act)
        self.update_tree()
        self.log(f"已將動作從第{from_idx+1}移到第{to_idx+1}")

    def rename_script(self):
        old_name = self.script_var.get()
        new_name = self.rename_var.get().strip()
        if not old_name or not new_name:
            messagebox.showinfo("提示", "請選擇腳本並輸入新名稱。")
            return
        if not new_name.startswith("p_"):
            new_name = "p_" + new_name
        old_path = os.path.join(SCRIPTS_DIR, old_name + ".json") if not old_name.endswith('.json') else os.path.join(SCRIPTS_DIR, old_name)
        new_path = os.path.join(SCRIPTS_DIR, new_name + ".json")
        if os.path.exists(new_path):
            messagebox.showerror("錯誤", "檔案已存在，請換個名稱。")
            return
        try:
            os.rename(old_path, new_path)
            self.log(f"腳本已更名為：{new_name}")
            self.refresh_script_menu()
            self.script_var.set(new_name)
        except Exception as e:
            messagebox.showerror("錯誤", f"更名失敗: {e}")
        self.rename_var.set("")  # 更名後清空輸入框

    def stop_actions(self):
        """停止所有動作執行"""
        self._stop_flag.set()
        self.btn_run.config(state=tk.NORMAL)
        self.status_label.config(text="狀態：已停止", foreground="#888888")
        self.log("已手動停止動作")

    def locate_element_or_image(self, template_path, confidence=0.8, search_text=None, app_title=None):
        # 1. UI Automation
        if search_text and not search_text.startswith("pic"):
            try:
                app = pywinauto.Application(backend="uia").connect(title=app_title) if app_title else pywinauto.Application(backend="uia").connect(path=sys.executable)
                dlg = app.top_window()
                ctrl = dlg.child_window(title=search_text)
                rect = ctrl.rectangle()
                center = ((rect.left + rect.right)//2, (rect.top + rect.bottom)//2)
                return center
            except Exception:
                pass
        # 2. SIFT 特徵點
        try:
            pos = locate_image_with_sift(template_path, confidence)
            if pos:
                return pos
        except Exception:
            pass
        # 3. 原本的圖片比對
        return locate_image_on_screen(template_path, confidence)

def locate_image_with_sift(template_path, confidence=0.8):
    import cv2
    screenshot = pyautogui.screenshot()
    screenshot_rgb = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
    file_bytes = np.fromfile(template_path, dtype=np.uint8)
    template = cv2.imdecode(file_bytes, cv2.IMREAD_COLOR)
    sift = cv2.SIFT_create()
    kp1, des1 = sift.detectAndCompute(template, None)
    kp2, des2 = sift.detectAndCompute(screenshot_rgb, None)
    bf = cv2.BFMatcher()
    matches = bf.knnMatch(des1, des2, k=2)
    good = []
    for m, n in matches:
        if m.distance < 0.75 * n.distance:
            good.append(m)
    if len(good) > 8:
        pts = [kp2[m.trainIdx].pt for m in good]
        x = int(np.mean([p[0] for p in pts]))
        y = int(np.mean([p[1] for p in pts]))
        return (x, y)
    return None

APP_VERSION = "1.2.0"  # ⚠️ 每次更新時記得手動改這裡
LATEST_JSON_URL = "https://raw.githubusercontent.com/你的帳號/你的repo/main/latest_version.json"  # TODO: 改成你的網址

def check_and_update():
    try:
        r = requests.get(LATEST_JSON_URL, timeout=5)
        r.raise_for_status()
        latest = r.json()
        latest_version = latest.get("version", "")
        download_url = latest.get("url", "")

        if not latest_version or not download_url:
            print("更新檢查失敗：資料缺失")
            return

        if latest_version > APP_VERSION:
            answer = messagebox.askyesno("更新提示", f"檢測到新版本 {latest_version}，是否下載更新？")
            if answer:
                tmp_path = os.path.join(tempfile.gettempdir(), "ChroLens_Sothoth_new.exe")
                with requests.get(download_url, stream=True) as r2:
                    r2.raise_for_status()
                    with open(tmp_path, "wb") as f:
                        for chunk in r2.iter_content(chunk_size=8192):
                            f.write(chunk)

                messagebox.showinfo("更新完成", "已下載更新，即將重啟程式完成更新。")
                os.startfile(tmp_path)
                sys.exit(0)

    except Exception as e:
        print(f"檢查更新錯誤：{e}")

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    pass


if __name__ == "__main__":
    check_and_update()
    app = ChroLens_SothothApp()
    app.mainloop()
