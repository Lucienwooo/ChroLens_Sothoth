# row0: 上方操作區（圖庫按鈕、Script按鈕、存檔按鈕、新增動作按鈕）
# row1: 腳本選單（腳本下拉選單、修改腳本名稱輸入框、修改按鈕）
# row2: 主區塊（動作表格/Treeview、日誌顯示區）
# row3: 下方執行區（重複次數、重複時間、回放速度、執行/停止按鈕、狀態顯示）


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

IMAGE_DIR = "images"
SCRIPTS_DIR = "scripts"
LAST_SESSION_FILE = "last_session.json"
if not os.path.exists(IMAGE_DIR):
    os.makedirs(IMAGE_DIR)
if not os.path.exists(SCRIPTS_DIR):
    os.makedirs(SCRIPTS_DIR)

class Action:
    def __init__(self, pic_key, img_path, action, delay=1.0):
        self.pic_key = pic_key
        self.img_path = img_path
        self.action = action
        self.delay = delay

def locate_image_on_screen(template_path, confidence=0.8):
    screenshot = pyautogui.screenshot()
    screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2GRAY)
    try:
        file_bytes = np.fromfile(template_path, dtype=np.uint8)
        template = cv2.imdecode(file_bytes, cv2.IMREAD_GRAYSCALE)
    except Exception as e:
        from tkinter import messagebox
        messagebox.showerror("錯誤", f"無法讀取圖片: {template_path}\n{e}")
        return None
    if template is None:
        from tkinter import messagebox
        messagebox.showerror("錯誤", f"無法讀取圖片: {template_path}")
        return None
    res = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
    if max_val >= confidence:
        h, w = template.shape[:2]
        center = (max_loc[0] + w // 2, max_loc[1] + h // 2)
        return center
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
        self.title("ChroLens_Sothoth1.0")
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
        self.btn_save = tb.Button(frm_top, text="存檔", width=8, bootstyle=SUCCESS, command=self.save_actions)
        self.btn_save.grid(row=0, column=2, padx=4)
        self.btn_add = tb.Button(frm_top, text="新增動作", width=16, bootstyle=PRIMARY, command=self.add_action)
        self.btn_add.grid(row=0, column=3, padx=4)

        # ====== 腳本選單獨立 row ======
        frm_script = tb.Frame(self, padding=(10, 0, 10, 5))
        frm_script.pack(fill="x")
        tb.Label(frm_script, text="腳本選單:", style="My.TLabel").pack(side="left")
        self.script_var = tk.StringVar()
        self.script_menu = tb.Combobox(frm_script, textvariable=self.script_var, width=24, state="readonly", style="My.TCombobox")
        self.script_menu.pack(side="left", padx=4)
        self.refresh_script_menu()
        self.script_menu.bind("<<ComboboxSelected>>", self.on_script_select)

        # 新增：修改腳本名稱輸入框與按鈕
        self.rename_var = tk.StringVar()
        self.rename_entry = tb.Entry(frm_script, textvariable=self.rename_var, width=20)
        self.rename_entry.pack(side="left", padx=4)
        self.btn_rename = tb.Button(frm_script, text="修改腳本名稱", command=self.rename_script, bootstyle=WARNING, width=12)
        self.btn_rename.pack(side="left", padx=4)

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

    def refresh_script_menu(self):
        scripts = [os.path.splitext(f)[0] for f in os.listdir(SCRIPTS_DIR) if f.endswith(".json")]
        self.script_menu["values"] = scripts
        if scripts:
            self.script_var.set(scripts[0])
        else:
            self.script_var.set("")

    def on_script_select(self, event=None):
        name = self.script_var.get()
        if name:
            self.load_script(name)

    def load_script(self, script_name):
        path = os.path.join(SCRIPTS_DIR, script_name + ".json")
        if not os.path.exists(path):
            messagebox.showerror("錯誤", f"找不到腳本檔案：{script_name}")
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
                item.get("delay", 0)
            )
            self.actions.append(act)
            self.tree.insert("", "end", values=(len(self.actions), act.pic_key, act.action, f"{act.delay:.1f}"))
        self.script_var.set(script_name)
        self.log(f"載入腳本：{script_name}")

    def update_tree(self):
        self.tree.delete(*self.tree.get_children())
        for idx, act in enumerate(self.actions, 1):
            self.tree.insert("", "end", values=(idx, act.pic_key, act.action, f"{act.delay:.1f}"))

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
        img_path = filedialog.askopenfilename(title="選擇圖片", filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp")])
        if not img_path:
            return
        pic_num = idx + 1
        base_name = os.path.basename(img_path)
        name8 = os.path.splitext(base_name)[0][:8]
        ext = os.path.splitext(base_name)[1]
        pic_key = f"pic{pic_num}_{name8}"
        save_name = f"{pic_key}{ext}"
        save_path = os.path.join(IMAGE_DIR, save_name)
        if not os.path.exists(save_path):
            shutil.copy(img_path, save_path)
        act.pic_key = pic_key
        act.img_path = save_path
        self.update_tree()
        self.log(f"編輯圖片：{pic_key}")

    def edit_action(self, act, idx):
        win = tk.Toplevel(self)
        win.title("編輯動作/按鍵/Script")
        win.geometry("500x300")
        win.resizable(False, False)

        # 1. 按鍵（只捕捉一個動作）
        frm_key = tb.Frame(win)
        frm_key.pack(fill="x", padx=10, pady=(20, 0))
        tb.Label(frm_key, text="按鍵", width=6, anchor="w").pack(side="left")
        key_var = tk.StringVar(value=act.action)
        key_entry = tb.Entry(frm_key, textvariable=key_var, width=20, font=("Microsoft JhengHei", 12), state="readonly")
        key_entry.pack(side="left", fill="x", expand=True)
        delay_key_var = tk.DoubleVar(value=act.delay)
        tb.Entry(frm_key, textvariable=delay_key_var, width=8).pack(side="left", padx=(8, 0))

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
        key_entry.bind("<KeyPress>", lambda e: "break")  # 禁止手動輸入

        # 滑鼠事件（只捕捉一個動作）
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
        frm_script.pack(fill="x", padx=10, pady=(30, 0))
        tb.Label(frm_script, text="Script", width=6, anchor="w").pack(side="left")
        script_var = tk.StringVar()
        script_files = [os.path.splitext(f)[0] for f in os.listdir(SCRIPTS_DIR) if f.endswith(".json")]
        script_combo = tb.Combobox(frm_script, textvariable=script_var, values=script_files, width=20, state="readonly")
        script_combo.pack(side="left", fill="x", expand=True)
        delay_script_var = tk.DoubleVar(value=act.delay)
        tb.Entry(frm_script, textvariable=delay_script_var, width=8).pack(side="left", padx=(8, 0))

        # 3. 錄製快捷鍵
        frm_record = tb.Frame(win)
        frm_record.pack(fill="x", padx=10, pady=(30, 0))
        import keyboard

        def trigger_record():
            self.start_record_script(script_var)
        def trigger_stop_record():
            self.stop_record_script()

        # 註冊快捷鍵（只在本視窗存活時）
        record_hotkey = keyboard.add_hotkey('f10', trigger_record, suppress=False)
        stop_hotkey = keyboard.add_hotkey('f9', trigger_stop_record, suppress=False)

        tb.Button(frm_record, text="錄製(F10)", width=14, bootstyle=PRIMARY, command=trigger_record).pack(side="left", padx=(0, 8))
        tb.Button(frm_record, text="停止錄製(F9)", width=14, bootstyle=WARNING, command=trigger_stop_record).pack(side="left")

        def on_ok():
            # 只會擇一
            key_action = key_var.get().strip()
            script_name = script_var.get().strip()
            if key_action:
                act.action = key_action
                act.delay = delay_key_var.get()
            elif script_name:
                act.action = f"[SCRIPT]{script_name}"
                act.delay = delay_script_var.get()
            else:
                act.action = ""
                act.delay = 0
            win.destroy()
            self.update_tree()
            self.log(f"編輯動作：{act.action} 延遲{act.delay}秒")

        tb.Button(win, text="確定", bootstyle=SUCCESS, width=12, command=on_ok).pack(pady=30)

        def on_close():
            try:
                keyboard.remove_hotkey(record_hotkey)
            except Exception:
                pass
            try:
                keyboard.remove_hotkey(stop_hotkey)
            except Exception:
                pass
            win.destroy()

        win.protocol("WM_DELETE_WINDOW", on_close)
        win.grab_set()

    def edit_delay_tree(self, act, idx):
        new_delay = simpledialog.askfloat("編輯延遲", "請輸入新的延遲秒數", minvalue=0, initialvalue=act.delay)
        if new_delay is not None:
            act.delay = new_delay
            self.update_tree()
            self.log(f"編輯延遲：{act.pic_key} - {act.action} - {act.delay}秒")

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
                                macro = json.load(f)
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
                            for _ in range(10):
                                pos = locate_image_on_screen(act.img_path)
                                if pos:
                                    found = True
                                    break
                                time.sleep(0.5 / speed_ratio)
                            if not found:
                                messagebox.showwarning("警告", f"找不到圖片: {os.path.basename(act.img_path)}")
                                on_finish()
                                return
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
        win.geometry("340x580")
        win.resizable(False, False)

        # 功能與預設快捷鍵
        labels = {
            "run": "執行",
            "stop": "停止",
            "add": "新增動作",
            "script": "Script",
            "gallery": "圖庫",
            "record": "錄製",
            "record_stop": "錄製停止"
        }
        # 讀取現有設定
        config_path = "config.json"
        if os.path.exists(config_path):
            with open(config_path, "r", encoding="utf-8") as f:
                config = json.load(f)
        else:
            config = default_hotkeys.copy()
        vars = {}
        row = 0

        def on_entry_key(event, key, var):
            # 只記錄實際按下的組合鍵或單鍵
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
            # 綁定事件
            entry.bind("<KeyRelease>", lambda e, k=key, v=var: on_entry_key(e, k, v))
            entry.bind("<FocusIn>", lambda e, v=var: on_entry_focus_in(e, v))
            entry.bind("<FocusOut>", lambda e, k=key, v=var: on_entry_focus_out(e, k, v))
            row += 1

        def save_and_apply():
            for key in default_hotkeys:
                val = vars[key].get()
                if val and val != "請輸入按鍵":
                    config[key] = val.lower()
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            # 更新到 self.hotkey_config
            self.hotkey_config = config
            self.register_hotkeys()
            # 更新主畫面按鈕顯示
            self.update_hotkey_labels()
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
        script_files = [f for f in os.listdir(SCRIPTS_DIR) if f.endswith(".json")]
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
            script_files = [f for f in os.listdir(SCRIPTS_DIR) if f.endswith(".json")]
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
        # 取得目前腳本名稱，預設填入
        current_name = self.script_var.get()
        name = simpledialog.askstring("存檔", "請輸入檔案名稱：", initialvalue=current_name)
        if not name:
            return
        save_path = os.path.join(SCRIPTS_DIR, name + ".json")
        # 直接覆蓋，不再提示重複
        data = [
            {
                "pic_key": act.pic_key,
                "img_path": act.img_path,
                "action": act.action,
                "delay": act.delay
            }
            for act in self.actions
        ]
        with open(save_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        self.refresh_script_menu()
        self.script_var.set(name)
        self.log(f"儲存腳本：{name}")

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
                # ...existing code...
                filtered_k_events = k_events
                events = [
                    {'type': 'keyboard', 'event': e.event_type, 'name': e.name, 'time': e.time}
                    for e in filtered_k_events
                ] + self._mouse_events
                all_events = sorted(events, key=lambda e: e['time'])
                # 存檔
                ts = datetime.datetime.now().strftime("%Y_%m%d_%H%M_%S")
                filename = f"script_{ts}.json"
                save_path = os.path.join(SCRIPTS_DIR, filename)
                with open(save_path, "w", encoding="utf-8") as f:
                    json.dump(all_events, f, ensure_ascii=False, indent=2)
                # 更新下拉選單
                script_files = [os.path.splitext(f)[0] for f in os.listdir(SCRIPTS_DIR) if f.endswith(".json")]
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
            self.refresh_script_menu()
            self.script_var.set(os.path.splitext(new_name)[0])
        except Exception as e:
            messagebox.showerror("錯誤", f"更名失敗: {e}")
        self.rename_var.set("")  # 更名後清空輸入框

    def stop_actions(self):
        """停止所有動作執行"""
        self._stop_flag.set()
        self.btn_run.config(state=tk.NORMAL)
        self.status_label.config(text="狀態：已停止", foreground="#888888")
        self.log("已手動停止動作")


if __name__ == "__main__":
    app = ChroLens_SothothApp()
    app.mainloop()
