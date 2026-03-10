"""
Auto Clicker Pro v2 - Modern Windows Desktop Auto Clicker + Web Automation

Tabs:
  1. Simple Clicker        - Interval-based repeated clicking
  2. Coordinate Clicker    - Click at specific screen coordinates
  3. Record & Replay       - Record mouse + keyboard, replay them
  4. Image Click           - Click based on image recognition
  5. Web Automation        - Selenium browser scripting
  6. Macro Editor          - Visual drag-and-drop macro builder
  7. Scheduler             - Run macros at specific times/dates

Requirements:
  pip install -r requirements.txt
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pyautogui
from pynput import mouse as pynput_mouse, keyboard as pynput_kb
import threading
import time
import json
import os
import copy
import random
import datetime
import keyboard as kb_hotkey

try:
    from screeninfo import get_monitors
    HAS_SCREENINFO = True
except ImportError:
    HAS_SCREENINFO = False

try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.common.action_chains import ActionChains
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.chrome.service import Service as ChromeService
    try:
        from webdriver_manager.chrome import ChromeDriverManager
        HAS_WDM = True
    except ImportError:
        HAS_WDM = False
    HAS_SELENIUM = True
except ImportError:
    HAS_SELENIUM = False
    HAS_WDM = False

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.03

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

HOTKEY_START_STOP = "F6"
HOTKEY_RECORD = "F7"
HOTKEY_PICK = "F8"
HOTKEY_REPLAY = "F9"
HOTKEY_STOP = "F10"

MACRO_ACTIONS = [
    "Mouse Click",
    "Mouse Move",
    "Key Press",
    "Type Text",
    "Wait",
    "Image Click",
    "Scroll",
]


# ================================================================
#  Dark Treeview Style
# ================================================================
def setup_dark_treeview_style():
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Dark.Treeview",
                    background="#1e1e2e", foreground="#cdd6f4",
                    fieldbackground="#1e1e2e", borderwidth=0,
                    font=("Consolas", 10))
    style.configure("Dark.Treeview.Heading",
                    background="#313244", foreground="#cdd6f4",
                    font=("Segoe UI", 10, "bold"))
    style.map("Dark.Treeview", background=[("selected", "#45475a")])


# ================================================================
#  Macro Step Edit Dialog
# ================================================================
class MacroStepDialog(ctk.CTkToplevel):
    def __init__(self, parent, step=None):
        super().__init__(parent)
        self.title("Edit Macro Step" if step else "Add Macro Step")
        self.geometry("440x360")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.result = None
        self._fields = {}

        top = ctk.CTkFrame(self)
        top.pack(fill="x", padx=12, pady=(12, 4))
        ctk.CTkLabel(top, text="Action:").pack(side="left")
        self._action_var = ctk.StringVar(
            value=step["action"] if step else MACRO_ACTIONS[0])
        ctk.CTkOptionMenu(
            top, variable=self._action_var, values=MACRO_ACTIONS,
            command=self._on_action_change
        ).pack(side="left", padx=8)

        self._param_frame = ctk.CTkFrame(self)
        self._param_frame.pack(fill="both", expand=True, padx=12, pady=4)

        btn_frame = ctk.CTkFrame(self)
        btn_frame.pack(fill="x", padx=12, pady=(4, 12))
        ctk.CTkButton(btn_frame, text="OK", width=90,
                      command=self._ok).pack(side="right", padx=4)
        ctk.CTkButton(btn_frame, text="Cancel", width=90,
                      fg_color="gray30", command=self.destroy
                      ).pack(side="right", padx=4)

        self._on_action_change(self._action_var.get())
        if step:
            self._populate(step)

    def _clear_params(self):
        for w in self._param_frame.winfo_children():
            w.destroy()
        self._fields.clear()

    def _add_field(self, label, default="", row=0):
        ctk.CTkLabel(self._param_frame, text=label).grid(
            row=row, column=0, sticky="e", padx=4, pady=4)
        var = ctk.StringVar(value=str(default))
        ctk.CTkEntry(self._param_frame, textvariable=var, width=240).grid(
            row=row, column=1, padx=4, pady=4)
        self._fields[label] = var

    def _on_action_change(self, action):
        self._clear_params()
        if action == "Mouse Click":
            self._add_field("X:", "0", 0)
            self._add_field("Y:", "0", 1)
            self._add_field("Button:", "left", 2)
            self._add_field("Clicks:", "1", 3)
        elif action == "Mouse Move":
            self._add_field("X:", "0", 0)
            self._add_field("Y:", "0", 1)
            self._add_field("Duration (s):", "0.2", 2)
        elif action == "Key Press":
            self._add_field("Key:", "enter", 0)
        elif action == "Type Text":
            self._add_field("Text:", "", 0)
            self._add_field("Interval (s):", "0.02", 1)
        elif action == "Wait":
            self._add_field("Seconds:", "1.0", 0)
        elif action == "Image Click":
            self._add_field("Image Path:", "", 0)
            self._add_field("Confidence:", "0.8", 1)
            self._add_field("Timeout (s):", "10", 2)
        elif action == "Scroll":
            self._add_field("Amount:", "300", 0)
            self._add_field("X (opt):", "", 1)
            self._add_field("Y (opt):", "", 2)

    def _field_key(self, label):
        return (label.rstrip(":").lower()
                .replace(" ", "_").replace("(opt)", "")
                .replace("(s)", "s").strip())

    def _populate(self, step):
        params = step.get("params", {})
        for label, var in self._fields.items():
            key = self._field_key(label)
            if key in params:
                var.set(str(params[key]))

    def _ok(self):
        action = self._action_var.get()
        params = {}
        for label, var in self._fields.items():
            params[self._field_key(label)] = var.get()
        self.result = {"action": action, "params": params}
        self.destroy()


# ================================================================
#  Main Application
# ================================================================
class AutoClickerPro(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Auto Clicker Pro v2")
        self.geometry("960x720")
        self.minsize(800, 600)

        self.running = False
        self.recording = False
        self.recorded_events = []
        self._record_start = 0
        self._mouse_listener = None
        self._kb_listener = None
        self._last_move_t = 0
        self.web_driver = None
        self.macro_steps = []
        self.scheduled_tasks = []
        self._scheduler_active = True

        setup_dark_treeview_style()
        self._build_ui()
        self._register_hotkeys()
        self._start_scheduler_thread()

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._log("Ready. Failsafe: move mouse to top-left corner to abort.")

    # ============================================================
    #  UI Layout
    # ============================================================
    def _build_ui(self):
        top = ctk.CTkFrame(self, height=32, fg_color="gray14")
        top.pack(fill="x")
        ctk.CTkLabel(
            top, text=f"  {HOTKEY_START_STOP} Start/Stop  |  "
                      f"{HOTKEY_RECORD} Record  |  "
                      f"{HOTKEY_REPLAY} Replay  |  "
                      f"{HOTKEY_STOP} Stop  |  "
                      f"{HOTKEY_PICK} Pick Coord",
            font=("Consolas", 11), text_color="gray60"
        ).pack(side="left", padx=8, pady=4)
        self._monitor_label = ctk.CTkLabel(
            top, text="", font=("Consolas", 11), text_color="gray60")
        self._monitor_label.pack(side="right", padx=8)
        self._refresh_monitor_info()

        self.tabview = ctk.CTkTabview(self, anchor="nw")
        self.tabview.pack(fill="both", expand=True, padx=8, pady=(4, 0))

        self._build_simple_tab()
        self._build_coord_tab()
        self._build_record_tab()
        self._build_image_tab()
        self._build_web_tab()
        self._build_macro_tab()
        self._build_scheduler_tab()

        log_frame = ctk.CTkFrame(self, height=80)
        log_frame.pack(fill="x", padx=8, pady=(4, 8))
        self.log_box = ctk.CTkTextbox(
            log_frame, height=65, font=("Consolas", 10),
            state="disabled", fg_color="gray12", text_color="gray70")
        self.log_box.pack(fill="both", expand=True, padx=4, pady=4)

    def _refresh_monitor_info(self):
        if HAS_SCREENINFO:
            mons = get_monitors()
            parts = [f"Monitors: {len(mons)}"]
            for i, m in enumerate(mons):
                parts.append(f"#{i+1}: {m.width}x{m.height}+{m.x}+{m.y}")
            self._monitor_label.configure(text="  |  ".join(parts))
        else:
            w, h = pyautogui.size()
            self._monitor_label.configure(text=f"Screen: {w}x{h}")

    # ============================================================
    #  Tab 1: Simple Clicker
    # ============================================================
    def _build_simple_tab(self):
        tab = self.tabview.add("Simple Clicker")

        sec = ctk.CTkFrame(tab)
        sec.pack(fill="x", padx=12, pady=8)
        ctk.CTkLabel(sec, text="Click Interval",
                     font=("Segoe UI", 14, "bold")).grid(
            row=0, column=0, columnspan=8, sticky="w", pady=(0, 8))

        self._simple_vars = {}
        for i, (lbl, default) in enumerate([
            ("Hours:", "0"), ("Min:", "0"), ("Sec:", "1"), ("Ms:", "0")
        ]):
            ctk.CTkLabel(sec, text=lbl).grid(
                row=1, column=i * 2, sticky="e", padx=(8, 2))
            var = ctk.StringVar(value=default)
            ctk.CTkEntry(sec, textvariable=var, width=60).grid(
                row=1, column=i * 2 + 1, padx=(2, 8))
            self._simple_vars[lbl] = var

        opt = ctk.CTkFrame(tab)
        opt.pack(fill="x", padx=12, pady=4)

        ctk.CTkLabel(opt, text="Mouse Button:").grid(
            row=0, column=0, sticky="e", padx=4, pady=4)
        self.simple_button = ctk.CTkOptionMenu(
            opt, values=["left", "right", "middle"], width=100)
        self.simple_button.set("left")
        self.simple_button.grid(row=0, column=1, padx=4, pady=4)

        ctk.CTkLabel(opt, text="Click Type:").grid(
            row=0, column=2, sticky="e", padx=4)
        self.simple_click_type = ctk.CTkOptionMenu(
            opt, values=["single", "double", "triple"], width=100)
        self.simple_click_type.set("single")
        self.simple_click_type.grid(row=0, column=3, padx=4)

        ctk.CTkLabel(opt, text="Repeat:").grid(
            row=1, column=0, sticky="e", padx=4, pady=4)
        self.simple_repeat = ctk.CTkOptionMenu(
            opt, values=["Infinite", "Count"], width=100)
        self.simple_repeat.set("Infinite")
        self.simple_repeat.grid(row=1, column=1, padx=4, pady=4)

        ctk.CTkLabel(opt, text="Count:").grid(
            row=1, column=2, sticky="e", padx=4)
        self.simple_count_var = ctk.StringVar(value="10")
        ctk.CTkEntry(opt, textvariable=self.simple_count_var, width=80).grid(
            row=1, column=3, padx=4)

        self.simple_random_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(opt, text="Random jitter",
                        variable=self.simple_random_var).grid(
            row=2, column=0, columnspan=2, sticky="w", padx=4, pady=4)
        ctk.CTkLabel(opt, text="Max jitter (ms):").grid(
            row=2, column=2, sticky="e", padx=4)
        self.simple_random_max = ctk.StringVar(value="500")
        ctk.CTkEntry(opt, textvariable=self.simple_random_max, width=80).grid(
            row=2, column=3, padx=4)

        btn = ctk.CTkFrame(tab)
        btn.pack(pady=16)
        ctk.CTkButton(
            btn, text=f"Start ({HOTKEY_START_STOP})",
            fg_color="#22c55e", hover_color="#16a34a",
            width=140, command=self._start_simple
        ).pack(side="left", padx=8)
        ctk.CTkButton(
            btn, text=f"Stop ({HOTKEY_START_STOP})",
            fg_color="#ef4444", hover_color="#dc2626",
            width=140, command=self._stop
        ).pack(side="left", padx=8)

    def _get_simple_interval(self):
        h = int(self._simple_vars["Hours:"].get() or 0)
        m = int(self._simple_vars["Min:"].get() or 0)
        s = int(self._simple_vars["Sec:"].get() or 0)
        ms = int(self._simple_vars["Ms:"].get() or 0)
        return h * 3600 + m * 60 + s + ms / 1000.0

    def _start_simple(self):
        if self.running:
            return
        interval = self._get_simple_interval()
        button = self.simple_button.get()
        click_type = self.simple_click_type.get()
        repeat = self.simple_repeat.get()
        count = int(self.simple_count_var.get() or 10) if repeat == "Count" else None
        use_random = self.simple_random_var.get()
        jitter_max = int(self.simple_random_max.get() or 500)

        def worker():
            i = 0
            self._log(f"Simple clicker: {interval}s, {button} {click_type}")
            while self.running:
                if count is not None and i >= count:
                    break
                clicks = {"single": 1, "double": 2, "triple": 3}.get(click_type, 1)
                pyautogui.click(button=button, clicks=clicks)
                i += 1
                wait = interval
                if use_random:
                    wait += random.randint(0, jitter_max) / 1000.0
                time.sleep(wait)
            self.running = False
            self._log(f"Simple clicker done. Clicked {i} times.")

        self._run_thread(worker)

    # ============================================================
    #  Tab 2: Coordinate Clicker
    # ============================================================
    def _build_coord_tab(self):
        tab = self.tabview.add("Coordinates")

        ctk.CTkLabel(tab, text="Click Sequence",
                     font=("Segoe UI", 14, "bold")).pack(
            anchor="w", padx=12, pady=(8, 0))
        ctk.CTkLabel(tab, text="Format: x, y, delay_sec   (one per line)",
                     text_color="gray60").pack(anchor="w", padx=12)

        self.coord_text = ctk.CTkTextbox(
            tab, height=180, font=("Consolas", 11))
        self.coord_text.pack(fill="both", expand=True, padx=12, pady=4)
        self.coord_text.insert(
            "0.0", "# Example:\n# 500, 300, 1.0\n# 800, 400, 0.5\n")

        pick = ctk.CTkFrame(tab)
        pick.pack(fill="x", padx=12, pady=4)
        ctk.CTkButton(pick, text=f"Pick Coordinate ({HOTKEY_PICK})",
                      width=180, command=self._pick_coordinate).pack(side="left")
        info = "  Click anywhere to capture"
        if HAS_SCREENINFO:
            info += " (multi-monitor)"
        ctk.CTkLabel(pick, text=info, text_color="gray60").pack(side="left")

        opt = ctk.CTkFrame(tab)
        opt.pack(fill="x", padx=12, pady=4)
        ctk.CTkLabel(opt, text="Loops:").pack(side="left", padx=4)
        self.coord_loops_var = ctk.StringVar(value="1")
        ctk.CTkEntry(opt, textvariable=self.coord_loops_var, width=60).pack(
            side="left", padx=4)
        ctk.CTkLabel(opt, text="Button:").pack(side="left", padx=(12, 4))
        self.coord_button = ctk.CTkOptionMenu(
            opt, values=["left", "right", "middle"], width=90)
        self.coord_button.set("left")
        self.coord_button.pack(side="left", padx=4)

        btn = ctk.CTkFrame(tab)
        btn.pack(pady=8)
        ctk.CTkButton(
            btn, text="Run Sequence", fg_color="#22c55e",
            hover_color="#16a34a", width=140, command=self._start_coord
        ).pack(side="left", padx=8)
        ctk.CTkButton(
            btn, text="Stop", fg_color="#ef4444",
            hover_color="#dc2626", width=140, command=self._stop
        ).pack(side="left", padx=8)

    def _parse_coord_lines(self):
        text = self.coord_text.get("0.0", "end").strip()
        coords = []
        for line in text.splitlines():
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            parts = [p.strip() for p in line.split(",")]
            if len(parts) >= 2:
                x, y = int(parts[0]), int(parts[1])
                delay = float(parts[2]) if len(parts) >= 3 else 0.5
                coords.append((x, y, delay))
        return coords

    def _start_coord(self):
        if self.running:
            return
        coords = self._parse_coord_lines()
        if not coords:
            messagebox.showwarning("No Coordinates",
                                   "Add at least one coordinate line.")
            return
        loops = int(self.coord_loops_var.get() or 1)
        button = self.coord_button.get()

        def worker():
            self._log(f"Coord sequence: {len(coords)} pts x {loops} loops")
            for _ in range(loops):
                for x, y, delay in coords:
                    if not self.running:
                        break
                    pyautogui.click(x, y, button=button)
                    time.sleep(delay)
                if not self.running:
                    break
            self.running = False
            self._log("Coord sequence done.")

        self._run_thread(worker)

    def _pick_coordinate(self):
        self._log("Click anywhere to capture coordinate...")

        def on_click(x, y, button, pressed):
            if pressed:
                self.coord_text.insert("end", f"{x}, {y}, 0.5\n")
                self._log(f"Captured: ({x}, {y})")
                return False

        pynput_mouse.Listener(on_click=on_click).start()

    # ============================================================
    #  Tab 3: Record & Replay (Mouse + Keyboard)
    # ============================================================
    def _build_record_tab(self):
        tab = self.tabview.add("Record & Replay")

        ctk.CTkLabel(
            tab, text="Record mouse clicks + keyboard, then replay",
            font=("Segoe UI", 14, "bold")
        ).pack(anchor="w", padx=12, pady=(8, 4))

        btn = ctk.CTkFrame(tab)
        btn.pack(fill="x", padx=12, pady=4)
        self.rec_btn = ctk.CTkButton(
            btn, text=f"Start Recording ({HOTKEY_RECORD})",
            fg_color="#ea580c", hover_color="#c2410c",
            width=200, command=self._toggle_record)
        self.rec_btn.pack(side="left", padx=4)
        ctk.CTkButton(
            btn, text=f"Replay ({HOTKEY_REPLAY})", fg_color="#22c55e",
            hover_color="#16a34a", width=120, command=self._start_replay
        ).pack(side="left", padx=4)
        ctk.CTkButton(
            btn, text=f"Stop ({HOTKEY_STOP})", fg_color="#ef4444",
            hover_color="#dc2626", width=120, command=self._stop
        ).pack(side="left", padx=4)

        rec_opt = ctk.CTkFrame(tab)
        rec_opt.pack(fill="x", padx=12, pady=4)
        self.rec_click_var = ctk.BooleanVar(value=True)
        self.rec_move_var = ctk.BooleanVar(value=False)
        self.rec_scroll_var = ctk.BooleanVar(value=True)
        self.rec_kb_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(rec_opt, text="Clicks",
                        variable=self.rec_click_var).pack(side="left", padx=6)
        ctk.CTkCheckBox(rec_opt, text="Mouse Move",
                        variable=self.rec_move_var).pack(side="left", padx=6)
        ctk.CTkCheckBox(rec_opt, text="Scroll",
                        variable=self.rec_scroll_var).pack(side="left", padx=6)
        ctk.CTkCheckBox(rec_opt, text="Keyboard",
                        variable=self.rec_kb_var).pack(side="left", padx=6)

        # Manual insert buttons
        insert_f = ctk.CTkFrame(tab)
        insert_f.pack(fill="x", padx=12, pady=2)
        ctk.CTkLabel(insert_f, text="Insert:", text_color="gray60").pack(
            side="left", padx=4)
        ctk.CTkButton(insert_f, text="+ Wait", width=70,
                      fg_color="gray30", command=self._rec_insert_wait
                      ).pack(side="left", padx=3)
        ctk.CTkButton(insert_f, text="+ Type Text", width=90,
                      fg_color="gray30", command=self._rec_insert_type
                      ).pack(side="left", padx=3)
        ctk.CTkButton(insert_f, text="+ Image Click", width=100,
                      fg_color="gray30", command=self._rec_insert_image
                      ).pack(side="left", padx=3)
        ctk.CTkButton(insert_f, text="+ Scroll", width=80,
                      fg_color="gray30", command=self._rec_insert_scroll
                      ).pack(side="left", padx=3)
        ctk.CTkButton(insert_f, text="Delete Selected", width=110,
                      fg_color="#7f1d1d", hover_color="#991b1b",
                      command=self._rec_delete_selected
                      ).pack(side="right", padx=3)

        list_frame = ctk.CTkFrame(tab)
        list_frame.pack(fill="both", expand=True, padx=12, pady=4)

        cols = ("time", "type", "detail")
        self.rec_tree = ttk.Treeview(
            list_frame, columns=cols, show="headings",
            height=8, style="Dark.Treeview")
        for c, w, txt in zip(cols, [80, 80, 300], ["Time", "Type", "Detail"]):
            self.rec_tree.heading(c, text=txt)
            self.rec_tree.column(
                c, width=w, anchor="center" if c != "detail" else "w")
        self.rec_tree.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(
            list_frame, orient="vertical", command=self.rec_tree.yview)
        sb.pack(side="right", fill="y")
        self.rec_tree.configure(yscrollcommand=sb.set)

        ctrl = ctk.CTkFrame(tab)
        ctrl.pack(fill="x", padx=12, pady=4)
        ctk.CTkLabel(ctrl, text="Loops:").pack(side="left", padx=4)
        self.replay_loops_var = ctk.StringVar(value="1")
        ctk.CTkEntry(ctrl, textvariable=self.replay_loops_var, width=50).pack(
            side="left", padx=4)
        ctk.CTkLabel(ctrl, text="Speed:").pack(side="left", padx=(12, 4))
        self.replay_speed = ctk.CTkOptionMenu(
            ctrl, values=["0.25x", "0.5x", "1x", "2x", "4x", "8x"], width=70)
        self.replay_speed.set("1x")
        self.replay_speed.pack(side="left", padx=4)
        ctk.CTkLabel(ctrl, text="Delay:").pack(side="left", padx=(12, 4))
        self.replay_delay_var = ctk.StringVar(value="0")
        ctk.CTkEntry(ctrl, textvariable=self.replay_delay_var, width=50).pack(
            side="left", padx=2)
        self.replay_delay_unit = ctk.CTkOptionMenu(
            ctrl, values=["sec", "min", "hr"], width=60)
        self.replay_delay_unit.set("sec")
        self.replay_delay_unit.pack(side="left", padx=2)
        ctk.CTkButton(ctrl, text="Save", width=70,
                      command=self._save_recording).pack(side="right", padx=4)
        ctk.CTkButton(ctrl, text="Load", width=70,
                      command=self._load_recording).pack(side="right", padx=4)
        ctk.CTkButton(ctrl, text="Clear", width=70, fg_color="gray30",
                      command=self._clear_recording).pack(side="right", padx=4)

    def _toggle_record(self):
        if self.recording:
            self._stop_recording()
        else:
            self._start_recording()

    def _start_recording(self):
        self.recorded_events = []
        self.recording = True
        self._record_start = time.time()
        self._last_move_t = 0
        self.rec_btn.configure(text=f"Stop Recording ({HOTKEY_RECORD})")
        self._log("Recording... Press " + HOTKEY_RECORD + " to stop.")

        hotkeys = {HOTKEY_START_STOP.lower(), HOTKEY_RECORD.lower(),
                   HOTKEY_PICK.lower(), HOTKEY_REPLAY.lower(),
                   HOTKEY_STOP.lower()}
        rec_clicks = self.rec_click_var.get()
        rec_moves = self.rec_move_var.get()
        rec_scroll = self.rec_scroll_var.get()

        # Mouse listeners
        if rec_clicks or rec_moves or rec_scroll:
            self._held_buttons = {}

            def on_click(x, y, button, pressed):
                if not self.recording:
                    return False
                if not rec_clicks:
                    return
                t = round(time.time() - self._record_start, 3)
                btn = button.name if hasattr(button, 'name') else str(button)
                if pressed:
                    self._held_buttons[btn] = t
                    evt = {"t": t, "type": "mouse_down",
                           "x": x, "y": y, "button": btn}
                    self.recorded_events.append(evt)
                    self.rec_tree.insert(
                        "", "end",
                        values=(f"{t}s", "Down", f"({x},{y}) {btn}"))
                else:
                    down_t = self._held_buttons.pop(btn, t)
                    hold = round(t - down_t, 3)
                    evt = {"t": t, "type": "mouse_up",
                           "x": x, "y": y, "button": btn,
                           "hold": hold}
                    self.recorded_events.append(evt)
                    self.rec_tree.insert(
                        "", "end",
                        values=(f"{t}s", "Up",
                                f"({x},{y}) {btn} held {hold}s"))

            def on_move(x, y):
                if not self.recording or not rec_moves:
                    return
                now = time.time()
                if now - self._last_move_t < 0.05:
                    return
                self._last_move_t = now
                t = round(now - self._record_start, 3)
                evt = {"t": t, "type": "mouse_move", "x": x, "y": y}
                self.recorded_events.append(evt)
                self.rec_tree.insert(
                    "", "end",
                    values=(f"{t}s", "Move", f"({x},{y})"))

            def on_scroll(x, y, dx, dy):
                if not self.recording or not rec_scroll:
                    return
                t = round(time.time() - self._record_start, 3)
                evt = {"t": t, "type": "mouse_scroll",
                       "x": x, "y": y, "dx": dx, "dy": dy}
                self.recorded_events.append(evt)
                direction = "up" if dy > 0 else "down"
                self.rec_tree.insert(
                    "", "end",
                    values=(f"{t}s", "Scroll", f"({x},{y}) {direction} {abs(dy)}"))

            self._mouse_listener = pynput_mouse.Listener(
                on_click=on_click if rec_clicks else None,
                on_move=on_move if rec_moves else None,
                on_scroll=on_scroll if rec_scroll else None)
            self._mouse_listener.start()

        if self.rec_kb_var.get():
            self._held_keys = {}  # vk/name -> key_str

            def _pynput_key_str(key):
                """Convert pynput key to a stable string name."""
                # Special keys (ctrl, shift, alt, enter, etc.)
                if isinstance(key, pynput_kb.Key):
                    return key.name  # e.g. 'ctrl_l', 'shift', 'space'
                # Regular keys - prefer vk code (reliable even
                # when modifiers change the char)
                vk = getattr(key, 'vk', None)
                if vk is not None:
                    if 65 <= vk <= 90:   # A-Z
                        return chr(vk).lower()
                    if 48 <= vk <= 57:   # 0-9
                        return chr(vk)
                    if 112 <= vk <= 123: # F1-F12
                        return f"f{vk - 111}"
                # Fallback to char if printable
                try:
                    if key.char is not None and key.char.isprintable():
                        return key.char
                except AttributeError:
                    pass
                # Last resort
                if vk is not None:
                    return f"vk_{vk}"
                return str(key).replace("'", "")

            def _key_id(key):
                """Unique ID for dedup (vk code or name)."""
                if isinstance(key, pynput_kb.Key):
                    return key.name
                vk = getattr(key, 'vk', None)
                return vk if vk is not None else str(key)

            def on_key_press(key):
                if not self.recording:
                    return False
                t = round(time.time() - self._record_start, 3)
                key_str = _pynput_key_str(key)
                kid = _key_id(key)
                if key_str.lower() in hotkeys:
                    return
                if kid in self._held_keys:
                    return
                self._held_keys[kid] = key_str
                evt = {"t": t, "type": "key_down", "key": key_str}
                self.recorded_events.append(evt)
                self.rec_tree.insert(
                    "", "end", values=(f"{t}s", "KeyDown", key_str))

            def on_key_release(key):
                if not self.recording:
                    return False
                t = round(time.time() - self._record_start, 3)
                kid = _key_id(key)
                # Use the same key_str from press for consistency
                key_str = self._held_keys.pop(kid, None)
                if key_str is None:
                    key_str = _pynput_key_str(key)
                if key_str.lower() in hotkeys:
                    return
                evt = {"t": t, "type": "key_up", "key": key_str}
                self.recorded_events.append(evt)
                self.rec_tree.insert(
                    "", "end", values=(f"{t}s", "KeyUp", key_str))

            self._kb_listener = pynput_kb.Listener(
                on_press=on_key_press, on_release=on_key_release)
            self._kb_listener.start()

    def _stop_recording(self):
        self.recording = False
        if self._mouse_listener:
            self._mouse_listener.stop()
            self._mouse_listener = None
        if self._kb_listener:
            self._kb_listener.stop()
            self._kb_listener = None
        self.rec_btn.configure(
            text=f"Start Recording ({HOTKEY_RECORD})")
        self._log(f"Recorded {len(self.recorded_events)} events.")

    def _start_replay(self):
        if self.running:
            return
        if not self.recorded_events:
            messagebox.showinfo("Empty", "Nothing recorded yet.")
            return
        loops = int(self.replay_loops_var.get() or 1)
        speed = float(self.replay_speed.get().replace("x", ""))
        delay_val = float(self.replay_delay_var.get() or 0)
        delay_unit = self.replay_delay_unit.get()
        if delay_unit == "min":
            delay_val *= 60
        elif delay_unit == "hr":
            delay_val *= 3600
        loop_delay = max(0, delay_val)

        def worker():
            self._log(
                f"Replaying {len(self.recorded_events)} events "
                f"x {loops} at {speed}x"
                + (f" (delay {delay_val}s between loops)" if loop_delay > 0 else ""))
            for loop_i in range(loops):
                if loop_i > 0 and loop_delay > 0:
                    self._log(f"Waiting {loop_delay}s before next loop...")
                    end_t = time.time() + loop_delay
                    while time.time() < end_t and self.running:
                        time.sleep(0.1)
                prev_t = 0
                for evt in self.recorded_events:
                    if not self.running:
                        break
                    wait = (evt["t"] - prev_t) / speed
                    if wait > 0:
                        time.sleep(wait)
                    prev_t = evt["t"]
                    etype = evt["type"]

                    if etype == "mouse_click":
                        pyautogui.click(
                            evt["x"], evt["y"],
                            button=evt.get("button", "left"))
                    elif etype == "mouse_down":
                        pyautogui.moveTo(evt["x"], evt["y"])
                        pyautogui.mouseDown(
                            button=evt.get("button", "left"))
                    elif etype == "mouse_up":
                        pyautogui.moveTo(evt["x"], evt["y"])
                        pyautogui.mouseUp(
                            button=evt.get("button", "left"))
                    elif etype == "mouse_move":
                        pyautogui.moveTo(evt["x"], evt["y"])
                    elif etype == "mouse_scroll":
                        clicks = evt.get("dy", 0)
                        pyautogui.scroll(
                            clicks, x=evt.get("x"), y=evt.get("y"))
                    elif etype == "key_down":
                        try:
                            pyautogui.keyDown(self._map_key(evt["key"]))
                        except Exception:
                            pass
                    elif etype == "key_up":
                        try:
                            pyautogui.keyUp(self._map_key(evt["key"]))
                        except Exception:
                            pass
                    elif etype == "key_press":
                        try:
                            pyautogui.press(evt["key"])
                        except Exception:
                            pyautogui.write(evt["key"])
                    elif etype == "type_text":
                        pyautogui.write(
                            evt.get("text", ""),
                            interval=float(evt.get("interval", 0.02)))
                    elif etype == "wait":
                        time.sleep(float(evt.get("seconds", 1)))
                    elif etype == "image_click":
                        img = evt.get("image", "")
                        conf = float(evt.get("confidence", 0.8))
                        timeout = float(evt.get("timeout", 10))
                        deadline = time.time() + timeout
                        while time.time() < deadline and self.running:
                            try:
                                loc = pyautogui.locateCenterOnScreen(
                                    img, confidence=conf)
                                if loc:
                                    pyautogui.click(loc[0], loc[1])
                                    break
                            except Exception:
                                pass
                            time.sleep(0.5)
                if not self.running:
                    break
            self.running = False
            self._log("Replay done.")

        self._run_thread(worker)

    def _clear_recording(self):
        self.recorded_events = []
        for item in self.rec_tree.get_children():
            self.rec_tree.delete(item)
        self._log("Recording cleared.")

    def _save_recording(self):
        if not self.recorded_events:
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".json", filetypes=[("JSON", "*.json")])
        if path:
            with open(path, "w") as f:
                json.dump(self.recorded_events, f, indent=2)
            self._log(f"Saved {len(self.recorded_events)} events.")

    def _load_recording(self):
        path = filedialog.askopenfilename(
            filetypes=[("JSON", "*.json")])
        if not path or not os.path.exists(path):
            return
        self._clear_recording()
        with open(path) as f:
            self.recorded_events = json.load(f)
        for evt in self.recorded_events:
            self._rec_tree_insert(evt)
        self._log(
            f"Loaded {len(self.recorded_events)} events "
            f"from {os.path.basename(path)}")

    def _rec_evt_display(self, evt):
        """Return (type_label, detail) for a recorded event."""
        etype = evt["type"]
        if etype == "mouse_click":
            return "Click", f"({evt['x']},{evt['y']}) {evt.get('button','left')}"
        elif etype == "mouse_down":
            return "Down", f"({evt['x']},{evt['y']}) {evt.get('button','left')}"
        elif etype == "mouse_up":
            hold = evt.get("hold", 0)
            return "Up", f"({evt['x']},{evt['y']}) {evt.get('button','left')} held {hold}s"
        elif etype == "mouse_move":
            return "Move", f"({evt['x']},{evt['y']})"
        elif etype == "mouse_scroll":
            d = "up" if evt.get("dy", 0) > 0 else "down"
            return "Scroll", f"({evt['x']},{evt['y']}) {d} {abs(evt.get('dy',0))}"
        elif etype == "key_down":
            return "KeyDown", evt.get("key", "")
        elif etype == "key_up":
            return "KeyUp", evt.get("key", "")
        elif etype == "key_press":
            return "Key", evt.get("key", "")
        elif etype == "type_text":
            txt = evt.get("text", "")
            preview = txt[:40] + "..." if len(txt) > 40 else txt
            return "Type", preview
        elif etype == "wait":
            return "Wait", f"{evt.get('seconds', 1)}s"
        elif etype == "image_click":
            return "ImgClick", os.path.basename(evt.get("image", ""))
        return etype, ""

    def _rec_tree_insert(self, evt):
        label, detail = self._rec_evt_display(evt)
        self.rec_tree.insert("", "end",
                             values=(f"{evt['t']}s", label, detail))

    def _rec_last_time(self):
        if self.recorded_events:
            return self.recorded_events[-1]["t"]
        return 0.0

    def _rec_insert_wait(self):
        dlg = ctk.CTkInputDialog(
            text="Wait duration (seconds):", title="Insert Wait")
        val = dlg.get_input()
        if not val:
            return
        try:
            secs = float(val)
        except ValueError:
            return
        t = round(self._rec_last_time() + 0.01, 3)
        evt = {"t": t, "type": "wait", "seconds": secs}
        self.recorded_events.append(evt)
        self._rec_tree_insert(evt)
        self._log(f"Inserted wait: {secs}s")

    def _rec_insert_type(self):
        dlg = ctk.CTkInputDialog(
            text="Text to type:", title="Insert Type Text")
        val = dlg.get_input()
        if not val:
            return
        t = round(self._rec_last_time() + 0.01, 3)
        evt = {"t": t, "type": "type_text", "text": val, "interval": 0.02}
        self.recorded_events.append(evt)
        self._rec_tree_insert(evt)
        self._log(f"Inserted type text: {val[:30]}")

    def _rec_insert_image(self):
        path = filedialog.askopenfilename(
            filetypes=[("Images", "*.png *.jpg *.jpeg *.bmp")])
        if not path:
            return
        t = round(self._rec_last_time() + 0.01, 3)
        evt = {"t": t, "type": "image_click", "image": path,
               "confidence": 0.8, "timeout": 10}
        self.recorded_events.append(evt)
        self._rec_tree_insert(evt)
        self._log(f"Inserted image click: {os.path.basename(path)}")

    def _rec_insert_scroll(self):
        dlg = ctk.CTkToplevel(self)
        dlg.title("Insert Scroll")
        dlg.geometry("300x200")
        dlg.resizable(False, False)
        dlg.grab_set()

        ctk.CTkLabel(dlg, text="Direction:").pack(anchor="w", padx=12, pady=(10, 2))
        dir_var = ctk.StringVar(value="Down")
        ctk.CTkOptionMenu(dlg, variable=dir_var,
                          values=["Up", "Down"], width=120).pack(padx=12, anchor="w")

        ctk.CTkLabel(dlg, text="Scroll amount (clicks):").pack(anchor="w", padx=12, pady=(8, 2))
        amount_var = ctk.StringVar(value="3")
        ctk.CTkEntry(dlg, textvariable=amount_var, width=120).pack(padx=12, anchor="w")

        def confirm():
            try:
                amount = int(amount_var.get())
            except ValueError:
                return
            dy = amount if dir_var.get() == "Up" else -amount
            t = round(self._rec_last_time() + 0.01, 3)
            x, y = pyautogui.position()
            evt = {"t": t, "type": "mouse_scroll",
                   "x": x, "y": y, "dx": 0, "dy": dy}
            self.recorded_events.append(evt)
            self._rec_tree_insert(evt)
            direction = "up" if dy > 0 else "down"
            self._log(f"Inserted scroll: {direction} {abs(dy)}")
            dlg.destroy()

        ctk.CTkButton(dlg, text="OK", width=100, command=confirm).pack(pady=12)

    def _rec_delete_selected(self):
        sel = self.rec_tree.selection()
        if not sel:
            return
        items = list(self.rec_tree.get_children())
        indices = sorted([items.index(s) for s in sel], reverse=True)
        for idx in indices:
            self.recorded_events.pop(idx)
        self._rec_refresh_tree()
        self._log(f"Deleted {len(indices)} event(s).")

    def _rec_refresh_tree(self):
        for item in self.rec_tree.get_children():
            self.rec_tree.delete(item)
        for evt in self.recorded_events:
            self._rec_tree_insert(evt)

    # ============================================================
    #  Tab 4: Image Click
    # ============================================================
    def _build_image_tab(self):
        tab = self.tabview.add("Image Click")

        ctk.CTkLabel(tab, text="Click when an image appears on screen",
                     font=("Segoe UI", 14, "bold")).pack(
            anchor="w", padx=12, pady=(8, 4))

        file_f = ctk.CTkFrame(tab)
        file_f.pack(fill="x", padx=12, pady=4)
        ctk.CTkLabel(file_f, text="Image:").pack(side="left", padx=4)
        self.img_path_var = ctk.StringVar()
        ctk.CTkEntry(file_f, textvariable=self.img_path_var, width=340).pack(
            side="left", padx=4)
        ctk.CTkButton(file_f, text="Browse", width=80,
                      command=self._browse_image).pack(side="left", padx=4)

        opt = ctk.CTkFrame(tab)
        opt.pack(fill="x", padx=12, pady=4)
        ctk.CTkLabel(opt, text="Confidence:").pack(side="left", padx=4)
        self.img_conf_var = ctk.StringVar(value="0.80")
        ctk.CTkEntry(opt, textvariable=self.img_conf_var, width=60).pack(
            side="left", padx=4)
        ctk.CTkLabel(opt, text="Check every (s):").pack(
            side="left", padx=(12, 4))
        self.img_interval_var = ctk.StringVar(value="1.0")
        ctk.CTkEntry(opt, textvariable=self.img_interval_var, width=60).pack(
            side="left", padx=4)

        act = ctk.CTkFrame(tab)
        act.pack(fill="x", padx=12, pady=4)
        ctk.CTkLabel(act, text="On found:").pack(side="left", padx=4)
        self.img_action = ctk.CTkOptionMenu(
            act, values=["Click", "Double-click", "Right-click", "Move only"],
            width=130)
        self.img_action.set("Click")
        self.img_action.pack(side="left", padx=4)
        self.img_repeat_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(act, text="Keep searching",
                        variable=self.img_repeat_var).pack(
            side="left", padx=12)

        region_f = ctk.CTkFrame(tab)
        region_f.pack(fill="x", padx=12, pady=4)
        self.img_region_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(region_f, text="Limit search region:",
                        variable=self.img_region_var).pack(side="left", padx=4)
        self.img_region_vals = {}
        for lbl in ["X:", "Y:", "W:", "H:"]:
            ctk.CTkLabel(region_f, text=lbl).pack(side="left", padx=2)
            var = ctk.StringVar(value="0")
            ctk.CTkEntry(region_f, textvariable=var, width=50).pack(
                side="left", padx=2)
            self.img_region_vals[lbl] = var

        btn = ctk.CTkFrame(tab)
        btn.pack(pady=12)
        ctk.CTkButton(
            btn, text="Start Watching", fg_color="#22c55e",
            hover_color="#16a34a", width=140, command=self._start_image_click
        ).pack(side="left", padx=8)
        ctk.CTkButton(
            btn, text="Stop", fg_color="#ef4444",
            hover_color="#dc2626", width=140, command=self._stop
        ).pack(side="left", padx=8)

        ctk.CTkLabel(
            tab,
            text="Tip: Take a small screenshot of the element "
                 "you want to click, save as PNG.",
            text_color="gray60"
        ).pack(anchor="w", padx=12, pady=4)

    def _browse_image(self):
        path = filedialog.askopenfilename(
            filetypes=[("Images", "*.png *.jpg *.jpeg *.bmp")])
        if path:
            self.img_path_var.set(path)

    def _start_image_click(self):
        if self.running:
            return
        img_path = self.img_path_var.get()
        if not img_path or not os.path.exists(img_path):
            messagebox.showwarning("No Image", "Select a valid image file.")
            return

        confidence = float(self.img_conf_var.get() or 0.8)
        interval = float(self.img_interval_var.get() or 1.0)
        action = self.img_action.get()
        repeat = self.img_repeat_var.get()
        region = None
        if self.img_region_var.get():
            rx = int(self.img_region_vals["X:"].get() or 0)
            ry = int(self.img_region_vals["Y:"].get() or 0)
            rw = int(self.img_region_vals["W:"].get() or 0)
            rh = int(self.img_region_vals["H:"].get() or 0)
            if rw > 0 and rh > 0:
                region = (rx, ry, rw, rh)

        def worker():
            self._log(
                f"Image watcher: every {interval}s, conf={confidence}")
            while self.running:
                try:
                    kwargs = {"confidence": confidence}
                    if region:
                        kwargs["region"] = region
                    location = pyautogui.locateCenterOnScreen(
                        img_path, **kwargs)
                except Exception:
                    location = None

                if location:
                    x, y = location
                    self._log(f"Image found at ({x}, {y})")
                    if action == "Click":
                        pyautogui.click(x, y)
                    elif action == "Double-click":
                        pyautogui.doubleClick(x, y)
                    elif action == "Right-click":
                        pyautogui.rightClick(x, y)
                    else:
                        pyautogui.moveTo(x, y)
                    if not repeat:
                        break
                time.sleep(interval)
            self.running = False
            self._log("Image watcher stopped.")

        self._run_thread(worker)

    # ============================================================
    #  Tab 5: Web Automation
    # ============================================================
    def _build_web_tab(self):
        tab = self.tabview.add("Web Automation")

        if not HAS_SELENIUM:
            ctk.CTkLabel(
                tab, text="Selenium not installed!",
                font=("Segoe UI", 16, "bold"), text_color="#ef4444"
            ).pack(pady=30)
            ctk.CTkLabel(
                tab, text="pip install selenium webdriver-manager\n\n"
                          "Then restart.",
                font=("Consolas", 12)
            ).pack()
            return

        ctk.CTkLabel(tab, text="Browser Automation (Selenium)",
                     font=("Segoe UI", 14, "bold")).pack(
            anchor="w", padx=12, pady=(8, 4))

        url_f = ctk.CTkFrame(tab)
        url_f.pack(fill="x", padx=12, pady=4)
        ctk.CTkLabel(url_f, text="URL:").pack(side="left", padx=4)
        self.web_url_var = ctk.StringVar(value="https://")
        ctk.CTkEntry(url_f, textvariable=self.web_url_var, width=300).pack(
            side="left", padx=4)
        ctk.CTkLabel(url_f, text="Browser:").pack(side="left", padx=(8, 4))
        self.web_browser = ctk.CTkOptionMenu(
            url_f, values=["Chrome", "Edge", "Firefox"], width=100)
        self.web_browser.set("Chrome")
        self.web_browser.pack(side="left", padx=4)
        self.web_headless_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(url_f, text="Headless",
                        variable=self.web_headless_var).pack(
            side="left", padx=8)

        launch_f = ctk.CTkFrame(tab)
        launch_f.pack(fill="x", padx=12, pady=4)
        ctk.CTkButton(
            launch_f, text="Launch Browser",
            fg_color="#22c55e", hover_color="#16a34a",
            width=130, command=self._web_launch
        ).pack(side="left", padx=4)
        ctk.CTkButton(
            launch_f, text="Close Browser",
            fg_color="#ef4444", hover_color="#dc2626",
            width=130, command=self._web_close
        ).pack(side="left", padx=4)
        ctk.CTkButton(launch_f, text="Navigate", width=90,
                      command=self._web_navigate).pack(side="left", padx=4)

        # -- Command Builder --
        ctk.CTkLabel(tab, text="Command Builder",
                     font=("Segoe UI", 12, "bold"),
                     text_color="gray70").pack(anchor="w", padx=12, pady=(8, 2))

        builder_f = ctk.CTkFrame(tab)
        builder_f.pack(fill="x", padx=12, pady=4)

        ctk.CTkLabel(builder_f, text="Action:").grid(
            row=0, column=0, padx=4, pady=2, sticky="w")
        self.web_action_var = ctk.StringVar(value="click")
        self.web_action_menu = ctk.CTkOptionMenu(
            builder_f,
            values=["click", "type", "keys", "clear", "submit",
                    "wait", "scroll", "hover", "select",
                    "gettext", "assert", "iframe", "alert",
                    "screenshot", "js", "navigate",
                    "back", "forward", "refresh"],
            variable=self.web_action_var, width=110,
            command=self._web_builder_update)
        self.web_action_menu.grid(row=0, column=1, padx=4, pady=2)

        ctk.CTkLabel(builder_f, text="Locator:").grid(
            row=0, column=2, padx=4, pady=2, sticky="w")
        self.web_locator_type = ctk.CTkOptionMenu(
            builder_f, values=["css=", "id=", "name=", "xpath=", "(none)"],
            width=90)
        self.web_locator_type.set("css=")
        self.web_locator_type.grid(row=0, column=3, padx=4, pady=2)

        self.web_selector_var = ctk.StringVar()
        self.web_selector_entry = ctk.CTkEntry(
            builder_f, textvariable=self.web_selector_var,
            placeholder_text="#selector", width=150)
        self.web_selector_entry.grid(row=0, column=4, padx=4, pady=2)

        ctk.CTkLabel(builder_f, text="Value:").grid(
            row=0, column=5, padx=4, pady=2, sticky="w")
        self.web_value_var = ctk.StringVar()
        self.web_value_entry = ctk.CTkEntry(
            builder_f, textvariable=self.web_value_var,
            placeholder_text="text / key / seconds", width=150)
        self.web_value_entry.grid(row=0, column=6, padx=4, pady=2)

        ctk.CTkButton(
            builder_f, text="+ Add", width=70,
            fg_color="#22c55e", hover_color="#16a34a",
            command=self._web_add_cmd
        ).grid(row=0, column=7, padx=4, pady=2)

        self.web_pick_btn = ctk.CTkButton(
            builder_f, text="Pick Element", width=100,
            fg_color="#f59e0b", hover_color="#d97706",
            command=self._web_pick_element)
        self.web_pick_btn.grid(row=0, column=8, padx=4, pady=2)

        # -- Templates --
        tpl_f = ctk.CTkFrame(tab)
        tpl_f.pack(fill="x", padx=12, pady=2)
        ctk.CTkLabel(tpl_f, text="Templates:",
                     font=("Segoe UI", 11, "bold")).pack(side="left", padx=4)
        for label, tpl_name in [
            ("Login", "login"),
            ("Register", "register"),
            ("Search", "search"),
            ("Form Fill", "form_fill"),
            ("Scrape Text", "scrape"),
        ]:
            ctk.CTkButton(
                tpl_f, text=label, width=80, height=26,
                font=("Segoe UI", 11),
                fg_color="#6366f1", hover_color="#4f46e5",
                command=lambda t=tpl_name: self._web_load_template(t)
            ).pack(side="left", padx=2)

        # -- Quick Insert Buttons --
        quick_f = ctk.CTkFrame(tab)
        quick_f.pack(fill="x", padx=12, pady=2)
        for label, cmd_text in [
            ("Click", "click css=#element"),
            ("Type", "type id=input Hello"),
            ("Enter", "keys id=input ENTER"),
            ("Wait 1s", "wait 1.0"),
            ("Scroll", "scroll 300"),
            ("Screenshot", "screenshot capture.png"),
            ("Back", "back"),
            ("Refresh", "refresh"),
        ]:
            ctk.CTkButton(
                quick_f, text=label, width=75, height=26,
                font=("Segoe UI", 11),
                command=lambda t=cmd_text: self._web_insert_line(t)
            ).pack(side="left", padx=2)
        ctk.CTkButton(
            quick_f, text="Clear All", width=75, height=26,
            font=("Segoe UI", 11),
            fg_color="#ef4444", hover_color="#dc2626",
            command=lambda: self.web_script.delete("0.0", "end")
        ).pack(side="left", padx=2)

        # -- Script Editor --
        self.web_script = ctk.CTkTextbox(tab, font=("Consolas", 11))
        self.web_script.pack(fill="both", expand=True, padx=12, pady=4)
        self.web_script.insert("0.0",
            "# Commands: action  locator  value\n"
            "# click  css=#submit-btn\n"
            "# type   id=search  Hello World\n"
            "# keys   id=search  ENTER\n"
            "# wait   2.0\n"
            "# scroll 500\n"
            "# hover  css=#menu\n"
            "# select css=#drop  Option Text\n"
            "# screenshot capture.png\n"
            "# js     document.title\n")

        # -- Run Controls --
        run_f = ctk.CTkFrame(tab)
        run_f.pack(fill="x", padx=12, pady=(4, 8))
        ctk.CTkButton(
            run_f, text="Run Script", fg_color="#22c55e",
            hover_color="#16a34a", width=120, command=self._web_run_script
        ).pack(side="left", padx=4)
        ctk.CTkButton(
            run_f, text="Stop", fg_color="#ef4444",
            hover_color="#dc2626", width=80, command=self._stop
        ).pack(side="left", padx=4)
        ctk.CTkLabel(run_f, text="Loops:").pack(side="left", padx=(12, 4))
        self.web_loops_var = ctk.StringVar(value="1")
        ctk.CTkEntry(run_f, textvariable=self.web_loops_var, width=50).pack(
            side="left", padx=4)
        ctk.CTkLabel(run_f, text="Delay (s):").pack(
            side="left", padx=(8, 4))
        self.web_delay_var = ctk.StringVar(value="1.0")
        ctk.CTkEntry(run_f, textvariable=self.web_delay_var, width=50).pack(
            side="left", padx=4)
        ctk.CTkButton(
            run_f, text="Save Script", width=90,
            command=self._web_save_script
        ).pack(side="left", padx=(12, 4))
        ctk.CTkButton(
            run_f, text="Load Script", width=90,
            command=self._web_load_script
        ).pack(side="left", padx=4)

    # -- Web helpers --

    _PICK_JS = """
    (function() {
        if (window.__pickerActive) return 'ALREADY_ACTIVE';
        window.__pickerActive = true;
        window.__pickedSelector = null;
        var overlay = document.createElement('div');
        overlay.id = '__picker_overlay';
        overlay.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;z-index:999998;cursor:crosshair;';
        var highlight = document.createElement('div');
        highlight.id = '__picker_highlight';
        highlight.style.cssText = 'position:fixed;z-index:999997;pointer-events:none;border:2px solid #f59e0b;background:rgba(245,158,11,0.15);transition:all 0.05s;display:none;';
        var label = document.createElement('div');
        label.id = '__picker_label';
        label.style.cssText = 'position:fixed;z-index:999999;background:#1e293b;color:#f59e0b;font:12px monospace;padding:4px 8px;border-radius:4px;pointer-events:none;display:none;';
        document.body.appendChild(highlight);
        document.body.appendChild(label);
        document.body.appendChild(overlay);

        function bestSelector(el) {
            if (el.id) return '#' + el.id;
            if (el.name) return el.tagName.toLowerCase() + '[name="' + el.name + '"]';
            var cls = Array.from(el.classList).filter(function(c){return c.indexOf('__picker')===-1;});
            if (cls.length) {
                var sel = el.tagName.toLowerCase() + '.' + cls.join('.');
                if (document.querySelectorAll(sel).length === 1) return sel;
            }
            var tag = el.tagName.toLowerCase();
            var parent = el.parentElement;
            if (!parent) return tag;
            var siblings = Array.from(parent.children).filter(function(c){return c.tagName===el.tagName;});
            if (siblings.length === 1) return bestSelector(parent) + ' > ' + tag;
            var idx = siblings.indexOf(el) + 1;
            return bestSelector(parent) + ' > ' + tag + ':nth-child(' + idx + ')';
        }

        overlay.addEventListener('mousemove', function(e) {
            overlay.style.pointerEvents = 'none';
            var target = document.elementFromPoint(e.clientX, e.clientY);
            overlay.style.pointerEvents = 'auto';
            if (!target || target.id && target.id.startsWith('__picker')) return;
            var rect = target.getBoundingClientRect();
            highlight.style.display = 'block';
            highlight.style.left = rect.left + 'px';
            highlight.style.top = rect.top + 'px';
            highlight.style.width = rect.width + 'px';
            highlight.style.height = rect.height + 'px';
            var sel = bestSelector(target);
            label.style.display = 'block';
            label.textContent = sel;
            label.style.left = Math.min(e.clientX + 12, window.innerWidth - 300) + 'px';
            label.style.top = Math.max(e.clientY - 30, 4) + 'px';
        });

        overlay.addEventListener('click', function(e) {
            e.preventDefault();
            e.stopPropagation();
            overlay.style.pointerEvents = 'none';
            var target = document.elementFromPoint(e.clientX, e.clientY);
            overlay.style.pointerEvents = 'auto';
            if (target && !(target.id && target.id.startsWith('__picker'))) {
                window.__pickedSelector = bestSelector(target);
            }
            overlay.remove();
            highlight.remove();
            label.remove();
            window.__pickerActive = false;
        });
        return 'PICKER_STARTED';
    })();
    """

    _PICK_RESULT_JS = "return window.__pickedSelector;"

    def _web_pick_element(self):
        if not self.web_driver:
            self._log("Launch a browser first.")
            return
        try:
            result = self.web_driver.execute_script(self._PICK_JS)
            if result == 'ALREADY_ACTIVE':
                self._log("Picker already active. Click an element in the browser.")
                return
        except Exception as e:
            self._log(f"Pick error: {e}")
            return

        self._log("Pick mode ON - click an element in the browser...")
        self.web_pick_btn.configure(text="Waiting...", state="disabled")

        def poll_pick():
            for _ in range(300):  # 30 seconds timeout
                time.sleep(0.1)
                try:
                    sel = self.web_driver.execute_script(self._PICK_RESULT_JS)
                except Exception:
                    break
                if sel:
                    self.after(0, lambda s=sel: self._web_pick_done(s))
                    return
            self.after(0, lambda: self._web_pick_done(None))

        threading.Thread(target=poll_pick, daemon=True).start()

    def _web_pick_done(self, selector):
        self.web_pick_btn.configure(text="Pick Element", state="normal")
        if selector:
            self.web_locator_type.set("css=")
            self.web_selector_var.set(selector)
            self.web_selector_entry.configure(state="normal")
            self._log(f"Picked: css={selector}")
        else:
            self._log("Pick cancelled or timed out.")

    def _web_builder_update(self, action):
        needs_locator = action in (
            "click", "type", "keys", "clear", "submit",
            "hover", "select", "gettext", "assert", "iframe")
        no_locator = action in (
            "wait", "scroll", "screenshot", "js", "navigate",
            "alert", "back", "forward", "refresh")
        state = "normal" if needs_locator else "disabled"
        self.web_selector_entry.configure(state=state)
        if no_locator:
            self.web_locator_type.set("(none)")
        elif self.web_locator_type.get() == "(none)":
            self.web_locator_type.set("css=")

    def _web_add_cmd(self):
        action = self.web_action_var.get()
        loc_type = self.web_locator_type.get()
        selector = self.web_selector_var.get().strip()
        value = self.web_value_var.get().strip()

        if action in ("back", "forward", "refresh"):
            line = action
        elif action == "navigate":
            line = f"navigate {value or self.web_url_var.get()}"
        elif action in ("wait", "scroll"):
            line = f"{action} {value or selector}"
        elif action == "screenshot":
            line = f"screenshot {value or 'capture.png'}"
        elif action == "js":
            line = f"js {value}"
        elif action == "alert":
            line = f"alert {value or 'accept'}"
        elif action == "iframe":
            locator = "" if loc_type == "(none)" else f"{loc_type}{selector}"
            line = f"iframe {locator}" if locator else "iframe default"
        else:
            locator = "" if loc_type == "(none)" else f"{loc_type}{selector}"
            if not locator:
                self._log("Please enter a selector.")
                return
            if value:
                line = f"{action} {locator} {value}"
            else:
                line = f"{action} {locator}"

        self._web_insert_line(line)

    _WEB_TEMPLATES = {
        "login": (
            "# === Login Template ===\n"
            "# Edit the selectors & values below to match your site\n"
            "navigate https://example.com/login\n"
            "wait 1.0\n"
            "type id=username your_username\n"
            "type id=password your_password\n"
            "click css=button[type=submit]\n"
            "wait 2.0\n"
            "screenshot login_result.png\n"
        ),
        "register": (
            "# === Register Template ===\n"
            "# Edit the selectors & values below to match your site\n"
            "navigate https://example.com/register\n"
            "wait 1.0\n"
            "type id=firstname John\n"
            "type id=lastname Doe\n"
            "type id=email john@example.com\n"
            "type id=username johndoe\n"
            "type id=password MyPassword123\n"
            "type id=confirm_password MyPassword123\n"
            "click css=input[type=checkbox]\n"
            "click css=button[type=submit]\n"
            "wait 2.0\n"
            "screenshot register_result.png\n"
        ),
        "search": (
            "# === Search Template ===\n"
            "navigate https://example.com\n"
            "wait 1.0\n"
            "type name=q search term here\n"
            "keys name=q ENTER\n"
            "wait 2.0\n"
            "screenshot search_result.png\n"
        ),
        "form_fill": (
            "# === Form Fill Template ===\n"
            "navigate https://example.com/form\n"
            "wait 1.0\n"
            "type id=name John Doe\n"
            "type id=email john@example.com\n"
            "type id=phone 0123456789\n"
            "type id=address 123 Main St\n"
            "select id=country Thailand\n"
            "click css=input[type=radio][value=male]\n"
            "click css=input[type=checkbox]#agree\n"
            "click css=button[type=submit]\n"
            "wait 2.0\n"
            "screenshot form_result.png\n"
        ),
        "scrape": (
            "# === Scrape Text Template ===\n"
            "navigate https://example.com\n"
            "wait 1.0\n"
            "gettext css=h1\n"
            "gettext css=.price\n"
            "gettext css=.description\n"
            "screenshot scraped_page.png\n"
        ),
    }

    def _web_load_template(self, name):
        tpl = self._WEB_TEMPLATES.get(name, "")
        if tpl:
            self.web_script.delete("0.0", "end")
            self.web_script.insert("0.0", tpl)
            self._log(f"Template '{name}' loaded. Edit selectors to match your site.")

    def _web_insert_line(self, text):
        content = self.web_script.get("0.0", "end").strip()
        if content and not content.endswith("\n"):
            self.web_script.insert("end", "\n")
        self.web_script.insert("end", text + "\n")

    def _web_save_script(self):
        from tkinter import filedialog
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if path:
            with open(path, "w", encoding="utf-8") as f:
                f.write(self.web_script.get("0.0", "end").strip())
            self._log(f"Script saved: {path}")

    def _web_load_script(self):
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if path:
            with open(path, "r", encoding="utf-8") as f:
                content = f.read()
            self.web_script.delete("0.0", "end")
            self.web_script.insert("0.0", content)
            self._log(f"Script loaded: {path}")

    def _web_launch(self):
        if self.web_driver:
            self._log("Browser already open. Close it first.")
            return
        browser = self.web_browser.get()
        headless = self.web_headless_var.get()
        url = self.web_url_var.get().strip()

        def launch():
            try:
                self._log(f"Launching {browser}...")
                if browser == "Chrome":
                    options = webdriver.ChromeOptions()
                    if headless:
                        options.add_argument("--headless=new")
                    options.add_argument("--disable-gpu")
                    options.add_argument("--no-sandbox")
                    if HAS_WDM:
                        svc = ChromeService(ChromeDriverManager().install())
                        self.web_driver = webdriver.Chrome(
                            service=svc, options=options)
                    else:
                        self.web_driver = webdriver.Chrome(options=options)
                elif browser == "Edge":
                    from selenium.webdriver.edge.service import (
                        Service as EdgeService)
                    options = webdriver.EdgeOptions()
                    if headless:
                        options.add_argument("--headless=new")
                    try:
                        from webdriver_manager.microsoft import (
                            EdgeChromiumDriverManager)
                        svc = EdgeService(
                            EdgeChromiumDriverManager().install())
                        self.web_driver = webdriver.Edge(
                            service=svc, options=options)
                    except ImportError:
                        self.web_driver = webdriver.Edge(options=options)
                elif browser == "Firefox":
                    from selenium.webdriver.firefox.service import (
                        Service as FFService)
                    options = webdriver.FirefoxOptions()
                    if headless:
                        options.add_argument("--headless")
                    try:
                        from webdriver_manager.firefox import (
                            GeckoDriverManager)
                        svc = FFService(GeckoDriverManager().install())
                        self.web_driver = webdriver.Firefox(
                            service=svc, options=options)
                    except ImportError:
                        self.web_driver = webdriver.Firefox(options=options)

                if url and url != "https://":
                    self.web_driver.get(url)
                self._log(f"{browser} launched.")
            except Exception as e:
                self._log(f"Launch error: {e}")
                self.web_driver = None

        threading.Thread(target=launch, daemon=True).start()

    def _web_close(self):
        if self.web_driver:
            try:
                self.web_driver.quit()
            except Exception:
                pass
            self.web_driver = None
            self._log("Browser closed.")

    def _web_navigate(self):
        if not self.web_driver:
            self._log("Launch a browser first.")
            return
        url = self.web_url_var.get().strip()
        if url:
            try:
                self.web_driver.get(url)
                self._log(f"Navigated to {url}")
            except Exception as e:
                self._log(f"Nav error: {e}")

    def _web_parse_locator(self, s):
        s = s.strip()
        prefixes = {
            "css=": By.CSS_SELECTOR, "xpath=": By.XPATH,
            "id=": By.ID, "name=": By.NAME,
            "class=": By.CLASS_NAME, "tag=": By.TAG_NAME,
            "link=": By.LINK_TEXT,
        }
        for prefix, by in prefixes.items():
            if s.startswith(prefix):
                return by, s[len(prefix):]
        return By.CSS_SELECTOR, s

    def _web_find(self, locator, timeout=10):
        by, val = self._web_parse_locator(locator)
        return WebDriverWait(self.web_driver, timeout).until(
            EC.presence_of_element_located((by, val)))

    def _web_find_clickable(self, locator, timeout=10):
        by, val = self._web_parse_locator(locator)
        return WebDriverWait(self.web_driver, timeout).until(
            EC.element_to_be_clickable((by, val)))

    def _web_key(self, name):
        mapping = {
            "ENTER": Keys.ENTER, "RETURN": Keys.RETURN,
            "TAB": Keys.TAB, "ESCAPE": Keys.ESCAPE,
            "ESC": Keys.ESCAPE, "BACKSPACE": Keys.BACKSPACE,
            "DELETE": Keys.DELETE, "SPACE": Keys.SPACE,
            "UP": Keys.UP, "DOWN": Keys.DOWN,
            "LEFT": Keys.LEFT, "RIGHT": Keys.RIGHT,
            "HOME": Keys.HOME, "END": Keys.END,
            "PAGE_UP": Keys.PAGE_UP, "PAGE_DOWN": Keys.PAGE_DOWN,
            "F1": Keys.F1, "F2": Keys.F2, "F3": Keys.F3,
            "F4": Keys.F4, "F5": Keys.F5, "F9": Keys.F9,
            "F10": Keys.F10, "F11": Keys.F11, "F12": Keys.F12,
            "CTRL": Keys.CONTROL, "ALT": Keys.ALT, "SHIFT": Keys.SHIFT,
        }
        return mapping.get(name.upper(), name)

    def _web_exec_line(self, line):
        line = line.strip()
        if not line or line.startswith("#"):
            return
        parts = line.split(None, 2)
        cmd = parts[0].lower()

        if cmd == "click":
            el = self._web_find_clickable(parts[1])
            self.web_driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", el)
            time.sleep(0.3)
            try:
                el.click()
            except Exception:
                self.web_driver.execute_script("arguments[0].click();", el)
        elif cmd == "type":
            el = self._web_find_clickable(parts[1])
            self.web_driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", el)
            el.click()
            el.clear()
            el.send_keys(parts[2] if len(parts) > 2 else "")
        elif cmd == "keys":
            el = self._web_find(parts[1])
            el.send_keys(
                self._web_key(parts[2].strip() if len(parts) > 2 else ""))
        elif cmd == "wait":
            arg = parts[1] if len(parts) > 1 else "1"
            try:
                time.sleep(float(arg))
            except ValueError:
                self._web_find(arg, timeout=10)
        elif cmd == "scroll":
            px = int(parts[1]) if len(parts) > 1 else 300
            self.web_driver.execute_script(f"window.scrollBy(0,{px})")
        elif cmd == "hover":
            el = self._web_find(parts[1])
            ActionChains(self.web_driver).move_to_element(el).perform()
        elif cmd == "select":
            from selenium.webdriver.support.ui import Select
            el = self._web_find(parts[1])
            Select(el).select_by_visible_text(
                parts[2] if len(parts) > 2 else "")
        elif cmd == "screenshot":
            fname = parts[1] if len(parts) > 1 else "screenshot.png"
            self.web_driver.save_screenshot(fname)
        elif cmd == "js":
            script = line[2:].strip()
            result = self.web_driver.execute_script(f"return {script}")
            self._log(f"JS: {result}")
        elif cmd == "navigate":
            url = parts[1] if len(parts) > 1 else ""
            if url:
                self.web_driver.get(url)
                self._log(f"Navigated to {url}")
        elif cmd == "clear":
            self._web_find(parts[1]).clear()
        elif cmd == "submit":
            self._web_find(parts[1]).submit()
        elif cmd == "gettext":
            el = self._web_find(parts[1])
            text = el.text
            self._log(f"Text: {text}")
        elif cmd == "assert":
            el = self._web_find(parts[1])
            expected = parts[2] if len(parts) > 2 else ""
            actual = el.text
            if expected and expected not in actual:
                self._log(f"ASSERT FAIL: expected '{expected}' in '{actual}'")
                self.running = False
                return
            self._log(f"ASSERT OK: '{actual}'")
        elif cmd == "iframe":
            arg = parts[1] if len(parts) > 1 else "default"
            if arg == "default":
                self.web_driver.switch_to.default_content()
                self._log("Switched to main page")
            else:
                frame = self._web_find(arg)
                self.web_driver.switch_to.frame(frame)
                self._log(f"Switched to iframe: {arg}")
        elif cmd == "alert":
            action = parts[1] if len(parts) > 1 else "accept"
            alert = WebDriverWait(self.web_driver, 5).until(
                EC.alert_is_present())
            self._log(f"Alert text: {alert.text}")
            if action == "dismiss":
                alert.dismiss()
            else:
                alert.accept()
        elif cmd == "back":
            self.web_driver.back()
            self._log("Navigated back")
        elif cmd == "forward":
            self.web_driver.forward()
            self._log("Navigated forward")
        elif cmd == "refresh":
            self.web_driver.refresh()
            self._log("Page refreshed")
        else:
            self._log(f"Unknown web cmd: {cmd}")

    def _web_run_script(self):
        if not self.web_driver:
            messagebox.showwarning("No Browser",
                                   "Launch a browser first.")
            return
        if self.running:
            return
        text = self.web_script.get("0.0", "end").strip()
        lines = [l for l in text.splitlines()
                 if l.strip() and not l.strip().startswith("#")]
        if not lines:
            return

        loops = int(self.web_loops_var.get() or 1)
        delay = float(self.web_delay_var.get() or 1.0)

        def worker():
            self._log(
                f"Web script: {len(lines)} actions x {loops} loops")
            for lp in range(loops):
                for i, line in enumerate(lines):
                    if not self.running:
                        break
                    try:
                        self._web_exec_line(line)
                    except Exception as e:
                        self._log(f"Web error line {i+1}: {e}")
                        self.running = False
                        return
                if not self.running:
                    break
                if lp < loops - 1:
                    time.sleep(delay)
            self.running = False
            self._log("Web script done.")

        self._run_thread(worker)

    # ============================================================
    #  Tab 6: Macro Editor
    # ============================================================
    def _build_macro_tab(self):
        tab = self.tabview.add("Macro Editor")

        ctk.CTkLabel(tab, text="Visual Macro Builder",
                     font=("Segoe UI", 14, "bold")).pack(
            anchor="w", padx=12, pady=(8, 4))

        main = ctk.CTkFrame(tab)
        main.pack(fill="both", expand=True, padx=12, pady=4)

        # Left: step list
        left = ctk.CTkFrame(main)
        left.pack(side="left", fill="both", expand=True, padx=(0, 4))
        list_frame = ctk.CTkFrame(left)
        list_frame.pack(fill="both", expand=True)

        cols = ("#", "action", "params")
        self.macro_tree = ttk.Treeview(
            list_frame, columns=cols, show="headings",
            height=10, style="Dark.Treeview")
        self.macro_tree.heading("#", text="#")
        self.macro_tree.column("#", width=30, anchor="center")
        self.macro_tree.heading("action", text="Action")
        self.macro_tree.column("action", width=100, anchor="center")
        self.macro_tree.heading("params", text="Parameters")
        self.macro_tree.column("params", width=280, anchor="w")
        self.macro_tree.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(
            list_frame, orient="vertical", command=self.macro_tree.yview)
        sb.pack(side="right", fill="y")
        self.macro_tree.configure(yscrollcommand=sb.set)

        # Drag reorder
        self._drag_item = None
        self.macro_tree.bind("<ButtonPress-1>", self._macro_drag_start)
        self.macro_tree.bind("<B1-Motion>", self._macro_drag_motion)
        self.macro_tree.bind("<ButtonRelease-1>", self._macro_drag_end)

        # Right: buttons
        right = ctk.CTkFrame(main, width=130)
        right.pack(side="right", fill="y", padx=(4, 0))
        right.pack_propagate(False)
        for text, cmd in [
            ("Add Step", self._macro_add),
            ("Edit Step", self._macro_edit),
            ("Delete", self._macro_delete),
            ("Move Up", self._macro_move_up),
            ("Move Down", self._macro_move_down),
            ("Duplicate", self._macro_duplicate),
        ]:
            ctk.CTkButton(right, text=text, width=115,
                          command=cmd).pack(pady=3)

        # Bottom controls
        bottom = ctk.CTkFrame(tab)
        bottom.pack(fill="x", padx=12, pady=(4, 8))
        ctk.CTkButton(
            bottom, text="Run Macro", fg_color="#22c55e",
            hover_color="#16a34a", width=120, command=self._macro_run
        ).pack(side="left", padx=4)
        ctk.CTkButton(
            bottom, text="Stop", fg_color="#ef4444",
            hover_color="#dc2626", width=80, command=self._stop
        ).pack(side="left", padx=4)
        ctk.CTkLabel(bottom, text="Loops:").pack(side="left", padx=(12, 4))
        self.macro_loops_var = ctk.StringVar(value="1")
        ctk.CTkEntry(bottom, textvariable=self.macro_loops_var, width=50
                     ).pack(side="left", padx=4)
        ctk.CTkButton(bottom, text="Save", width=70,
                      command=self._macro_save).pack(side="right", padx=4)
        ctk.CTkButton(bottom, text="Load", width=70,
                      command=self._macro_load).pack(side="right", padx=4)
        ctk.CTkButton(bottom, text="Clear", width=70, fg_color="gray30",
                      command=self._macro_clear).pack(side="right", padx=4)

    def _macro_refresh(self):
        for item in self.macro_tree.get_children():
            self.macro_tree.delete(item)
        for i, step in enumerate(self.macro_steps):
            params_str = ", ".join(
                f"{k}={v}" for k, v in step.get("params", {}).items())
            self.macro_tree.insert(
                "", "end", values=(i + 1, step["action"], params_str))

    def _macro_drag_start(self, event):
        self._drag_item = self.macro_tree.identify_row(event.y)

    def _macro_drag_motion(self, event):
        if not self._drag_item:
            return
        target = self.macro_tree.identify_row(event.y)
        if target and target != self._drag_item:
            items = list(self.macro_tree.get_children())
            src = items.index(self._drag_item)
            dst = items.index(target)
            self.macro_steps[src], self.macro_steps[dst] = (
                self.macro_steps[dst], self.macro_steps[src])
            self._macro_refresh()
            new_items = list(self.macro_tree.get_children())
            if dst < len(new_items):
                self._drag_item = new_items[dst]

    def _macro_drag_end(self, _event):
        self._drag_item = None

    def _macro_sel_idx(self):
        sel = self.macro_tree.selection()
        if not sel:
            return None
        return list(self.macro_tree.get_children()).index(sel[0])

    def _macro_add(self):
        dlg = MacroStepDialog(self)
        self.wait_window(dlg)
        if dlg.result:
            self.macro_steps.append(dlg.result)
            self._macro_refresh()

    def _macro_edit(self):
        idx = self._macro_sel_idx()
        if idx is None:
            return
        dlg = MacroStepDialog(self, step=self.macro_steps[idx])
        self.wait_window(dlg)
        if dlg.result:
            self.macro_steps[idx] = dlg.result
            self._macro_refresh()

    def _macro_delete(self):
        idx = self._macro_sel_idx()
        if idx is not None:
            self.macro_steps.pop(idx)
            self._macro_refresh()

    def _macro_move_up(self):
        idx = self._macro_sel_idx()
        if idx is not None and idx > 0:
            self.macro_steps[idx], self.macro_steps[idx - 1] = (
                self.macro_steps[idx - 1], self.macro_steps[idx])
            self._macro_refresh()
            items = self.macro_tree.get_children()
            self.macro_tree.selection_set(items[idx - 1])

    def _macro_move_down(self):
        idx = self._macro_sel_idx()
        if idx is not None and idx < len(self.macro_steps) - 1:
            self.macro_steps[idx], self.macro_steps[idx + 1] = (
                self.macro_steps[idx + 1], self.macro_steps[idx])
            self._macro_refresh()
            items = self.macro_tree.get_children()
            self.macro_tree.selection_set(items[idx + 1])

    def _macro_duplicate(self):
        idx = self._macro_sel_idx()
        if idx is not None:
            self.macro_steps.insert(
                idx + 1, copy.deepcopy(self.macro_steps[idx]))
            self._macro_refresh()

    def _macro_clear(self):
        self.macro_steps.clear()
        self._macro_refresh()

    def _macro_save(self):
        if not self.macro_steps:
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON Macro", "*.json")])
        if path:
            with open(path, "w") as f:
                json.dump({"macro": self.macro_steps}, f, indent=2)
            self._log(f"Macro saved: {os.path.basename(path)}")

    def _macro_load(self):
        path = filedialog.askopenfilename(
            filetypes=[("JSON Macro", "*.json")])
        if not path or not os.path.exists(path):
            return
        with open(path) as f:
            data = json.load(f)
        if isinstance(data, dict):
            self.macro_steps = data.get("macro", [])
        elif isinstance(data, list):
            self.macro_steps = data
        self._macro_refresh()
        self._log(f"Macro loaded: {len(self.macro_steps)} steps")

    def _macro_exec_step(self, step):
        action = step["action"]
        p = step.get("params", {})

        if action == "Mouse Click":
            pyautogui.click(
                int(p.get("x", 0)), int(p.get("y", 0)),
                button=p.get("button", "left"),
                clicks=int(p.get("clicks", 1)))
        elif action == "Mouse Move":
            pyautogui.moveTo(
                int(p.get("x", 0)), int(p.get("y", 0)),
                duration=float(p.get("duration_s", 0.2)))
        elif action == "Key Press":
            pyautogui.press(p.get("key", "enter"))
        elif action == "Type Text":
            pyautogui.write(
                p.get("text", ""),
                interval=float(p.get("interval_s", 0.02)))
        elif action == "Wait":
            time.sleep(float(p.get("seconds", 1)))
        elif action == "Image Click":
            img = p.get("image_path", "")
            conf = float(p.get("confidence", 0.8))
            timeout = float(p.get("timeout_s", 10))
            deadline = time.time() + timeout
            while time.time() < deadline:
                try:
                    loc = pyautogui.locateCenterOnScreen(
                        img, confidence=conf)
                    if loc:
                        pyautogui.click(loc[0], loc[1])
                        break
                except Exception:
                    pass
                time.sleep(0.5)
        elif action == "Scroll":
            amt = int(p.get("amount", 300))
            x = int(p["x"]) if p.get("x") else None
            y = int(p["y"]) if p.get("y") else None
            pyautogui.scroll(amt, x=x, y=y)

    def _macro_run(self):
        if self.running:
            return
        if not self.macro_steps:
            messagebox.showinfo("Empty", "Add steps first.")
            return
        loops = int(self.macro_loops_var.get() or 1)

        def worker():
            self._log(
                f"Macro: {len(self.macro_steps)} steps x {loops} loops")
            for _ in range(loops):
                for i, step in enumerate(self.macro_steps):
                    if not self.running:
                        break
                    try:
                        self._macro_exec_step(step)
                    except Exception as e:
                        self._log(f"Macro error step {i+1}: {e}")
                        self.running = False
                        return
                if not self.running:
                    break
            self.running = False
            self._log("Macro done.")

        self._run_thread(worker)

    def _run_macro_file(self, path):
        """Run a saved macro file (used by scheduler)."""
        if not os.path.exists(path):
            self._log(f"Scheduler: file not found: {path}")
            return
        with open(path) as f:
            data = json.load(f)
        if isinstance(data, dict):
            steps = data.get("macro", [])
        elif isinstance(data, list):
            steps = data
        else:
            return
        if not steps:
            return

        def worker():
            self._log(f"Scheduled: {os.path.basename(path)}")
            for i, step in enumerate(steps):
                if not self.running:
                    break
                try:
                    self._macro_exec_step(step)
                except Exception as e:
                    self._log(f"Scheduled error step {i+1}: {e}")
                    self.running = False
                    return
            self.running = False
            self._log("Scheduled macro done.")

        self._run_thread(worker)

    # ============================================================
    #  Tab 7: Scheduler
    # ============================================================
    def _build_scheduler_tab(self):
        tab = self.tabview.add("Scheduler")

        ctk.CTkLabel(tab, text="Schedule Macros at Specific Times",
                     font=("Segoe UI", 14, "bold")).pack(
            anchor="w", padx=12, pady=(8, 4))

        form = ctk.CTkFrame(tab)
        form.pack(fill="x", padx=12, pady=4)

        ctk.CTkLabel(form, text="Name:").grid(
            row=0, column=0, sticky="e", padx=4, pady=4)
        self.sched_name_var = ctk.StringVar(value="Task 1")
        ctk.CTkEntry(form, textvariable=self.sched_name_var, width=140).grid(
            row=0, column=1, padx=4, pady=4)

        ctk.CTkLabel(form, text="Macro File:").grid(
            row=0, column=2, sticky="e", padx=4)
        self.sched_file_var = ctk.StringVar()
        ctk.CTkEntry(form, textvariable=self.sched_file_var, width=200).grid(
            row=0, column=3, padx=4)
        ctk.CTkButton(
            form, text="Browse", width=70,
            command=lambda: self.sched_file_var.set(
                filedialog.askopenfilename(
                    filetypes=[("JSON", "*.json")]
                ) or self.sched_file_var.get())
        ).grid(row=0, column=4, padx=4)

        ctk.CTkLabel(form, text="Time (HH:MM):").grid(
            row=1, column=0, sticky="e", padx=4, pady=4)
        self.sched_time_var = ctk.StringVar(value="09:00")
        ctk.CTkEntry(form, textvariable=self.sched_time_var, width=80).grid(
            row=1, column=1, sticky="w", padx=4, pady=4)

        ctk.CTkLabel(form, text="Repeat:").grid(
            row=1, column=2, sticky="e", padx=4)
        self.sched_repeat = ctk.CTkOptionMenu(
            form, values=["Once", "Daily", "Weekdays", "Weekends",
                          "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"],
            width=110)
        self.sched_repeat.set("Once")
        self.sched_repeat.grid(row=1, column=3, sticky="w", padx=4, pady=4)

        ctk.CTkButton(
            form, text="Add Schedule", fg_color="#22c55e",
            hover_color="#16a34a", width=120, command=self._sched_add
        ).grid(row=1, column=4, padx=4, pady=4)

        list_frame = ctk.CTkFrame(tab)
        list_frame.pack(fill="both", expand=True, padx=12, pady=4)

        cols = ("name", "file", "time", "repeat", "next_run", "on")
        self.sched_tree = ttk.Treeview(
            list_frame, columns=cols, show="headings",
            height=8, style="Dark.Treeview")
        for c, w, txt in zip(
            cols, [100, 150, 60, 80, 120, 50],
            ["Name", "Macro File", "Time", "Repeat", "Next Run", "On"]
        ):
            self.sched_tree.heading(c, text=txt)
            self.sched_tree.column(c, width=w, anchor="center")
        self.sched_tree.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(
            list_frame, orient="vertical", command=self.sched_tree.yview)
        sb.pack(side="right", fill="y")
        self.sched_tree.configure(yscrollcommand=sb.set)

        ctrl = ctk.CTkFrame(tab)
        ctrl.pack(fill="x", padx=12, pady=(4, 8))
        ctk.CTkButton(
            ctrl, text="Remove", fg_color="#ef4444",
            hover_color="#dc2626", width=100, command=self._sched_remove
        ).pack(side="left", padx=4)
        ctk.CTkButton(ctrl, text="Toggle On/Off", width=110,
                      command=self._sched_toggle).pack(side="left", padx=4)
        self._sched_status = ctk.CTkLabel(
            ctrl, text="Scheduler: active", text_color="#22c55e")
        self._sched_status.pack(side="right", padx=8)

    def _sched_add(self):
        name = self.sched_name_var.get().strip()
        fpath = self.sched_file_var.get().strip()
        time_str = self.sched_time_var.get().strip()
        repeat = self.sched_repeat.get()

        if not name or not fpath or not time_str:
            messagebox.showwarning("Missing",
                                   "Fill in name, file, and time.")
            return
        try:
            h, m = map(int, time_str.split(":"))
            if not (0 <= h < 24 and 0 <= m < 60):
                raise ValueError
        except (ValueError, TypeError):
            messagebox.showwarning("Invalid Time", "Use HH:MM (e.g. 09:00)")
            return

        self.scheduled_tasks.append({
            "name": name, "file": fpath, "time": time_str,
            "repeat": repeat, "enabled": True, "last_run": None,
        })
        self._sched_refresh()
        self._log(f"Scheduled: {name} at {time_str} ({repeat})")

    def _sched_remove(self):
        sel = self.sched_tree.selection()
        if not sel:
            return
        idx = list(self.sched_tree.get_children()).index(sel[0])
        self.scheduled_tasks.pop(idx)
        self._sched_refresh()

    def _sched_toggle(self):
        sel = self.sched_tree.selection()
        if not sel:
            return
        idx = list(self.sched_tree.get_children()).index(sel[0])
        self.scheduled_tasks[idx]["enabled"] = (
            not self.scheduled_tasks[idx]["enabled"])
        self._sched_refresh()

    def _sched_refresh(self):
        for item in self.sched_tree.get_children():
            self.sched_tree.delete(item)
        for task in self.scheduled_tasks:
            self.sched_tree.insert("", "end", values=(
                task["name"], os.path.basename(task["file"]),
                task["time"], task["repeat"],
                self._sched_next_run(task),
                "Yes" if task["enabled"] else "No"))

    def _sched_next_run(self, task):
        now = datetime.datetime.now()
        h, m = map(int, task["time"].split(":"))
        target = now.replace(hour=h, minute=m, second=0, microsecond=0)
        if target <= now:
            target += datetime.timedelta(days=1)

        repeat = task["repeat"]
        day_map = {"Mon": 0, "Tue": 1, "Wed": 2, "Thu": 3,
                   "Fri": 4, "Sat": 5, "Sun": 6}
        if repeat in day_map:
            while target.weekday() != day_map[repeat]:
                target += datetime.timedelta(days=1)
        elif repeat == "Weekdays":
            while target.weekday() >= 5:
                target += datetime.timedelta(days=1)
        elif repeat == "Weekends":
            while target.weekday() < 5:
                target += datetime.timedelta(days=1)

        return target.strftime("%Y-%m-%d %H:%M")

    def _sched_should_run(self, task):
        if not task["enabled"]:
            return False
        now = datetime.datetime.now()
        h, m = map(int, task["time"].split(":"))
        if now.hour != h or now.minute != m:
            return False
        if task["last_run"]:
            last = datetime.datetime.fromisoformat(task["last_run"])
            if (now - last).total_seconds() < 90:
                return False

        repeat = task["repeat"]
        wd = now.weekday()
        day_map = {"Mon": 0, "Tue": 1, "Wed": 2, "Thu": 3,
                   "Fri": 4, "Sat": 5, "Sun": 6}
        if repeat in ("Daily", "Once"):
            return True
        if repeat == "Weekdays" and wd < 5:
            return True
        if repeat == "Weekends" and wd >= 5:
            return True
        if repeat in day_map and wd == day_map[repeat]:
            return True
        return False

    def _start_scheduler_thread(self):
        def loop():
            while self._scheduler_active:
                for task in self.scheduled_tasks[:]:
                    if self._sched_should_run(task):
                        task["last_run"] = datetime.datetime.now().isoformat()
                        if task["repeat"] == "Once":
                            task["enabled"] = False
                        try:
                            self._sched_refresh()
                        except Exception:
                            pass
                        self._run_macro_file(task["file"])
                time.sleep(30)

        threading.Thread(target=loop, daemon=True).start()

    # ============================================================
    #  Shared Utilities
    # ============================================================
    def _log(self, msg):
        ts = time.strftime("%H:%M:%S")
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"[{ts}] {msg}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    # pynput key names -> pyautogui key names
    _KEY_MAP = {
        "ctrl_l": "ctrlleft", "ctrl_r": "ctrlright", "ctrl": "ctrl",
        "alt_l": "altleft", "alt_r": "altright", "alt": "alt",
        "alt_gr": "altright",
        "shift_l": "shiftleft", "shift_r": "shiftright", "shift": "shift",
        "cmd": "win", "cmd_l": "winleft", "cmd_r": "winright",
        "caps_lock": "capslock", "num_lock": "numlock",
        "scroll_lock": "scrolllock",
        "page_up": "pageup", "page_down": "pagedown",
        "print_screen": "printscreen",
        "enter": "enter", "return": "return",
        "space": "space", "tab": "tab",
        "backspace": "backspace", "delete": "delete",
        "up": "up", "down": "down", "left": "left", "right": "right",
        "home": "home", "end": "end",
        "esc": "escape", "escape": "escape",
        "insert": "insert",
        "menu": "apps",
    }

    def _map_key(self, key_str):
        return self._KEY_MAP.get(key_str, key_str)

    def _stop(self):
        self.running = False
        self._log("Stopped.")

    def _run_thread(self, target):
        self.running = True
        threading.Thread(target=target, daemon=True).start()

    def _register_hotkeys(self):
        kb_hotkey.add_hotkey(HOTKEY_START_STOP, self._hotkey_toggle)
        kb_hotkey.add_hotkey(HOTKEY_RECORD, self._toggle_record)
        kb_hotkey.add_hotkey(HOTKEY_PICK, self._pick_coordinate)
        kb_hotkey.add_hotkey(HOTKEY_REPLAY, self._start_replay)
        kb_hotkey.add_hotkey(HOTKEY_STOP, self._stop)

    def _hotkey_toggle(self):
        if self.running:
            self._stop()
        else:
            self._start_simple()

    def _on_close(self):
        self.running = False
        self.recording = False
        self._scheduler_active = False
        kb_hotkey.unhook_all()
        if self._mouse_listener:
            self._mouse_listener.stop()
        if self._kb_listener:
            self._kb_listener.stop()
        if self.web_driver:
            try:
                self.web_driver.quit()
            except Exception:
                pass
        self.destroy()


if __name__ == "__main__":
    app = AutoClickerPro()
    app.mainloop()
