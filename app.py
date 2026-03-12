import json
import threading
import time
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk

from pynput import keyboard, mouse


DB_FILE = Path("macros_db.json")
MIN_LOOP_INTERVAL_MS = 100
DEFAULT_STOP_KEY_ID = "key:f8"


def serialize_key(key):
    if isinstance(key, keyboard.KeyCode):
        if key.char is not None:
            return {"kind": "char", "value": key.char}
        if key.vk is not None:
            return {"kind": "vk", "value": key.vk}
        return {"kind": "vk", "value": None}
    return {"kind": "key", "value": key.name}


def deserialize_key(data):
    kind = data.get("kind")
    value = data.get("value")
    if kind == "char":
        return keyboard.KeyCode.from_char(value)
    if kind == "vk":
        return keyboard.KeyCode.from_vk(value)
    return keyboard.Key[value]


def normalize_key(key):
    if isinstance(key, keyboard.KeyCode):
        if key.vk is not None:
            return f"vk:{key.vk}"
        if key.char:
            return f"char:{key.char.lower()}"
        return "unknown"
    return f"key:{key.name}"


def key_id_to_display(key_id):
    if key_id.startswith("vk:"):
        vk = key_id.split(":", 1)[1]
        numpad_map = {
            "96": "Numpad 0",
            "97": "Numpad 1",
            "98": "Numpad 2",
            "99": "Numpad 3",
            "100": "Numpad 4",
            "101": "Numpad 5",
            "102": "Numpad 6",
            "103": "Numpad 7",
            "104": "Numpad 8",
            "105": "Numpad 9",
            "110": "Numpad .",
            "107": "Numpad +",
            "109": "Numpad -",
            "106": "Numpad *",
            "111": "Numpad /",
        }
        return numpad_map.get(vk, f"VK {vk}")
    if key_id.startswith("char:"):
        return key_id.split(":", 1)[1].upper()
    if key_id.startswith("key:"):
        return key_id.split(":", 1)[1].replace("_", " ").title()
    return key_id


def key_data_to_display(key_data):
    kind = key_data.get("kind")
    value = key_data.get("value")
    if kind == "vk":
        return key_id_to_display(f"vk:{value}")
    if kind == "char":
        return str(value).upper()
    if kind == "key":
        return key_id_to_display(f"key:{value}")
    return str(value)


def format_event_detail(event):
    device = event.get("device")
    action = event.get("action")
    if device == "keyboard":
        return f"{key_data_to_display(event['key'])}"
    if device == "mouse":
        if action == "move":
            return f"x={event['x']} y={event['y']}"
        if action == "click":
            state = "down" if event.get("pressed") else "up"
            return f"{event['button']} {state} @ ({event['x']}, {event['y']})"
        if action == "scroll":
            return f"dx={event['dx']} dy={event['dy']} @ ({event['x']}, {event['y']})"
    return "-"


class MacroRecorder:
    def __init__(self):
        self.events = []
        self._lock = threading.Lock()
        self._recording = False
        self._last_time = None
        self._kb_listener = None
        self._ms_listener = None
        self.record_mouse = False

    @property
    def recording(self):
        return self._recording

    def _elapsed(self):
        now = time.perf_counter()
        delay = 0.0 if self._last_time is None else now - self._last_time
        self._last_time = now
        return delay

    def _append_event(self, event):
        with self._lock:
            self.events.append(event)

    def _on_key_press(self, key):
        if not self._recording:
            return False
        self._append_event(
            {
                "device": "keyboard",
                "action": "press",
                "key": serialize_key(key),
                "delay": self._elapsed(),
            }
        )

    def _on_key_release(self, key):
        if not self._recording:
            return False
        self._append_event(
            {
                "device": "keyboard",
                "action": "release",
                "key": serialize_key(key),
                "delay": self._elapsed(),
            }
        )

    def _on_mouse_move(self, x, y):
        if not self._recording or not self.record_mouse:
            return False
        self._append_event(
            {
                "device": "mouse",
                "action": "move",
                "x": x,
                "y": y,
                "delay": self._elapsed(),
            }
        )

    def _on_mouse_click(self, x, y, button, pressed):
        if not self._recording or not self.record_mouse:
            return False
        self._append_event(
            {
                "device": "mouse",
                "action": "click",
                "x": x,
                "y": y,
                "button": button.name,
                "pressed": pressed,
                "delay": self._elapsed(),
            }
        )

    def _on_mouse_scroll(self, x, y, dx, dy):
        if not self._recording or not self.record_mouse:
            return False
        self._append_event(
            {
                "device": "mouse",
                "action": "scroll",
                "x": x,
                "y": y,
                "dx": dx,
                "dy": dy,
                "delay": self._elapsed(),
            }
        )

    def start(self, record_mouse=False):
        self.events = []
        self.record_mouse = record_mouse
        self._recording = True
        self._last_time = time.perf_counter()
        self._kb_listener = keyboard.Listener(
            on_press=self._on_key_press,
            on_release=self._on_key_release,
        )
        self._ms_listener = mouse.Listener(
            on_move=self._on_mouse_move,
            on_click=self._on_mouse_click,
            on_scroll=self._on_mouse_scroll,
        )
        self._kb_listener.start()
        self._ms_listener.start()

    def stop(self):
        self._recording = False
        if self._kb_listener:
            self._kb_listener.stop()
            self._kb_listener = None
        if self._ms_listener:
            self._ms_listener.stop()
            self._ms_listener = None
        with self._lock:
            return list(self.events)


class MacroPlayer:
    def __init__(self):
        self.kb = keyboard.Controller()
        self.ms = mouse.Controller()
        self._running = False
        self._looping = False
        self._stop_event = threading.Event()

    @property
    def running(self):
        return self._running

    @property
    def looping(self):
        return self._looping

    def stop(self):
        self._stop_event.set()

    def _sleep_interruptible(self, seconds):
        seconds = max(0.0, float(seconds))
        end_time = time.perf_counter() + seconds
        while time.perf_counter() < end_time:
            if self._stop_event.is_set():
                return False
            remaining = end_time - time.perf_counter()
            time.sleep(min(0.01, max(0.0, remaining)))
        return not self._stop_event.is_set()

    def _execute_event(self, event):
        device = event.get("device")
        action = event.get("action")
        if device == "keyboard":
            key = deserialize_key(event["key"])
            if action == "press":
                self.kb.press(key)
            elif action == "release":
                self.kb.release(key)
            return
        if device == "mouse":
            if action == "move":
                self.ms.position = (event["x"], event["y"])
            elif action == "click":
                self.ms.position = (event["x"], event["y"])
                button = mouse.Button[event["button"]]
                if event["pressed"]:
                    self.ms.press(button)
                else:
                    self.ms.release(button)
            elif action == "scroll":
                self.ms.position = (event["x"], event["y"])
                self.ms.scroll(event["dx"], event["dy"])

    def play(self, events, loop=False, loop_interval_ms=MIN_LOOP_INTERVAL_MS):
        if self._running:
            return False
        if not events:
            return False

        self._running = True
        self._looping = bool(loop)
        self._stop_event.clear()
        target_loop_interval_s = max(
            float(loop_interval_ms) / 1000.0,
            MIN_LOOP_INTERVAL_MS / 1000.0,
        )

        def _run():
            try:
                while not self._stop_event.is_set():
                    cycle_start = time.perf_counter()
                    for event in events:
                        if self._stop_event.is_set():
                            break
                        if not self._sleep_interruptible(event.get("delay", 0)):
                            break
                        self._execute_event(event)
                    if not loop:
                        break
                    cycle_elapsed = time.perf_counter() - cycle_start
                    cycle_wait = target_loop_interval_s - cycle_elapsed
                    if cycle_wait > 0 and not self._sleep_interruptible(cycle_wait):
                        break
            finally:
                self._running = False
                self._looping = False
                self._stop_event.clear()

        threading.Thread(target=_run, daemon=True).start()
        return True


class MacroApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PyMacro Studio")
        self.geometry("1200x760")
        self.minsize(1020, 620)

        self.recorder = MacroRecorder()
        self.player = MacroPlayer()

        self.macros = {}
        self.hotkey_map = {}
        self.stop_key_id = DEFAULT_STOP_KEY_ID

        self.mapping_key_order = []
        self.active_pressed_keys = set()
        self.selected_event_index = None

        self.capturing_hotkey = False
        self.capturing_stop_key = False

        self.name_var = tk.StringVar()
        self.record_mouse_var = tk.BooleanVar(value=False)
        self.loop_var = tk.BooleanVar(value=False)
        self.loop_interval_var = tk.StringVar(value=str(MIN_LOOP_INTERVAL_MS))
        self.delay_ms_var = tk.StringVar(value="0")
        self.status_var = tk.StringVar(value="Pronto.")
        self.event_info_var = tk.StringVar(value="Selecione um evento para editar o tempo.")
        self.stop_key_var = tk.StringVar()

        self._configure_style()
        self._build_ui()
        self._start_global_hotkey_listener()
        self.load_db(silent=True)
        self._update_stop_key_label()

    def _configure_style(self):
        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("TLabel", font=("Segoe UI", 10))
        style.configure("TButton", font=("Segoe UI", 10), padding=6)
        style.configure("TLabelframe.Label", font=("Segoe UI Semibold", 10))
        style.configure("Treeview.Heading", font=("Segoe UI Semibold", 10))
        style.configure("Treeview", rowheight=26, font=("Segoe UI", 10))

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        top = ttk.LabelFrame(self, text="Controle de Gravação e Execução", padding=12)
        top.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 8))
        for i in range(7):
            top.columnconfigure(i, weight=1)

        ttk.Label(top, text="Nome da macro").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.name_var).grid(
            row=0, column=1, columnspan=2, sticky="ew", padx=(6, 12)
        )

        ttk.Checkbutton(
            top,
            text="Gravar eventos do mouse",
            variable=self.record_mouse_var,
        ).grid(row=0, column=3, sticky="w")

        ttk.Checkbutton(
            top,
            text="Executar em loop",
            variable=self.loop_var,
        ).grid(row=0, column=4, sticky="w")

        ttk.Label(top, text="Loop (ms)").grid(row=0, column=5, sticky="e")
        ttk.Spinbox(
            top,
            from_=MIN_LOOP_INTERVAL_MS,
            to=999999,
            increment=50,
            textvariable=self.loop_interval_var,
            width=10,
        ).grid(row=0, column=6, sticky="w", padx=(6, 0))

        ttk.Button(top, text="Iniciar Gravação", command=self.start_recording).grid(
            row=1, column=0, sticky="ew", pady=(12, 0), padx=(0, 6)
        )
        ttk.Button(top, text="Parar Gravação", command=self.stop_recording).grid(
            row=1, column=1, sticky="ew", pady=(12, 0), padx=6
        )
        ttk.Button(top, text="Executar Selecionada", command=self.play_selected_macro).grid(
            row=1, column=2, sticky="ew", pady=(12, 0), padx=6
        )
        ttk.Button(top, text="Parar Execução", command=self.stop_playback).grid(
            row=1, column=3, sticky="ew", pady=(12, 0), padx=6
        )
        ttk.Button(top, text="Salvar", command=self.save_db).grid(
            row=1, column=4, sticky="ew", pady=(12, 0), padx=6
        )
        ttk.Button(top, text="Carregar", command=self.load_db).grid(
            row=1, column=5, sticky="ew", pady=(12, 0), padx=6
        )
        ttk.Button(top, text="Excluir Macro", command=self.delete_selected_macro).grid(
            row=1, column=6, sticky="ew", pady=(12, 0), padx=(6, 0)
        )

        stop_row = ttk.Frame(top)
        stop_row.grid(row=2, column=0, columnspan=7, sticky="ew", pady=(12, 0))
        stop_row.columnconfigure(1, weight=1)
        ttk.Button(stop_row, text="Tecla de Parada do Loop", command=self.capture_stop_key).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(stop_row, textvariable=self.stop_key_var).grid(
            row=0, column=1, sticky="w", padx=(12, 0)
        )

        body = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
        body.grid(row=1, column=0, sticky="nsew", padx=12, pady=8)

        left_panel = ttk.Frame(body, padding=8)
        center_panel = ttk.Frame(body, padding=8)
        right_panel = ttk.Frame(body, padding=8)
        body.add(left_panel, weight=2)
        body.add(center_panel, weight=5)
        body.add(right_panel, weight=3)

        self._build_macro_panel(left_panel)
        self._build_event_panel(center_panel)
        self._build_mapping_panel(right_panel)

        status = ttk.Frame(self, padding=(12, 6, 12, 10))
        status.grid(row=2, column=0, sticky="ew")
        status.columnconfigure(0, weight=1)
        ttk.Label(status, textvariable=self.status_var).grid(row=0, column=0, sticky="w")

    def _build_macro_panel(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        ttk.Label(
            parent,
            text="Macros Salvas",
            font=("Segoe UI Semibold", 11),
        ).grid(row=0, column=0, sticky="w", pady=(0, 6))

        list_frame = ttk.Frame(parent)
        list_frame.grid(row=1, column=0, sticky="nsew")
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.macro_listbox = tk.Listbox(
            list_frame,
            font=("Segoe UI", 10),
            activestyle="none",
            exportselection=False,
        )
        self.macro_listbox.grid(row=0, column=0, sticky="nsew")
        self.macro_listbox.bind("<<ListboxSelect>>", self._on_macro_select)

        macro_scroll = ttk.Scrollbar(
            list_frame, orient="vertical", command=self.macro_listbox.yview
        )
        macro_scroll.grid(row=0, column=1, sticky="ns")
        self.macro_listbox.configure(yscrollcommand=macro_scroll.set)

        ttk.Label(
            parent,
            text="Selecione uma macro para ver e editar os eventos.",
            foreground="#4a5568",
        ).grid(row=2, column=0, sticky="w", pady=(8, 0))

    def _build_event_panel(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        ttk.Label(
            parent,
            text="Eventos da Macro",
            font=("Segoe UI Semibold", 11),
        ).grid(row=0, column=0, sticky="w", pady=(0, 6))

        event_table_frame = ttk.Frame(parent)
        event_table_frame.grid(row=1, column=0, sticky="nsew")
        event_table_frame.columnconfigure(0, weight=1)
        event_table_frame.rowconfigure(0, weight=1)

        columns = ("idx", "delay", "total", "device", "action", "detail")
        self.event_tree = ttk.Treeview(
            event_table_frame,
            columns=columns,
            show="headings",
            selectmode="browse",
        )
        self.event_tree.grid(row=0, column=0, sticky="nsew")
        self.event_tree.bind("<<TreeviewSelect>>", self._on_event_select)

        self.event_tree.heading("idx", text="#")
        self.event_tree.heading("delay", text="Atraso (ms)")
        self.event_tree.heading("total", text="Acumulado (ms)")
        self.event_tree.heading("device", text="Dispositivo")
        self.event_tree.heading("action", text="Ação")
        self.event_tree.heading("detail", text="Detalhes")

        self.event_tree.column("idx", width=42, stretch=False, anchor="center")
        self.event_tree.column("delay", width=110, stretch=False, anchor="e")
        self.event_tree.column("total", width=130, stretch=False, anchor="e")
        self.event_tree.column("device", width=110, stretch=False, anchor="center")
        self.event_tree.column("action", width=90, stretch=False, anchor="center")
        self.event_tree.column("detail", width=400, stretch=True, anchor="w")

        event_scroll = ttk.Scrollbar(
            event_table_frame, orient="vertical", command=self.event_tree.yview
        )
        event_scroll.grid(row=0, column=1, sticky="ns")
        self.event_tree.configure(yscrollcommand=event_scroll.set)

        editor = ttk.LabelFrame(parent, text="Tempo do Evento", padding=10)
        editor.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        editor.columnconfigure(1, weight=1)

        ttk.Label(editor, textvariable=self.event_info_var).grid(
            row=0, column=0, columnspan=4, sticky="w", pady=(0, 8)
        )
        ttk.Label(editor, text="Atraso (ms)").grid(row=1, column=0, sticky="w")
        ttk.Entry(editor, textvariable=self.delay_ms_var, width=12).grid(
            row=1, column=1, sticky="w", padx=(8, 12)
        )
        ttk.Button(editor, text="Aplicar no Evento", command=self.apply_delay_to_selected).grid(
            row=1, column=2, sticky="ew"
        )
        ttk.Button(editor, text="Aplicar em Todos", command=self.apply_delay_to_all).grid(
            row=1, column=3, sticky="ew", padx=(8, 0)
        )

    def _build_mapping_panel(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        ttk.Label(
            parent,
            text="Atalhos de Teclado",
            font=("Segoe UI Semibold", 11),
        ).grid(row=0, column=0, sticky="w", pady=(0, 6))

        list_frame = ttk.Frame(parent)
        list_frame.grid(row=1, column=0, sticky="nsew")
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.mapping_listbox = tk.Listbox(
            list_frame,
            font=("Segoe UI", 10),
            activestyle="none",
            exportselection=False,
        )
        self.mapping_listbox.grid(row=0, column=0, sticky="nsew")

        mapping_scroll = ttk.Scrollbar(
            list_frame, orient="vertical", command=self.mapping_listbox.yview
        )
        mapping_scroll.grid(row=0, column=1, sticky="ns")
        self.mapping_listbox.configure(yscrollcommand=mapping_scroll.set)

        actions = ttk.Frame(parent)
        actions.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        actions.columnconfigure(0, weight=1)
        actions.columnconfigure(1, weight=1)

        ttk.Button(actions, text="Mapear Tecla", command=self.capture_hotkey_for_selected).grid(
            row=0, column=0, sticky="ew", padx=(0, 6)
        )
        ttk.Button(actions, text="Remover Mapeamento", command=self.remove_selected_mapping).grid(
            row=0, column=1, sticky="ew", padx=(6, 0)
        )

        tip = (
            "UX rápido:\n"
            "1) Selecione macro\n"
            "2) Mapeie tecla\n"
            "3) Ative loop se quiser repetição\n"
            "4) Ajuste atraso por evento na tabela"
        )
        ttk.Label(parent, text=tip, foreground="#4a5568", justify="left").grid(
            row=3, column=0, sticky="w", pady=(10, 0)
        )

    def set_status(self, text):
        self.status_var.set(text)

    def _selected_macro_name(self):
        selected = self.macro_listbox.curselection()
        if not selected:
            return None
        return self.macro_listbox.get(selected[0])

    def _get_loop_interval_ms(self):
        try:
            value = int(float(self.loop_interval_var.get()))
        except (TypeError, ValueError):
            value = MIN_LOOP_INTERVAL_MS
        if value < MIN_LOOP_INTERVAL_MS:
            value = MIN_LOOP_INTERVAL_MS
        self.loop_interval_var.set(str(value))
        return value

    def _refresh_macro_list(self, selected_name=None):
        names = sorted(self.macros.keys())
        self.macro_listbox.delete(0, tk.END)
        for name in names:
            self.macro_listbox.insert(tk.END, name)
        if selected_name and selected_name in names:
            idx = names.index(selected_name)
            self.macro_listbox.selection_clear(0, tk.END)
            self.macro_listbox.selection_set(idx)
            self.macro_listbox.activate(idx)
            self.macro_listbox.see(idx)
            self._refresh_event_table(selected_name)
        elif names:
            self.macro_listbox.selection_clear(0, tk.END)
            self.macro_listbox.selection_set(0)
            self.macro_listbox.activate(0)
            self._refresh_event_table(names[0])
        else:
            self._refresh_event_table(None)

    def _refresh_mapping_list(self):
        self.mapping_key_order = []
        self.mapping_listbox.delete(0, tk.END)
        for key_id in sorted(self.hotkey_map.keys()):
            self.mapping_key_order.append(key_id)
            macro_name = self.hotkey_map[key_id]
            self.mapping_listbox.insert(
                tk.END,
                f"{key_id_to_display(key_id)} -> {macro_name}",
            )

    def _refresh_event_table(self, macro_name):
        self.selected_event_index = None
        self.event_tree.delete(*self.event_tree.get_children())
        self.event_info_var.set("Selecione um evento para editar o tempo.")
        self.delay_ms_var.set("0")
        if not macro_name or macro_name not in self.macros:
            return

        events = self.macros[macro_name].get("events", [])
        cumulative_ms = 0.0
        for idx, event in enumerate(events):
            delay_ms = float(event.get("delay", 0.0)) * 1000.0
            cumulative_ms += delay_ms
            self.event_tree.insert(
                "",
                tk.END,
                iid=str(idx),
                values=(
                    idx + 1,
                    f"{delay_ms:.1f}",
                    f"{cumulative_ms:.1f}",
                    event.get("device", "-"),
                    event.get("action", "-"),
                    format_event_detail(event),
                ),
            )

    def _on_macro_select(self, _event=None):
        name = self._selected_macro_name()
        if not name:
            return
        self.name_var.set(name)
        self._refresh_event_table(name)

    def _on_event_select(self, _event=None):
        name = self._selected_macro_name()
        if not name or name not in self.macros:
            return
        selected = self.event_tree.selection()
        if not selected:
            self.selected_event_index = None
            self.event_info_var.set("Selecione um evento para editar o tempo.")
            return

        idx = int(selected[0])
        events = self.macros[name].get("events", [])
        if idx < 0 or idx >= len(events):
            return
        self.selected_event_index = idx
        delay_ms = float(events[idx].get("delay", 0.0)) * 1000.0
        self.delay_ms_var.set(f"{delay_ms:.1f}")
        self.event_info_var.set(
            f"Evento #{idx + 1}: {events[idx].get('device', '-')} / {events[idx].get('action', '-')}"
        )

    def _parse_delay_ms(self):
        try:
            value = float(self.delay_ms_var.get().replace(",", "."))
        except (TypeError, ValueError):
            return None
        if value < 0:
            return None
        return value

    def apply_delay_to_selected(self):
        macro_name = self._selected_macro_name()
        if not macro_name:
            messagebox.showwarning("Eventos", "Selecione uma macro.")
            return
        if self.selected_event_index is None:
            messagebox.showwarning("Eventos", "Selecione um evento na tabela.")
            return
        delay_ms = self._parse_delay_ms()
        if delay_ms is None:
            messagebox.showwarning("Eventos", "Informe um atraso em ms (>= 0).")
            return

        events = self.macros[macro_name].get("events", [])
        if self.selected_event_index >= len(events):
            return
        events[self.selected_event_index]["delay"] = delay_ms / 1000.0
        self._refresh_event_table(macro_name)
        self.event_tree.selection_set(str(self.selected_event_index))
        self.event_tree.see(str(self.selected_event_index))
        self._on_event_select()
        self.save_db(silent=True)
        self.set_status(
            f"Atraso do evento #{self.selected_event_index + 1} atualizado para {delay_ms:.1f} ms."
        )

    def apply_delay_to_all(self):
        macro_name = self._selected_macro_name()
        if not macro_name:
            messagebox.showwarning("Eventos", "Selecione uma macro.")
            return
        delay_ms = self._parse_delay_ms()
        if delay_ms is None:
            messagebox.showwarning("Eventos", "Informe um atraso em ms (>= 0).")
            return

        events = self.macros[macro_name].get("events", [])
        for event in events:
            event["delay"] = delay_ms / 1000.0
        self._refresh_event_table(macro_name)
        self.save_db(silent=True)
        self.set_status(
            f"Atraso de {len(events)} eventos atualizado para {delay_ms:.1f} ms."
        )

    def start_recording(self):
        if self.player.running:
            self.set_status("Pare a execução antes de iniciar uma nova gravação.")
            return
        name = self.name_var.get().strip()
        if not name:
            messagebox.showwarning("Macro", "Informe um nome para a macro.")
            return
        if self.recorder.recording:
            messagebox.showinfo("Macro", "A gravação já está em andamento.")
            return
        if name in self.macros and not messagebox.askyesno(
            "Sobrescrever macro",
            f"A macro '{name}' já existe. Deseja sobrescrever?",
        ):
            return
        self.recorder.start(record_mouse=self.record_mouse_var.get())
        mode = "teclado + mouse" if self.record_mouse_var.get() else "somente teclado"
        self.set_status(f"Gravando ({mode})... clique em 'Parar Gravação' ao finalizar.")

    def stop_recording(self):
        if not self.recorder.recording:
            messagebox.showinfo("Macro", "Nenhuma gravação em andamento.")
            return
        name = self.name_var.get().strip()
        if not name:
            name = "Macro sem nome"
            self.name_var.set(name)
        events = self.recorder.stop()
        self.macros[name] = {"name": name, "events": events}
        self._refresh_macro_list(selected_name=name)
        self.save_db(silent=True)
        self.set_status(f"Macro '{name}' salva com {len(events)} eventos.")

    def play_selected_macro(self):
        name = self._selected_macro_name()
        if not name:
            messagebox.showwarning("Macro", "Selecione uma macro para executar.")
            return
        self.play_macro_by_name(name)

    def play_macro_by_name(self, name):
        macro = self.macros.get(name)
        if not macro:
            return
        if self.recorder.recording:
            self.set_status("Finalize a gravação antes de executar.")
            return
        if self.player.running:
            self.set_status("Uma macro já está em execução.")
            return

        events = macro.get("events", [])
        if not events:
            self.set_status(f"A macro '{name}' não possui eventos.")
            return

        loop_enabled = self.loop_var.get()
        loop_interval_ms = self._get_loop_interval_ms()
        if loop_enabled:
            self.set_status(
                f"Executando '{name}' em loop ({loop_interval_ms} ms por ciclo, min {MIN_LOOP_INTERVAL_MS} ms). "
                f"Tecla de parada: {key_id_to_display(self.stop_key_id)}."
            )
        else:
            self.set_status(f"Executando '{name}'...")

        started = self.player.play(
            events,
            loop=loop_enabled,
            loop_interval_ms=loop_interval_ms,
        )
        if started:
            self.after(120, self._poll_player, name)

    def _poll_player(self, name):
        if self.player.running:
            self.after(120, self._poll_player, name)
            return
        self.set_status(f"Execução da macro '{name}' finalizada.")

    def stop_playback(self):
        if not self.player.running:
            self.set_status("Nenhuma macro em execução.")
            return
        self.player.stop()
        self.set_status("Parando execução...")

    def delete_selected_macro(self):
        name = self._selected_macro_name()
        if not name:
            messagebox.showwarning("Macro", "Selecione uma macro para excluir.")
            return
        if not messagebox.askyesno("Excluir macro", f"Deseja excluir a macro '{name}'?"):
            return

        del self.macros[name]
        keys_to_remove = [k for k, macro_name in self.hotkey_map.items() if macro_name == name]
        for key_id in keys_to_remove:
            del self.hotkey_map[key_id]

        self._refresh_macro_list()
        self._refresh_mapping_list()
        self.save_db(silent=True)
        self.set_status(f"Macro '{name}' excluída.")

    def _capture_single_key(self, callback):
        def _on_press(key):
            callback(key)
            return False

        listener = keyboard.Listener(on_press=_on_press)
        listener.start()
        listener.join()

    def capture_hotkey_for_selected(self):
        name = self._selected_macro_name()
        if not name:
            messagebox.showwarning("Atalho", "Selecione uma macro para mapear.")
            return
        if self.capturing_hotkey or self.capturing_stop_key:
            return

        self.capturing_hotkey = True
        self.set_status("Pressione a tecla que deve disparar a macro...")

        def _capture():
            def _done(key):
                key_id = normalize_key(key)
                self.hotkey_map[key_id] = name
                self.after(0, self._refresh_mapping_list)
                self.after(
                    0,
                    lambda: self.set_status(
                        f"Tecla '{key_id_to_display(key_id)}' mapeada para '{name}'."
                    ),
                )
                self.after(0, lambda: self.save_db(silent=True))
                self.after(0, lambda: setattr(self, "capturing_hotkey", False))

            self._capture_single_key(_done)

        threading.Thread(target=_capture, daemon=True).start()

    def capture_stop_key(self):
        if self.capturing_hotkey or self.capturing_stop_key:
            return

        self.capturing_stop_key = True
        self.set_status("Pressione a tecla para parar execução em loop...")

        def _capture():
            def _done(key):
                key_id = normalize_key(key)
                self.after(0, lambda: self._set_stop_key(key_id))
                self.after(0, lambda: setattr(self, "capturing_stop_key", False))

            self._capture_single_key(_done)

        threading.Thread(target=_capture, daemon=True).start()

    def _set_stop_key(self, key_id):
        self.stop_key_id = key_id
        self._update_stop_key_label()
        self.save_db(silent=True)
        self.set_status(f"Tecla de parada definida para '{key_id_to_display(key_id)}'.")

    def _update_stop_key_label(self):
        self.stop_key_var.set(
            f"Parada atual: {key_id_to_display(self.stop_key_id)}"
        )

    def remove_selected_mapping(self):
        selected = self.mapping_listbox.curselection()
        if not selected:
            messagebox.showwarning("Atalho", "Selecione um mapeamento para remover.")
            return
        idx = selected[0]
        if idx >= len(self.mapping_key_order):
            return
        key_id = self.mapping_key_order[idx]
        key_display = key_id_to_display(key_id)
        del self.hotkey_map[key_id]
        self._refresh_mapping_list()
        self.save_db(silent=True)
        self.set_status(f"Mapeamento '{key_display}' removido.")

    def save_db(self, silent=False):
        payload = {
            "macros": self.macros,
            "hotkey_map": self.hotkey_map,
            "stop_key_id": self.stop_key_id,
            "loop_interval_ms": self._get_loop_interval_ms(),
        }
        with DB_FILE.open("w", encoding="utf-8") as file:
            json.dump(payload, file, ensure_ascii=False, indent=2)
        if not silent:
            self.set_status(f"Dados salvos em {DB_FILE.resolve()}.")

    def load_db(self, silent=False):
        if not DB_FILE.exists():
            if not silent:
                self.set_status("Nenhum banco encontrado ainda.")
            return
        with DB_FILE.open("r", encoding="utf-8") as file:
            payload = json.load(file)

        self.macros = payload.get("macros", {})
        self.hotkey_map = payload.get("hotkey_map", {})
        self.stop_key_id = payload.get("stop_key_id", DEFAULT_STOP_KEY_ID)
        loop_interval = payload.get("loop_interval_ms", MIN_LOOP_INTERVAL_MS)
        self.loop_interval_var.set(str(max(int(loop_interval), MIN_LOOP_INTERVAL_MS)))

        self._refresh_macro_list()
        self._refresh_mapping_list()
        self._update_stop_key_label()
        if not silent:
            self.set_status("Dados carregados com sucesso.")

    def _on_global_press(self, key):
        key_id = normalize_key(key)
        if self.player.running and self.player.looping and key_id == self.stop_key_id:
            self.after(0, self.stop_playback)
            return

        if self.capturing_hotkey or self.capturing_stop_key:
            return
        if self.recorder.recording or self.player.running:
            return
        if key_id in self.active_pressed_keys:
            return

        self.active_pressed_keys.add(key_id)
        macro_name = self.hotkey_map.get(key_id)
        if macro_name:
            self.after(0, lambda: self.play_macro_by_name(macro_name))

    def _on_global_release(self, key):
        key_id = normalize_key(key)
        if key_id in self.active_pressed_keys:
            self.active_pressed_keys.remove(key_id)

    def _start_global_hotkey_listener(self):
        listener = keyboard.Listener(
            on_press=self._on_global_press,
            on_release=self._on_global_release,
        )
        listener.daemon = True
        listener.start()
        self._global_listener = listener


if __name__ == "__main__":
    app = MacroApp()
    app.mainloop()
