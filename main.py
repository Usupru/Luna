##Paquetes

import os
import sys
import ctypes
import random
import re
import socket
import subprocess as sp
import threading
import time
import urllib.request
import unicodedata
import webbrowser
from datetime import datetime
import tkinter as tk
from tkinter import Tk, Toplevel, ttk, filedialog, messagebox, StringVar, simpledialog
import wave
from piper import PiperVoice
from piper import SynthesisConfig
import json
import winreg
import keyboard
import tempfile
import uuid
import psutil
import pyaudio
import pyautogui
import pyperclip
import pyttsx3
import requests
import win32clipboard
import wikipedia
try:
    from vosk import KaldiRecognizer, Model
except Exception:
    KaldiRecognizer = None
    Model = None
from winotify import Notification
import spotipy
from spotipy import SpotifyOAuth
import sounddevice as sd
import numpy
try:
    import pystray
    from PIL import Image, ImageDraw
except Exception:
    pystray = None
    Image = None

#Preparativos

wikipedia.set_lang("es")

#Rutas y configuraciones

APP_NAME = "Luna"
IS_FROZEN = getattr(sys, "frozen", False)
APP_DIR = os.path.dirname(sys.executable if IS_FROZEN else os.path.abspath(__file__))
BUNDLE_DIR = getattr(sys, "_MEIPASS", APP_DIR)

def resolve_default_data_dir():
    local_app_data = os.environ.get("LOCALAPPDATA", "").strip()
    if local_app_data:
        return os.path.join(local_app_data, APP_NAME, "data")
    return os.path.join(APP_DIR, "data")

def resolve_default_assets_dir():
    configured = os.environ.get("LUNA_ASSETS_DIR", "").strip()
    if configured:
        return configured
    markers = ("Sonidos", "Modelos", "Imagenes")
    if any(os.path.exists(os.path.join(APP_DIR, marker)) for marker in markers):
        return APP_DIR
    if IS_FROZEN and os.path.isdir(BUNDLE_DIR):
        return BUNDLE_DIR
    return APP_DIR

DEFAULT_DATA_DIR = resolve_default_data_dir()
DATA_DIR = os.environ.get("LUNA_DATA_DIR", DEFAULT_DATA_DIR).strip() or DEFAULT_DATA_DIR

def build_data_paths(base_dir):
    return {
        "data": base_dir,
        "config": os.path.join(base_dir, "config"),
        "secrets": os.path.join(base_dir, "secrets"),
        "state": os.path.join(base_dir, "state"),
        "cache": os.path.join(base_dir, "cache"),
    }

DATA_PATHS = build_data_paths(DATA_DIR)
CONFIG_DIR = DATA_PATHS["config"]
SECRETS_DIR = DATA_PATHS["secrets"]
STATE_DIR = DATA_PATHS["state"]
CACHE_DIR = DATA_PATHS["cache"]
OPENWEATHER_API_KEY_FILE_ENV = os.environ.get("LUNA_OPENWEATHER_API_KEY_FILE", "").strip()

ASSETS_DIR = resolve_default_assets_dir()
SOUNDS_DIR = os.environ.get("LUNA_SOUNDS_DIR", os.path.join(ASSETS_DIR, "Sonidos"))
MODELS_DIR = os.environ.get("LUNA_MODELS_DIR", os.path.join(ASSETS_DIR, "Modelos"))

OPENWEATHER_API_ENV = "OPENWEATHER_API_KEY"
OPENWEATHER_API_KEY_FILE = OPENWEATHER_API_KEY_FILE_ENV or os.path.join(SECRETS_DIR, "openweather_api_key.txt")
CITY_FILE = os.path.join(STATE_DIR, "city.txt")
APP_CONFIG_FILE = os.path.join(CONFIG_DIR, "app_config.json")

DEFAULT_APP_CONFIG = {
    "setup_complete": False,
    "assistant": {
        "name": "Luna",
        "gender": "femenino",
    },
    "voice": {
        "model_path": os.path.join(MODELS_DIR, "es_AR-daniela-high.onnx"),
        "length_scale": 1.2,
        "volume": 1.0,
    },
    "apis": {
        "openweather_api_key": "",
        "spotify_client_id": "",
        "spotify_client_secret": "",
        "spotify_redirect_uri": "http://127.0.0.1:8888/callback",
    },
    "weather": {
        "city": "",
    },
    "keyword_actions": [],
    "ui": {
        "theme": "pycharm_dark",
        "run_in_background_on_close": True,
        "run_on_startup": False,
    },
}

APP_CONFIG = {}
APP_CONFIG_LOCK = threading.Lock()
APP_RUNNING = True
CONTROL_PANEL = None

def asset_path(*parts):
    return os.path.join(ASSETS_DIR, *parts)

def ensure_runtime_structure():
    os.makedirs(CONFIG_DIR, exist_ok=True)
    os.makedirs(SECRETS_DIR, exist_ok=True)
    os.makedirs(STATE_DIR, exist_ok=True)
    os.makedirs(CACHE_DIR, exist_ok=True)

    if not os.path.isfile(APP_CONFIG_FILE):
        with open(APP_CONFIG_FILE, "w", encoding="utf-8") as file:
            json.dump(DEFAULT_APP_CONFIG, file, ensure_ascii=False, indent=2)
    if not os.path.isfile(CITY_FILE):
        with open(CITY_FILE, "w", encoding="utf-8") as file:
            file.write("")
    if not os.path.isfile(OPENWEATHER_API_KEY_FILE):
        with open(OPENWEATHER_API_KEY_FILE, "w", encoding="utf-8") as file:
            file.write("")
    if not os.path.isfile(SPOTIPY_CACHE_PATH):
        with open(SPOTIPY_CACHE_PATH, "w", encoding="utf-8") as file:
            file.write("{}")

def deep_merge(base, override):
    merged = dict(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(merged.get(key), dict):
            merged[key] = deep_merge(merged[key], value)
        else:
            merged[key] = value
    return merged

def load_app_config():
    if not os.path.isfile(APP_CONFIG_FILE):
        return deep_merge(DEFAULT_APP_CONFIG, {})
    try:
        with open(APP_CONFIG_FILE, "r", encoding="utf-8") as file:
            raw = json.load(file)
        return deep_merge(DEFAULT_APP_CONFIG, raw)
    except Exception:
        return deep_merge(DEFAULT_APP_CONFIG, {})

def save_app_config(config):
    os.makedirs(CONFIG_DIR, exist_ok=True)
    with open(APP_CONFIG_FILE, "w", encoding="utf-8") as file:
        json.dump(config, file, ensure_ascii=False, indent=2)

def get_config_value(path, default=None):
    with APP_CONFIG_LOCK:
        cursor = APP_CONFIG
        for part in path:
            if not isinstance(cursor, dict) or part not in cursor:
                return default
            cursor = cursor[part]
        return cursor

def load_openweather_api_key():
    config_key = get_config_value(("apis", "openweather_api_key"), "").strip()
    if config_key:
        return config_key
    api_key = os.environ.get(OPENWEATHER_API_ENV, "").strip()
    if api_key:
        return api_key
    if os.path.isfile(OPENWEATHER_API_KEY_FILE):
        with open(OPENWEATHER_API_KEY_FILE, "r", encoding="utf-8") as file:
            return file.read().strip()
    return ""

SPOTIPY_CACHE_PATH = os.environ.get(
    "SPOTIPY_CACHE_PATH",
    os.path.join(CACHE_DIR, "spotify_token.json"),
).strip()

ensure_runtime_structure()
APP_CONFIG = load_app_config()
OPENWEATHER_API_KEY = load_openweather_api_key()

SPOTIFY_SCOPE = "user-read-playback-state user-modify-playback-state user-read-currently-playing"

def get_spotify_client():
    client_id = get_config_value(("apis", "spotify_client_id"), "").strip() or os.environ.get("SPOTIPY_CLIENT_ID", "").strip()
    client_secret = get_config_value(("apis", "spotify_client_secret"), "").strip() or os.environ.get("SPOTIPY_CLIENT_SECRET", "").strip()
    redirect_uri = get_config_value(("apis", "spotify_redirect_uri"), "").strip() or os.environ.get("SPOTIPY_REDIRECT_URI", "").strip()
    if not (client_id and client_secret and redirect_uri):
        return None
    try:
        auth = SpotifyOAuth(
            client_id=client_id,
            client_secret=client_secret,
            redirect_uri=redirect_uri,
            scope=SPOTIFY_SCOPE,
            cache_path=SPOTIPY_CACHE_PATH,
            open_browser=True,
        )
        return spotipy.Spotify(auth_manager=auth)
    except Exception:
        return None

def get_active_spotify_device(sp):
    try:
        devices = sp.devices().get("devices", [])
    except Exception:
        return None
    if not devices:
        return None
    for device in devices:
        if device.get("is_active"):
            return device
    return devices[0]

def spotify_play_query(query):
    sp = get_spotify_client()
    if not sp:
        speak("Spotify no estÃ¡ configurado. Falta credenciales.")
        return
    device = get_active_spotify_device(sp)
    if not device:
        speak("No encontrÃ© un dispositivo Spotify activo.")
        return
    results = sp.search(q=query, type="track", limit=1)
    items = results.get("tracks", {}).get("items", [])
    if not items:
        speak("No encontrÃ© esa canciÃ³n en Spotify.")
        return
    track = items[0]
    sp.start_playback(device_id=device["id"], uris=[track["uri"]])
    speak("Reproduciendo " + track.get("name", ""))

def spotify_pause():
    sp = get_spotify_client()
    if not sp:
        speak("Spotify no estÃ¡ configurado. Falta credenciales.")
        return
    device = get_active_spotify_device(sp)
    if not device:
        speak("No encontrÃ© un dispositivo Spotify activo.")
        return
    sp.pause_playback(device_id=device["id"])
    speak("Pausado.")

def spotify_resume():
    sp = get_spotify_client()
    if not sp:
        speak("Spotify no estÃ¡ configurado. Falta credenciales.")
        return
    device = get_active_spotify_device(sp)
    if not device:
        speak("No encontrÃ© un dispositivo Spotify activo.")
        return
    sp.start_playback(device_id=device["id"])
    speak("Reanudando.")

def spotify_next():
    sp = get_spotify_client()
    if not sp:
        speak("Spotify no estÃ¡ configurado. Falta credenciales.")
        return
    device = get_active_spotify_device(sp)
    if not device:
        speak("No encontrÃ© un dispositivo Spotify activo.")
        return
    sp.next_track(device_id=device["id"])
    speak("Siguiente.")

def spotify_prev():
    sp = get_spotify_client()
    if not sp:
        speak("Spotify no estÃ¡ configurado. Falta credenciales.")
        return
    device = get_active_spotify_device(sp)
    if not device:
        speak("No encontrÃ© un dispositivo Spotify activo.")
        return
    sp.previous_track(device_id=device["id"])
    speak("Anterior.")

def spotify_set_volume(percent):
    sp = get_spotify_client()
    if not sp:
        speak("Spotify no estÃ¡ configurado. Falta credenciales.")
        return
    device = get_active_spotify_device(sp)
    if not device:
        speak("No encontrÃ© un dispositivo Spotify activo.")
        return
    sp.volume(int(percent), device_id=device["id"])
    speak("Volumen " + str(int(percent)))

VOSK_MODEL_PATH = os.environ.get(
    "VOSK_MODEL_PATH",
    os.path.join(MODELS_DIR, "vosk-model-small-es-0.42"),
)

def resolve_vosk_model_path():
    candidates = []
    env_path = os.environ.get("VOSK_MODEL_PATH", "").strip()
    if env_path:
        candidates.append(env_path)
    candidates.append(VOSK_MODEL_PATH)
    try:
        if os.path.isdir(MODELS_DIR):
            for entry in os.listdir(MODELS_DIR):
                full = os.path.join(MODELS_DIR, entry)
                if os.path.isdir(full) and entry.lower().startswith("vosk-model"):
                    candidates.append(full)
    except Exception:
        pass

    seen = set()
    for candidate in candidates:
        if not candidate:
            continue
        normalized = os.path.normpath(candidate)
        if normalized in seen:
            continue
        seen.add(normalized)
        if os.path.isdir(normalized):
            return normalized
    return ""

def load_vosk_model():
    if Model is None:
        return None
    model_path = resolve_vosk_model_path()
    if model_path:
        try:
            return Model(model_path)
        except Exception:
            return None
    return None

VOSK_MODEL = None
VOSK_MODEL_LOCK = threading.Lock()

def get_vosk_model():
    global VOSK_MODEL
    if VOSK_MODEL is not None:
        return VOSK_MODEL
    with VOSK_MODEL_LOCK:
        if VOSK_MODEL is None:
            VOSK_MODEL = load_vosk_model()
    return VOSK_MODEL

THEME = {
    "bg": "#0F1115",
    "surface": "#151821",
    "surface_alt": "#1D2130",
    "accent": "#FFB86B",
    "accent_2": "#7CDFFF",
    "text": "#E7E9EE",
    "muted": "#A4A9B6",
    "border": "#2A3142",
    "danger": "#FF6B6B",
}

FONT_BASE = "Bahnschrift"
FONT_BOLD = "Bahnschrift SemiBold"
COMMON_OPENWEATHER_CITIES = [
    "Buenos Aires,AR",
    "Cordoba,AR",
    "Rosario,AR",
    "Mendoza,AR",
    "Santiago,CL",
    "Montevideo,UY",
    "Asuncion,PY",
    "Lima,PE",
    "Bogota,CO",
    "Madrid,ES",
    "Barcelona,ES",
    "Mexico City,MX",
]

def apply_modern_theme(root):
    root.configure(bg=THEME["bg"])
    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure(".", background=THEME["bg"], foreground=THEME["text"], font=(FONT_BASE, 10))
    style.configure("TFrame", background=THEME["bg"])
    style.configure("TLabel", background=THEME["bg"], foreground=THEME["text"])
    style.configure("Header.TLabel", background=THEME["bg"], foreground=THEME["accent"], font=(FONT_BOLD, 13))
    style.configure("SubHeader.TLabel", background=THEME["bg"], foreground=THEME["accent_2"], font=(FONT_BOLD, 11))
    style.configure("Muted.TLabel", background=THEME["bg"], foreground=THEME["muted"])
    style.configure("TButton", background=THEME["surface_alt"], foreground=THEME["text"], borderwidth=0, padding=(12, 6))
    style.map("TButton", background=[("active", THEME["surface"])])
    style.configure("Primary.TButton", background=THEME["accent"], foreground="#1B1B1B")
    style.map("Primary.TButton", background=[("active", "#FFC889")])
    style.configure("Ghost.TButton", background=THEME["surface"], foreground=THEME["text"])
    style.map("Ghost.TButton", background=[("active", THEME["surface_alt"])])
    style.configure("Danger.TButton", background=THEME["danger"], foreground="#1B1B1B")
    style.map("Danger.TButton", background=[("active", "#FF8686")])
    style.configure("TEntry", fieldbackground=THEME["surface_alt"], foreground=THEME["text"], insertcolor=THEME["text"], bordercolor=THEME["border"])
    style.configure("TCombobox", fieldbackground=THEME["surface_alt"], foreground=THEME["text"])
    style.map("TCombobox", fieldbackground=[("readonly", THEME["surface_alt"])], foreground=[("readonly", THEME["text"])])
    style.configure("Treeview", background=THEME["surface_alt"], fieldbackground=THEME["surface_alt"], foreground=THEME["text"])
    style.map("Treeview", background=[("selected", "#2B3953")])
    style.configure("Treeview.Heading", background=THEME["surface"], foreground=THEME["text"], font=(FONT_BOLD, 10))

def center_window(window, width, height):
    window.update_idletasks()
    screen_w = window.winfo_screenwidth()
    screen_h = window.winfo_screenheight()
    x = max(0, int((screen_w - width) / 2))
    y = max(0, int((screen_h - height) / 2))
    window.geometry(f"{width}x{height}+{x}+{y}")

def set_rounded_corners(window):
    try:
        DWMWA_WINDOW_CORNER_PREFERENCE = 33
        DWM_WINDOW_CORNER_PREFERENCE = 3
        hwnd = window.winfo_id()
        preference = ctypes.c_int(DWM_WINDOW_CORNER_PREFERENCE)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(
            hwnd,
            DWMWA_WINDOW_CORNER_PREFERENCE,
            ctypes.byref(preference),
            ctypes.sizeof(preference),
        )
    except Exception:
        pass

def _enable_drag(widget, window):
    def on_press(event):
        window._drag_start_x = event.x_root
        window._drag_start_y = event.y_root
        window._drag_win_x = window.winfo_x()
        window._drag_win_y = window.winfo_y()

    def on_drag(event):
        dx = event.x_root - window._drag_start_x
        dy = event.y_root - window._drag_start_y
        window.geometry(f"+{window._drag_win_x + dx}+{window._drag_win_y + dy}")

    widget.bind("<ButtonPress-1>", on_press)
    widget.bind("<B1-Motion>", on_drag)

def build_window_shell(window, title, on_close=None, on_minimize=None):
    container = tk.Frame(window, bg=THEME["bg"], highlightthickness=1, highlightbackground=THEME["border"])
    container.pack(fill="both", expand=True)

    titlebar = tk.Frame(container, bg=THEME["surface"], height=40)
    titlebar.pack(fill="x", side="top")

    title_label = tk.Label(titlebar, text=title, bg=THEME["surface"], fg=THEME["text"], font=(FONT_BOLD, 11))
    title_label.pack(side="left", padx=12)

    actions = tk.Frame(titlebar, bg=THEME["surface"])
    actions.pack(side="right", padx=8)

    def do_minimize():
        if on_minimize:
            on_minimize()
        else:
            window.withdraw()

    def do_close():
        if on_close:
            on_close()
        else:
            window.destroy()

    min_btn = tk.Button(actions, text="-", command=do_minimize, bg=THEME["surface"], fg=THEME["text"],
                        activebackground=THEME["surface_alt"], activeforeground=THEME["text"], bd=0, font=(FONT_BOLD, 10))
    min_btn.pack(side="left", padx=(0, 6))
    close_btn = tk.Button(actions, text="X", command=do_close, bg=THEME["surface"], fg=THEME["text"],
                          activebackground=THEME["danger"], activeforeground="#1B1B1B", bd=0, font=(FONT_BOLD, 10))
    close_btn.pack(side="left")

    _enable_drag(titlebar, window)
    _enable_drag(title_label, window)

    content = tk.Frame(container, bg=THEME["bg"])
    content.pack(fill="both", expand=True, padx=16, pady=16)
    return content

def create_modern_window(window, title, size, minsize=None, on_close=None, resizable=True, on_minimize=None):
    window.overrideredirect(True)
    window.configure(bg=THEME["bg"])
    apply_modern_theme(window)
    if minsize:
        window.minsize(minsize[0], minsize[1])
    if size:
        center_window(window, size[0], size[1])
    window.resizable(resizable, resizable)
    content = build_window_shell(window, title, on_close=on_close, on_minimize=on_minimize)
    if on_close:
        window.bind("<Escape>", lambda _event: on_close())
    window.update_idletasks()
    set_rounded_corners(window)
    return content

def set_run_on_startup(enabled):
    try:
        if IS_FROZEN:
            startup_command = f"\"{os.path.abspath(sys.executable)}\""
        else:
            startup_command = f"\"{os.path.abspath(sys.executable)}\" \"{os.path.abspath(sys.argv[0])}\""
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_SET_VALUE)
        if enabled:
            winreg.SetValueEx(key, "Luna", 0, winreg.REG_SZ, startup_command)
        else:
            try:
                winreg.DeleteValue(key, "Luna")
            except FileNotFoundError:
                pass
        winreg.CloseKey(key)
    except Exception:
        pass

def launch_program_path(program_path):
    path = program_path.strip()
    if not path:
        return False
    if re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", path):
        try:
            os.startfile(path)
        except Exception:
            return False
        return True
    if not os.path.isfile(path):
        return False
    extension = os.path.splitext(path)[1].lower()
    try:
        if extension in (".lnk", ".url"):
            os.startfile(path)
        else:
            sp.Popen([path])
    except Exception:
        return False
    return True

def normalize_keywords_list(raw_keywords):
    keywords = []
    for item in raw_keywords:
        if not isinstance(item, str):
            continue
        normalized = normalize_intent_text(item)
        if normalized:
            keywords.append(normalized)
    return keywords

def load_keyword_actions():
    raw_actions = get_config_value(("keyword_actions",), []) or []
    actions = []
    for entry in raw_actions:
        if not isinstance(entry, dict):
            continue
        keywords = normalize_keywords_list(entry.get("keywords", []))
        if not keywords:
            continue
        action = entry.get("action", {}) if isinstance(entry.get("action"), dict) else {}
        action_type = action.get("type", "")
        if action_type not in ("response_fixed", "response_random", "launch_app", "run_command"):
            continue
        actions.append({
            "id": entry.get("id") or str(uuid.uuid4()),
            "keywords": keywords,
            "run_on_start": bool(entry.get("run_on_start", False)),
            "action": {
                "type": action_type,
                "response": action.get("response", ""),
                "responses": action.get("responses", []) if isinstance(action.get("responses", []), list) else [],
                "exe_path": action.get("exe_path", ""),
                "command": action.get("command", ""),
            },
        })
    return actions

def format_keyword_action(entry):
    action = entry.get("action", {})
    action_type = action.get("type", "")
    prefix = "Inicio + " if entry.get("run_on_start") else ""
    if action_type == "response_fixed":
        return f"{prefix}Respuesta fija"
    if action_type == "response_random":
        count = len([r for r in action.get("responses", []) if isinstance(r, str) and r.strip()])
        return f"{prefix}Respuesta random ({count})"
    if action_type == "launch_app":
        raw_target = action.get("exe_path", "")
        program = raw_target if re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", raw_target) else os.path.basename(raw_target)
        program = program or ".exe/.lnk/.url/steam://"
        return f"{prefix}Abrir {program}"
    if action_type == "run_command":
        return f"{prefix}Ejecutar comando"
    return "Accion"

def execute_keyword_action(entry, warn_on_fail=False, announce_launch=True):
    action = entry.get("action", {}) if isinstance(entry.get("action"), dict) else {}
    action_type = action.get("type", "")
    if action_type == "response_fixed":
        response = action.get("response", "").strip()
        if not response:
            responses = [r for r in action.get("responses", []) if isinstance(r, str) and r.strip()]
            response = responses[0] if responses else ""
        if response:
            speak(response)
            return True
    elif action_type == "response_random":
        responses = [r for r in action.get("responses", []) if isinstance(r, str) and r.strip()]
        if responses:
            speak(random.choice(responses))
            return True
    elif action_type == "launch_app":
        exe_path = action.get("exe_path", "").strip()
        if launch_program_path(exe_path):
            if announce_launch:
                speak("Abriendo programa")
            return True
        if warn_on_fail:
            speak("No encontre el programa configurado.")
    elif action_type == "run_command":
        command = action.get("command", "").strip()
        if command:
            sp.Popen(command, shell=True)
            speak("Ejecutando comando")
            return True
        if warn_on_fail:
            speak("No hay comando configurado.")
    return False

def handle_keyword_actions(intent_statement):
    actions = get_config_value(("keyword_actions",), []) or []
    for entry in actions:
        keywords = normalize_keywords_list(entry.get("keywords", []))
        if not keywords:
            continue
        if not any(keyword in intent_statement for keyword in keywords):
            continue
        executed = execute_keyword_action(entry, warn_on_fail=True)
        if executed:
            return True
    return False

def run_startup_keyword_actions():
    actions = get_config_value(("keyword_actions",), []) or []
    for entry in actions:
        if entry.get("run_on_start"):
            execute_keyword_action(entry, announce_launch=False)

class KeywordActionDialog:
    ACTION_LABELS = {
        "Respuesta fija": "response_fixed",
        "Respuesta random (lista)": "response_random",
        "Iniciar programa o enlace": "launch_app",
        "Ejecutar comando": "run_command",
    }
    ACTION_LABELS_REVERSE = {v: k for k, v in ACTION_LABELS.items()}

    def __init__(self, parent, existing=None):
        self.result = None
        self.window = Toplevel(parent)
        self.window.title("Accion por palabra clave")
        self.window.protocol("WM_DELETE_WINDOW", self._cancel)
        content = create_modern_window(self.window, "Accion por palabra clave", (640, 520), (600, 480), on_close=self._cancel, resizable=False)

        self.keywords_var = StringVar()
        self.action_type_var = StringVar()
        self.exe_path_var = StringVar()
        self.command_var = StringVar()
        self.run_on_start_var = tk.BooleanVar(value=False)

        form = ttk.Frame(content)
        form.pack(fill="both", expand=True)

        ttk.Label(form, text="Palabras clave (separadas por coma)", style="SubHeader.TLabel").pack(anchor="w")
        ttk.Entry(form, textvariable=self.keywords_var).pack(fill="x", pady=(4, 12))

        ttk.Label(form, text="Accion", style="SubHeader.TLabel").pack(anchor="w")
        action_box = ttk.Combobox(
            form,
            textvariable=self.action_type_var,
            values=list(self.ACTION_LABELS.keys()),
            state="readonly",
        )
        action_box.pack(fill="x", pady=(4, 12))
        action_box.bind("<<ComboboxSelected>>", lambda _event: self._update_action_fields())

        self.response_frame = ttk.Frame(form)
        ttk.Label(self.response_frame, text="Respuesta(s)", style="SubHeader.TLabel").pack(anchor="w")
        self.response_text = tk.Text(
            self.response_frame,
            height=6,
            bg=THEME["surface_alt"],
            fg=THEME["text"],
            insertbackground=THEME["text"],
            relief="flat",
        )
        self.response_text.pack(fill="both", expand=True, pady=(4, 0))
        ttk.Label(self.response_frame, text="Para respuesta random, cada linea es una opcion.", style="Muted.TLabel").pack(anchor="w", pady=(6, 0))

        self.exe_frame = ttk.Frame(form)
        ttk.Label(self.exe_frame, text="Programa/enlace (.exe, .lnk, .url o steam://...)", style="SubHeader.TLabel").pack(anchor="w")
        exe_row = ttk.Frame(self.exe_frame)
        exe_row.pack(fill="x", pady=(4, 0))
        ttk.Entry(exe_row, textvariable=self.exe_path_var).pack(side="left", fill="x", expand=True)
        ttk.Button(exe_row, text="Examinar", command=self._pick_exe, style="Ghost.TButton").pack(side="left", padx=(8, 0))
        ttk.Label(self.exe_frame, text="Tambien puedes escribir una URI como steam://rungameid/553850", style="Muted.TLabel").pack(anchor="w", pady=(6, 0))

        self.command_frame = ttk.Frame(form)
        ttk.Label(self.command_frame, text="Comando", style="SubHeader.TLabel").pack(anchor="w")
        ttk.Entry(self.command_frame, textvariable=self.command_var).pack(fill="x", pady=(4, 0))

        self.run_on_start_chk = ttk.Checkbutton(
            form,
            text="Ejecutar al iniciar Luna",
            variable=self.run_on_start_var,
        )
        self.run_on_start_chk.pack(anchor="w", pady=(0, 12))

        actions = ttk.Frame(content)
        actions.pack(fill="x", pady=(12, 0))
        ttk.Button(actions, text="Cancelar", command=self._cancel, style="Ghost.TButton").pack(side="right")
        ttk.Button(actions, text="Guardar", command=self._save, style="Primary.TButton").pack(side="right", padx=(0, 8))

        if existing:
            self._load_existing(existing)
        else:
            self.action_type_var.set("Respuesta fija")
            self._update_action_fields()

        self.window.grab_set()

    def _load_existing(self, existing):
        keywords = ", ".join(existing.get("keywords", []))
        self.keywords_var.set(keywords)
        action = existing.get("action", {})
        action_type = action.get("type", "response_fixed")
        self.action_type_var.set(self.ACTION_LABELS_REVERSE.get(action_type, "Respuesta fija"))
        if action_type == "response_fixed":
            response = action.get("response", "")
            self.response_text.delete("1.0", "end")
            self.response_text.insert("1.0", response)
        elif action_type == "response_random":
            responses = action.get("responses", [])
            self.response_text.delete("1.0", "end")
            self.response_text.insert("1.0", "\n".join([r for r in responses if isinstance(r, str)]))
        elif action_type == "launch_app":
            self.exe_path_var.set(action.get("exe_path", ""))
        elif action_type == "run_command":
            self.command_var.set(action.get("command", ""))
        self.run_on_start_var.set(bool(existing.get("run_on_start", False)))
        self._update_action_fields()

    def _update_action_fields(self):
        for frame in (self.response_frame, self.exe_frame, self.command_frame):
            frame.pack_forget()
        action_type = self.ACTION_LABELS.get(self.action_type_var.get(), "response_fixed")
        if action_type in ("response_fixed", "response_random"):
            self.response_frame.pack(fill="both", expand=True, pady=(0, 12))
        elif action_type == "launch_app":
            self.exe_frame.pack(fill="x", pady=(0, 12))
        elif action_type == "run_command":
            self.command_frame.pack(fill="x", pady=(0, 12))

    def _pick_exe(self):
        chosen = filedialog.askopenfilename(
            title="Seleccionar programa o acceso directo",
            filetypes=[("Programas y accesos", "*.exe *.lnk *.url"), ("Ejecutable", "*.exe"), ("Acceso directo", "*.lnk"), ("Acceso de internet", "*.url"), ("Todos", "*.*")],
        )
        if chosen:
            self.exe_path_var.set(chosen)

    def _save(self):
        keywords_raw = self.keywords_var.get().split(",")
        keywords = normalize_keywords_list(keywords_raw)
        if not keywords:
            messagebox.showinfo("Accion", "Ingresa al menos una palabra clave.")
            return
        action_type = self.ACTION_LABELS.get(self.action_type_var.get(), "response_fixed")
        action = {"type": action_type, "response": "", "responses": [], "exe_path": "", "command": ""}

        if action_type == "response_fixed":
            response = self.response_text.get("1.0", "end").strip()
            if not response:
                messagebox.showinfo("Accion", "Ingresa una respuesta fija.")
                return
            action["response"] = response
        elif action_type == "response_random":
            responses = [line.strip() for line in self.response_text.get("1.0", "end").splitlines() if line.strip()]
            if not responses:
                messagebox.showinfo("Accion", "Ingresa al menos una respuesta.")
                return
            action["responses"] = responses
        elif action_type == "launch_app":
            exe_path = self.exe_path_var.get().strip()
            if re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", exe_path):
                action["exe_path"] = exe_path
                self.result = {
                    "id": str(uuid.uuid4()),
                    "keywords": keywords,
                    "run_on_start": bool(self.run_on_start_var.get()),
                    "action": action,
                }
                self.window.destroy()
                return
            extension = os.path.splitext(exe_path)[1].lower()
            if not exe_path or extension not in (".exe", ".lnk", ".url"):
                messagebox.showinfo("Accion", "Selecciona un .exe/.lnk/.url o escribe una URI (por ejemplo steam://...).")
                return
            if not os.path.isfile(exe_path):
                messagebox.showinfo("Accion", "La ruta seleccionada no existe.")
                return
            action["exe_path"] = exe_path
        elif action_type == "run_command":
            command = self.command_var.get().strip()
            if not command:
                messagebox.showinfo("Accion", "Ingresa un comando.")
                return
            action["command"] = command

        self.result = {
            "id": str(uuid.uuid4()),
            "keywords": keywords,
            "run_on_start": bool(self.run_on_start_var.get()),
            "action": action,
        }
        self.window.destroy()

    def _cancel(self):
        self.window.destroy()

class SetupWizard:
    def __init__(self, root):
        self.root = root
        self.window = Tk() if root is None else Toplevel(root)
        self.window.title("Luna Setup")
        self.content = create_modern_window(
            self.window,
            "Luna Setup",
            (900, 660),
            (820, 600),
            on_close=self.window.destroy,
            on_minimize=self.window.withdraw,
            resizable=True,
        )

        self.assistant_name = StringVar(value=get_config_value(("assistant", "name"), "Luna"))
        self.assistant_gender = StringVar(value=get_config_value(("assistant", "gender"), "femenino"))
        self.voice_model_path = StringVar(value=get_config_value(("voice", "model_path"), DEFAULT_APP_CONFIG["voice"]["model_path"]))
        self.voice_length_scale = StringVar(value=str(get_config_value(("voice", "length_scale"), 1.2)))
        self.spotify_client_id = StringVar(value=get_config_value(("apis", "spotify_client_id"), ""))
        self.spotify_client_secret = StringVar(value=get_config_value(("apis", "spotify_client_secret"), ""))
        self.spotify_redirect = StringVar(value=get_config_value(("apis", "spotify_redirect_uri"), "http://127.0.0.1:8888/callback"))
        self.openweather_key = StringVar(value=get_config_value(("apis", "openweather_api_key"), ""))
        self.designated_city = StringVar(value=read_city())
        self.run_on_startup_var = tk.BooleanVar(value=bool(get_config_value(("ui", "run_on_startup"), False)))

        self.keyword_actions = load_keyword_actions()
        self.current_step = 0
        self.steps = []

        self._build_layout()
        self._show_step(0)
        self.window.after(50, self._show_on_top)

    def _show_on_top(self):
        self.window.deiconify()
        self.window.attributes("-topmost", True)
        self.window.lift()
        self.window.focus_force()
        self.window.after(200, lambda: self.window.attributes("-topmost", False))

    def _build_layout(self):
        container = ttk.Frame(self.content, padding=8)
        container.pack(fill="both", expand=True)

        header = tk.Frame(container, bg=THEME["surface_alt"], highlightthickness=1, highlightbackground=THEME["border"])
        header.pack(fill="x", pady=(0, 12))
        accent = tk.Frame(header, bg=THEME["accent"], width=8, height=64)
        accent.pack(side="left", fill="y")
        header_content = tk.Frame(header, bg=THEME["surface_alt"])
        header_content.pack(side="left", fill="x", expand=True, padx=12, pady=10)
        tk.Label(header_content, text="Luna Setup", bg=THEME["surface_alt"], fg=THEME["text"], font=(FONT_BOLD, 16)).pack(anchor="w")
        tk.Label(header_content, text="Configura tu asistente como un launcher de juegos.", bg=THEME["surface_alt"], fg=THEME["muted"], font=(FONT_BASE, 10)).pack(anchor="w")

        self.title_label = ttk.Label(container, text="Setup inicial", style="Header.TLabel")
        self.title_label.pack(anchor="w", pady=(0, 8))

        self.steps_container = ttk.Frame(container)
        self.steps_container.pack(fill="both", expand=True)

        self.steps = [
            self._build_step_general(self.steps_container),
            self._build_step_keywords(self.steps_container),
            self._build_step_apis(self.steps_container),
            self._build_step_summary(self.steps_container),
        ]

        buttons = ttk.Frame(container)
        buttons.pack(fill="x", pady=(10, 0))
        self.back_btn = ttk.Button(buttons, text="Atras", command=self._back, style="Ghost.TButton")
        self.back_btn.pack(side="left")
        self.next_btn = ttk.Button(buttons, text="Siguiente", command=self._next, style="Primary.TButton")
        self.next_btn.pack(side="right")

    def _build_step_general(self, parent):
        frame = ttk.Frame(parent)
        ttk.Label(frame, text="1) Configuracion de asistente", style="Header.TLabel").pack(anchor="w", pady=(0, 8))

        form = ttk.Frame(frame)
        form.pack(fill="x")

        ttk.Label(form, text="Nombre de asistente").grid(row=0, column=0, sticky="w", pady=6)
        ttk.Entry(form, textvariable=self.assistant_name, width=36).grid(row=0, column=1, sticky="we", padx=(8, 0))

        ttk.Label(form, text="Genero de voz").grid(row=1, column=0, sticky="w", pady=6)
        ttk.Combobox(form, textvariable=self.assistant_gender, values=["femenino", "masculino", "neutro"], state="readonly").grid(row=1, column=1, sticky="we", padx=(8, 0))

        ttk.Label(form, text="Ruta modelo Piper (.onnx)").grid(row=2, column=0, sticky="w", pady=6)
        ttk.Entry(form, textvariable=self.voice_model_path).grid(row=2, column=1, sticky="we", padx=(8, 0))
        ttk.Button(form, text="Examinar", command=self._select_voice_model).grid(row=2, column=2, padx=(8, 0))

        ttk.Label(form, text="Velocidad voz (1.0 normal)").grid(row=3, column=0, sticky="w", pady=6)
        ttk.Entry(form, textvariable=self.voice_length_scale, width=12).grid(row=3, column=1, sticky="w", padx=(8, 0))

        ttk.Checkbutton(form, text="Iniciar Luna al iniciar Windows", variable=self.run_on_startup_var).grid(row=4, column=0, columnspan=2, sticky="w", pady=(12, 0))

        form.columnconfigure(1, weight=1)
        return frame

    def _build_step_keywords(self, parent):
        frame = ttk.Frame(parent)
        ttk.Label(frame, text="2) Acciones por palabras clave", style="Header.TLabel").pack(anchor="w", pady=(0, 8))
        ttk.Label(frame, text="Define palabras clave y que accion ejecutar cuando Luna las reconozca.", style="Muted.TLabel").pack(anchor="w", pady=(0, 12))

        list_frame = ttk.Frame(frame)
        list_frame.pack(fill="both", expand=True)

        self.keyword_tree = ttk.Treeview(list_frame, columns=("keywords", "action"), show="headings", height=10)
        self.keyword_tree.heading("keywords", text="Palabras clave")
        self.keyword_tree.heading("action", text="Accion")
        self.keyword_tree.column("keywords", width=380, anchor="w")
        self.keyword_tree.column("action", width=220, anchor="w")
        self.keyword_tree.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.keyword_tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.keyword_tree.configure(yscrollcommand=scrollbar.set)

        actions = ttk.Frame(frame)
        actions.pack(fill="x", pady=(12, 0))
        ttk.Button(actions, text="Agregar", command=self._add_keyword_action, style="Primary.TButton").pack(side="left")
        ttk.Button(actions, text="Editar", command=self._edit_keyword_action, style="Ghost.TButton").pack(side="left", padx=6)
        ttk.Button(actions, text="Eliminar", command=self._delete_keyword_action, style="Danger.TButton").pack(side="left")

        self._render_keyword_actions()
        return frame

    def _build_step_apis(self, parent):
        frame = ttk.Frame(parent)
        ttk.Label(frame, text="3) Credenciales de APIs (opcional)", style="Header.TLabel").pack(anchor="w", pady=(0, 8))
        ttk.Label(frame, text="La validacion de ciudad es opcional. Puedes guardar sin validar.").pack(anchor="w", pady=(0, 12))

        form = ttk.Frame(frame)
        form.pack(fill="x")
        ttk.Label(form, text="OpenWeather API Key").grid(row=0, column=0, sticky="w", pady=6)
        ttk.Entry(form, textvariable=self.openweather_key).grid(row=0, column=1, sticky="we", padx=(8, 0))

        ttk.Label(form, text="Ciudad Designada").grid(row=1, column=0, sticky="w", pady=6)
        ttk.Combobox(form, textvariable=self.designated_city, values=COMMON_OPENWEATHER_CITIES).grid(row=1, column=1, sticky="we", padx=(8, 0))
        ttk.Button(form, text="Validar ciudad", command=self._validate_city).grid(row=1, column=2, padx=(8, 0))

        ttk.Label(form, text="Spotify Client ID").grid(row=2, column=0, sticky="w", pady=6)
        ttk.Entry(form, textvariable=self.spotify_client_id).grid(row=2, column=1, sticky="we", padx=(8, 0))

        ttk.Label(form, text="Spotify Client Secret").grid(row=3, column=0, sticky="w", pady=6)
        ttk.Entry(form, textvariable=self.spotify_client_secret, show="*").grid(row=3, column=1, sticky="we", padx=(8, 0))

        ttk.Label(form, text="Spotify Redirect URI").grid(row=4, column=0, sticky="w", pady=6)
        ttk.Entry(form, textvariable=self.spotify_redirect).grid(row=4, column=1, sticky="we", padx=(8, 0))
        form.columnconfigure(1, weight=1)
        return frame

    def _build_step_summary(self, parent):
        frame = ttk.Frame(parent)
        ttk.Label(frame, text="4) Finalizar", style="Header.TLabel").pack(anchor="w", pady=(0, 8))
        ttk.Label(frame, text="Cuando cierres la ventana, Luna sigue en background.").pack(anchor="w", pady=(0, 8))
        ttk.Label(frame, text="Atajo para reabrir ajustes: Ctrl + Shift + |").pack(anchor="w")
        return frame

    def _render_keyword_actions(self):
        for item in self.keyword_tree.get_children():
            self.keyword_tree.delete(item)
        for entry in self.keyword_actions:
            keywords = ", ".join(entry.get("keywords", []))
            action_label = format_keyword_action(entry)
            self.keyword_tree.insert("", "end", iid=entry.get("id"), values=(keywords, action_label))

    def _add_keyword_action(self):
        dialog = KeywordActionDialog(self.window)
        self.window.wait_window(dialog.window)
        if dialog.result:
            self.keyword_actions.append(dialog.result)
            self._render_keyword_actions()

    def _edit_keyword_action(self):
        selected = self.keyword_tree.selection()
        if not selected:
            messagebox.showinfo("Accion", "Selecciona una accion para editar.")
            return
        action_id = selected[0]
        existing = next((entry for entry in self.keyword_actions if entry.get("id") == action_id), None)
        if not existing:
            return
        dialog = KeywordActionDialog(self.window, existing=existing)
        self.window.wait_window(dialog.window)
        if dialog.result:
            dialog.result["id"] = action_id
            self.keyword_actions = [
                dialog.result if entry.get("id") == action_id else entry for entry in self.keyword_actions
            ]
            self._render_keyword_actions()

    def _delete_keyword_action(self):
        selected = self.keyword_tree.selection()
        if not selected:
            messagebox.showinfo("Accion", "Selecciona una accion para eliminar.")
            return
        action_id = selected[0]
        self.keyword_actions = [entry for entry in self.keyword_actions if entry.get("id") != action_id]
        self._render_keyword_actions()

    def _show_step(self, index):
        self.current_step = max(0, min(index, len(self.steps) - 1))
        for step in self.steps:
            step.pack_forget()
        self.steps[self.current_step].pack(fill="both", expand=True)
        self.back_btn.configure(state="normal" if self.current_step > 0 else "disabled")
        self.next_btn.configure(text="Finalizar" if self.current_step == len(self.steps) - 1 else "Siguiente")

    def _next(self):
        if self.current_step < len(self.steps) - 1:
            self._show_step(self.current_step + 1)
            return
        if self._persist():
            self.window.destroy()

    def _back(self):
        self._show_step(self.current_step - 1)

    def _select_voice_model(self):
        chosen = filedialog.askopenfilename(title="Seleccionar modelo Piper", filetypes=[("ONNX model", "*.onnx"), ("All", "*.*")])
        if chosen:
            self.voice_model_path.set(chosen)

    def _validate_city(self):
        valid, message = validate_openweather_city(self.designated_city.get(), self.openweather_key.get())
        if valid:
            self.designated_city.set(message)
            messagebox.showinfo("Ciudad", f"Ciudad valida: {message}")
            return True
        messagebox.showerror("Ciudad", message)
        return False

    def _persist(self):
        global APP_CONFIG, OPENWEATHER_API_KEY, voice, voice_loaded
        city = self.designated_city.get().strip()
        with APP_CONFIG_LOCK:
            APP_CONFIG["setup_complete"] = True
            APP_CONFIG["assistant"]["name"] = self.assistant_name.get().strip() or "Luna"
            APP_CONFIG["assistant"]["gender"] = self.assistant_gender.get().strip() or "femenino"
            APP_CONFIG["voice"]["model_path"] = self.voice_model_path.get().strip()
            try:
                APP_CONFIG["voice"]["length_scale"] = float(self.voice_length_scale.get().strip())
            except ValueError:
                APP_CONFIG["voice"]["length_scale"] = 1.2
            APP_CONFIG["keyword_actions"] = self.keyword_actions
            APP_CONFIG["ui"]["run_on_startup"] = bool(self.run_on_startup_var.get())
            APP_CONFIG["apis"]["openweather_api_key"] = self.openweather_key.get().strip()
            APP_CONFIG["apis"]["spotify_client_id"] = self.spotify_client_id.get().strip()
            APP_CONFIG["apis"]["spotify_client_secret"] = self.spotify_client_secret.get().strip()
            APP_CONFIG["apis"]["spotify_redirect_uri"] = self.spotify_redirect.get().strip() or "http://127.0.0.1:8888/callback"
            APP_CONFIG.setdefault("weather", {})
            APP_CONFIG["weather"]["city"] = city
            save_app_config(APP_CONFIG)
        write_city(city)
        set_run_on_startup(bool(self.run_on_startup_var.get()))
        OPENWEATHER_API_KEY = load_openweather_api_key()
        voice = load_voice()
        voice_loaded = True
        return True

def launch_setup_if_needed():
    if get_config_value(("setup_complete",), False):
        return
    wizard = SetupWizard(None)
    wizard.window.mainloop()

class ControlPanel:
    def __init__(self):
        self.root = Tk()
        self.root.title("Luna Assistant")
        self.content = create_modern_window(
            self.root,
            "Luna Assistant",
            (860, 580),
            (820, 540),
            on_close=self._on_close,
            on_minimize=self._on_minimize,
            resizable=True,
        )
        self.tray_icon = None
        self._build_ui()
        keyboard.add_hotkey("ctrl+shift+|", self._show_from_hotkey)
        self.root.bind_all("<Button-1>", self._on_global_click, add="+")

    def _build_ui(self):
        frame = ttk.Frame(self.content, padding=8)
        frame.pack(fill="both", expand=True)
        name = get_config_value(("assistant", "name"), "Luna")

        hero = tk.Frame(frame, bg=THEME["surface_alt"], highlightthickness=1, highlightbackground=THEME["border"])
        hero.pack(fill="x", pady=(0, 12))
        accent = tk.Frame(hero, bg=THEME["accent_2"], width=8, height=80)
        accent.pack(side="left", fill="y")
        hero_content = tk.Frame(hero, bg=THEME["surface_alt"])
        hero_content.pack(side="left", fill="x", expand=True, padx=12, pady=12)
        tk.Label(hero_content, text=f"{name}", bg=THEME["surface_alt"], fg=THEME["text"], font=(FONT_BOLD, 18)).pack(anchor="w")
        tk.Label(hero_content, text="Panel de control", bg=THEME["surface_alt"], fg=THEME["muted"], font=(FONT_BASE, 10)).pack(anchor="w")
        tk.Label(hero_content, text="Atajo de voz: Ctrl + |    |    Reabrir panel: Ctrl + Shift + |", bg=THEME["surface_alt"], fg=THEME["muted"], font=(FONT_BASE, 9)).pack(anchor="w", pady=(6, 0))

        actions = ttk.Frame(frame)
        actions.pack(fill="x", pady=12)
        ttk.Button(actions, text="Editar configuracion", command=self.open_settings, style="Primary.TButton").pack(side="left")
        ttk.Button(actions, text="Ver comandos", command=self.show_commands, style="Ghost.TButton").pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Salir", command=self.exit_app, style="Danger.TButton").pack(side="right")

        self.info = ttk.Label(frame, text="", justify="left")
        self.info.pack(anchor="w", pady=(12, 0))
        self.refresh_info()

    def refresh_info(self):
        keyword_count = len(get_config_value(("keyword_actions",), []) or [])
        spotify_ready = bool(get_spotify_client())
        self.info.configure(
            text=(
                f"Setup completo: {'si' if get_config_value(('setup_complete',), False) else 'no'}\n"
                f"Acciones por palabras clave: {keyword_count}\n"
                f"Spotify API configurada: {'si' if spotify_ready else 'no'}\n"
                f"Ciudad designada: {read_city() or 'no configurada'}"
            )
        )

    def open_settings(self):
        wizard = SetupWizard(self.root)
        self.root.wait_window(wizard.window)
        self.refresh_info()

    def show_commands(self):
        window = Toplevel(self.root)
        window.title("Comandos de Luna")
        content = create_modern_window(
            window,
            "Comandos Disponibles",
            (760, 560),
            (680, 500),
            on_close=window.destroy,
            on_minimize=window.withdraw,
            resizable=True,
        )
        frame = ttk.Frame(content, padding=8)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Comandos nativos", style="Header.TLabel").pack(anchor="w", pady=(0, 8))
        text = tk.Text(
            frame,
            bg=THEME["surface_alt"],
            fg=THEME["text"],
            insertbackground=THEME["text"],
            wrap="word",
            height=16,
            relief="flat",
        )
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=scrollbar.set)
        text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        builtin_commands = [
            ("Busca en Google", "Abre una busqueda en Google con la consulta dictada."),
            ("Busca en YouTube", "Abre resultados de YouTube con la consulta dictada."),
            ("Busca en Wikipedia", "Busca un tema y abre el primer resultado en Wikipedia."),
            ("Spotify: pon/reproduce", "Reproduce una cancion en Spotify."),
            ("Spotify: pausa/reanuda/siguiente/anterior", "Controla la reproduccion de Spotify."),
            ("Spotify: volumen <0-100>", "Cambia el volumen del dispositivo activo."),
            ("Escanea puertos", "Pide IP o dominio por popup y escanea puertos 1-99."),
            ("Prueba de conexion", "Pide URL por popup y verifica acceso web."),
            ("Clima/Tiempo", "Consulta clima de la Ciudad Designada con OpenWeather."),
            ("YouTube/Google/Instagram/Twitter/Twitch/Netflix", "Abre el sitio correspondiente."),
            ("Direccion IP", "Dice la IP local y la copia al portapapeles."),
            ("Hora/Numero es hoy", "Indica hora actual o dia del mes."),
            ("Bateria/CPU", "Informa bateria o uso actual de CPU."),
            ("Cambia de pantalla", "Ejecuta Alt+Tab."),
            ("Captura", "Abre herramienta de recorte de Windows."),
            ("A con tilde, E con tilde, etc.", "Copia letras acentuadas al portapapeles."),
            ("Apaga/Reinicia equipo", "Ejecuta apagado o reinicio del sistema."),
            ("Apaga sistemas", "Cierra Luna."),
        ]

        for command_name, description in builtin_commands:
            text.insert("end", f"- {command_name}: {description}\n")

        text.insert("end", "\nComandos adicionales configurados por el usuario\n")
        user_actions = get_config_value(("keyword_actions",), []) or []
        if not user_actions:
            text.insert("end", "- No hay comandos adicionales configurados.\n")
        else:
            for entry in user_actions:
                keywords = ", ".join(entry.get("keywords", [])) or "(sin palabras clave)"
                action_label = format_keyword_action(entry)
                text.insert("end", f"- {keywords}: {action_label}\n")
        text.configure(state="disabled")

    def _show_from_hotkey(self):
        self.root.after(0, self.show)

    def show(self):
        self.root.deiconify()
        self.root.attributes("-topmost", True)
        self.root.lift()
        self.root.focus_force()
        self.root.after(200, lambda: self.root.attributes("-topmost", False))

    def _on_close(self):
        if get_config_value(("ui", "run_in_background_on_close"), True):
            self.root.withdraw()
            notify("Luna", "La app sigue corriendo en segundo plano.")
            self._start_tray_icon()
            return
        self.exit_app()

    def _on_minimize(self):
        self.root.withdraw()

    def _on_global_click(self, event):
        try:
            if not self.root.winfo_viewable():
                return
            widget = event.widget
            if widget and widget.winfo_toplevel() is self.root:
                return
            self._on_minimize()
        except Exception:
            self._on_minimize()

    def _start_tray_icon(self):
        if pystray is None or Image is None or self.tray_icon is not None:
            return
        image = Image.new("RGB", (64, 64), "#2B2B2B")
        draw = ImageDraw.Draw(image)
        draw.ellipse((14, 14, 50, 50), fill="#FFC66D")
        draw.ellipse((24, 24, 40, 40), fill="#2B2B2B")

        menu = pystray.Menu(
            pystray.MenuItem("Mostrar", lambda icon, item: self.show()),
            pystray.MenuItem("Salir", lambda icon, item: self.exit_app()),
        )
        self.tray_icon = pystray.Icon("Luna", image, "Luna", menu)
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def exit_app(self):
        global APP_RUNNING
        APP_RUNNING = False
        try:
            if self.tray_icon is not None:
                self.tray_icon.stop()
        except Exception:
            pass
        self.root.quit()
        self.root.destroy()

voice = None
voice_loaded = False
voice_lock = threading.Lock()
voiceFile = os.path.join(tempfile.gettempdir(), "luna_output.wav")
fallback_engine = None

def load_voice():
    model_path = get_config_value(("voice", "model_path"), DEFAULT_APP_CONFIG["voice"]["model_path"])
    if not os.path.isfile(model_path):
        return None
    try:
        return PiperVoice.load(model_path)
    except Exception:
        return None

USE_PIPER_TTS = True

def get_voice():
    global voice, voice_loaded
    if voice_loaded:
        return voice
    with voice_lock:
        if not voice_loaded:
            voice = load_voice()
            voice_loaded = True
    return voice

def get_fallback_engine():
    global fallback_engine
    if fallback_engine is not None:
        return fallback_engine
    fallback_engine = pyttsx3.init()
    fallback_engine.setProperty("rate", 155)
    return fallback_engine

##Funciones

def play(target_filename):
    if not target_filename or not os.path.isfile(target_filename):
        return False
    try:
        # Play the wav using sounddevice
        with wave.open(target_filename, 'rb') as wf:
            sample_rate = wf.getframerate()
            n_channels = wf.getnchannels()
            n_frames = wf.getnframes()
            audio = wf.readframes(n_frames)
            audio_np = numpy.frombuffer(audio, dtype=numpy.int16)
            if n_channels == 2:
                audio_np = audio_np.reshape(-1, 2)
            audio_np = audio_np.astype(numpy.float32) / 32768.0
            sd.play(audio_np, sample_rate)
            sd.wait()
        return True
    except Exception:
        return False

def speak(text):
    if not text:
        return
    current_voice = get_voice() if USE_PIPER_TTS else None
    if current_voice is not None:
        syn_config = SynthesisConfig(
            volume=float(get_config_value(("voice", "volume"), 1.0)),
            length_scale=float(get_config_value(("voice", "length_scale"), 1.2)),
            normalize_audio=False,
        )
        with wave.open(voiceFile, mode="wb") as file:
            current_voice.synthesize_wav(text=text, wav_file=file, syn_config=syn_config)
        play(voiceFile)
        return

    engine = get_fallback_engine()
    engine.say(text)
    engine.runAndWait()

def takeCommand():
    if KaldiRecognizer is None:
        return "none"
    model = get_vosk_model()
    if not model:
        return "none"

    recognizer = KaldiRecognizer(model, 16000)
    audio = pyaudio.PyAudio()
    stream = audio.open(
        format=pyaudio.paInt16,
        channels=1,
        rate=16000,
        input=True,
        frames_per_buffer=8000,
    )
    stream.start_stream()

    result_text = ""
    start_time = time.time()

    try:
        while time.time() - start_time < 7:
            data = stream.read(4000, exception_on_overflow=False)
            if recognizer.AcceptWaveform(data):
                result = json.loads(recognizer.Result())
                result_text = result.get("text", "")
                if result_text:
                    break
    except Exception:
        result_text = ""
    finally:
        stream.stop_stream()
        stream.close()
        audio.terminate()

    if not result_text:
        final = json.loads(recognizer.FinalResult())
        result_text = final.get("text", "")

    return result_text if result_text else "none"

def clipBoard():
    try:
        win32clipboard.OpenClipboard()
        data = win32clipboard.GetClipboardData()
        win32clipboard.CloseClipboard()
        return data
    except TypeError:
        return "None"

def scannerPuertos(target):
    target = str(target).strip()
    if not target:
        return "error1"
    puertos = []
    try:

        # will scan ports between 1 to 65,535
        for port in range(1, 100):
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            socket.setdefaulttimeout(0.0000001)

            if port == 10001:
                quit()

            # returns an error indicator
            result = s.connect_ex((target, port))
            if result == 0:
                puertos.append(str(port))
            s.close()

        return puertos

    except socket.gaierror:
        speak("Error, la direcciÃ³n aipi no es vÃ¡lida")
        return "error1"

    except socket.error:
        speak("Error, el servidor no responde")
        return "error2"

def notify(title, msg):
    notification = Notification(
        title=title,
        msg=msg,
        icon=asset_path("Imagenes", "LunaLogo.png"),
        app_id="Luna",
    )
    notification.show()

def normalize_statement(statement):
    replacements = (
        ("niÃƒÂ±a", "luna"),
        ("niÃƒÂ±o", "luna"),
        ("lyna", "luna"),
        ("una", "luna"),
    )
    for old, new in replacements:
        if old in statement and "luna" not in statement:
            statement = statement.replace(old, new)
    return statement

def normalize_intent_text(text):
    text = unicodedata.normalize("NFD", text.lower())
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return " ".join(text.split())

def read_city():
    city = get_config_value(("weather", "city"), "").strip()
    if city:
        return city
    if not os.path.isfile(CITY_FILE):
        return ""
    with open(CITY_FILE, "r", encoding="utf-8") as file:
        return file.read().strip()

def write_city(city):
    with APP_CONFIG_LOCK:
        APP_CONFIG.setdefault("weather", {})
        APP_CONFIG["weather"]["city"] = city.strip()
        save_app_config(APP_CONFIG)
    os.makedirs(STATE_DIR, exist_ok=True)
    with open(CITY_FILE, "w", encoding="utf-8") as file:
        file.write(city.strip())

def validate_openweather_city(city, api_key):
    clean_city = city.strip()
    if not clean_city:
        return False, "Ingresa una ciudad."
    if not api_key.strip():
        return False, "Necesitas una OpenWeather API Key para validar la ciudad."
    try:
        response = requests.get(
            "https://api.openweathermap.org/data/2.5/weather",
            params={"q": clean_city, "appid": api_key.strip()},
            timeout=8,
        )
        data = response.json()
    except Exception:
        return False, "No se pudo validar la ciudad por un error de red."
    if response.status_code == 200 and data.get("main"):
        resolved_name = data.get("name", clean_city)
        country = data.get("sys", {}).get("country", "")
        if country:
            resolved_name = f"{resolved_name},{country}"
        return True, resolved_name
    if response.status_code == 404:
        return False, "OpenWeather no encontro esa ciudad. Usa Ciudad,CC. Ej: Buenos Aires,AR"
    return False, "OpenWeather rechazo la validacion. Revisa API key y ciudad."

def prompt_user_text(title, prompt, initial_value=""):
    result = {"value": None}
    done = threading.Event()

    def ask_with_parent():
        try:
            parent = CONTROL_PANEL.root if CONTROL_PANEL is not None else None
            result["value"] = simpledialog.askstring(
                title,
                prompt,
                parent=parent,
                initialvalue=initial_value,
            )
        finally:
            done.set()

    if CONTROL_PANEL is not None and hasattr(CONTROL_PANEL, "root"):
        try:
            if threading.current_thread() is threading.main_thread():
                ask_with_parent()
            else:
                CONTROL_PANEL.root.after(0, ask_with_parent)
                done.wait()
            return result["value"]
        except Exception:
            pass

    temp_root = Tk()
    temp_root.withdraw()
    try:
        return simpledialog.askstring(title, prompt, parent=temp_root, initialvalue=initial_value)
    finally:
        temp_root.destroy()

##Main

def assistant_loop():
    global APP_RUNNING
    name = get_config_value(("assistant", "name"), "Luna")
    speak(f"Hola, soy {name}, que puedo hacer por ti?")

    while APP_RUNNING:

        keyboard.wait('ctrl+|')
        play(os.path.join(SOUNDS_DIR, "IniciarSound.wav"))
        statement = takeCommand().lower()

        ##Ajuste de comandos

        statement = normalize_statement(statement)
        intent_statement = normalize_intent_text(statement)

        if statement == "none":
            notify("Error de reconocimiento", "No se ha podido reconocer ninguna frase")
            speak("No te entendi, puedes repetirlo?")
            continue

        notify("Texto reconocido", statement)

        if handle_keyword_actions(intent_statement):
            time.sleep(1)
            continue

        #Comandos de busqueda



        if 'busca' in statement and "google" in statement and not "wikipedia" in statement and not "youtube" in statement:
            statement = statement.replace("busca", "")
            statement = statement.replace("google", "")
            statement = statement.replace("en", "")
            statement = statement.replace("luna", "")
            webbrowser.open_new_tab("https://www.google.com/search?q=" + statement)
            speak('Buscando en gugel ' + statement)
            time.sleep(3)
            continue

        if 'busca' in statement and "youtube" in statement:
            statement = statement.replace("busca", "")
            statement = statement.replace("youtube", "")
            statement = statement.replace("en", "")
            statement = statement.replace("luna", "")
            webbrowser.open_new_tab("https://www.youtube.com/results?search_query=" + statement)
            speak('Buscando en youtube ' + statement)
            time.sleep(3)
            continue

        spotify_context = (
            "spotify" in intent_statement
            or "musica" in intent_statement
            or "cancion" in intent_statement
        )
        if spotify_context:
            if any(token in intent_statement for token in ("pausa", "pausar", "frena", "deten")):
                spotify_pause()
                time.sleep(1)
                continue

            if any(token in intent_statement for token in ("reanuda", "continuar", "resume", "seguir")):
                spotify_resume()
                time.sleep(1)
                continue

            if "siguiente" in intent_statement:
                spotify_next()
                time.sleep(1)
                continue

            if any(token in intent_statement for token in ("anterior", "previa", "atras", "atraz")):
                spotify_prev()
                time.sleep(1)
                continue

            if "volumen" in intent_statement:
                match = re.search(r"\b(\d{1,3})\b", intent_statement)
                if not match:
                    speak("Decime un volumen entre 0 y 100.")
                else:
                    percent = max(0, min(100, int(match.group(1))))
                    spotify_set_volume(percent)
                time.sleep(1)
                continue

            if any(token in intent_statement for token in ("pon", "pone", "reproduce", "reproducir")):
                query = intent_statement
                for token in (
                    "spotify", "musica", "cancion", "pon", "pone", "reproduce", "reproducir",
                    "en", "la", "el", "una", "un", "luna",
                ):
                    query = re.sub(rf"\b{token}\b", " ", query)
                query = " ".join(query.split())
                spotify_play_query(query)
                time.sleep(1)
                continue

        if 'busca' in statement and "wikipedia" in statement or "quÃ© es" in statement or "que es" in statement or "que son" in statement or "quÃ© son" in statement:
            statement = statement.replace("busca", "")
            statement = statement.replace("wikipedia", "")
            statement = statement.replace("en", "")
            statement = statement.replace("luna", "")
            statement = statement.replace("lo", "")
            if "que" in statement or "quÃ©" in statement:
                statement = statement.replace("que", "")
                statement = statement.replace("quÃ©", "")
                statement = statement.replace("es", "")
                statement = statement.replace("los", "")

            speak("buscando en wikipedia " + statement + ", por favor espera un momento")

            try:

                results = wikipedia.search(statement)
                page = wikipedia.page(results[0])
                webbrowser.open_new_tab(page.url)
                speak(page.summary)

            except:

                speak("No se han encontrado resultados en wikipedia")

            time.sleep(3)
            continue


        ##Comandos sociales

        assistant_name = get_config_value(("assistant", "name"), "luna").lower()
        if intent_statement == assistant_name:
            speak("En que puedo ayudarte?")
            time.sleep(3)
            continue
        
        #Comandos de apagado

        elif "apaga" in statement and "equipo" in statement:
            speak("Apagando el equipo")
            os.system("shutdown /s /t 1")

        elif "reinicia" in statement and "equipo" in statement:
            speak("Reiniciando el equipo")
            os.system("shutdown /r /t 1")

        elif "apaga" in statement or "apÃ¡galo" in statement and "sistema" in statement or "sistemas" in statement:
            speak("Apagando sistemas")
            APP_RUNNING = False
            break

        ##Comandos hacking

        elif "escanea" in statement and "puerto" in statement:
            target = prompt_user_text("Escaneo de puertos", "Ingresa IP o dominio a escanear:")
            if not target:
                speak("Escaneo cancelado")
                continue
            speak("Ejecutando un escaneo de puertos en la aipi seleccionada")
            resultado = scannerPuertos(target)

            if resultado not in ("error1", "error2"):

                speak("Los puertos")
                for puertoAbierto in resultado:
                     speak(puertoAbierto)
                speak("estan abiertos")


        elif "prueba" in statement and ("conexiÃ³n" in statement or "conexion" in statement):
            link = prompt_user_text("Prueba de conexion", "Ingresa URL para probar conexion:", "https://")
            if not link:
                speak("Prueba cancelada")
                continue
            try:
                urllib.request.urlopen(link)
                speak("conexiÃ³n exitosa")

            except:
                speak("no ha sido posible conectarse con la wueb seleccionada")

        ##Comandos generales


        elif 'youtube' in statement:
            speak("abriendo youtube")
            webbrowser.open_new_tab("https://youtube.com")
            time.sleep(3)
            continue

        elif 'google' in statement and not "earth" in statement and not "busca" in statement:
            speak("abriendo gugel")
            webbrowser.open_new_tab("https://www.google.com")
            time.sleep(3)
            continue

        elif "instagram" in statement:
            speak("iniciando instagram")
            webbrowser.open_new_tab("https://www.instagram.com")

        elif 'direcciÃ³n ip' in statement or "direcciÃ³n id" in statement:
            speak("Tu direcciÃ³n aipi es " + socket.gethostbyname(
                socket.gethostname()) + ". lo he copiado en tu portapapeles")
            pyperclip.copy(socket.gethostbyname(socket.gethostname()))
            time.sleep(3)
            continue

        elif 'twitter' in statement:
            speak("abriendo tuiter")
            webbrowser.open_new_tab("https://www.twitter.com")
            time.sleep(3)
            continue

        elif 'twitch' in statement:
            speak("abriendo tuich")
            webbrowser.open_new_tab("https://www.twitch.com")
            time.sleep(3)
            continue

        elif 'nÃºmero es hoy' in statement:
            speak("hoy es " + str(datetime.now().strftime("%d")))
            time.sleep(3)
            continue

        elif 'hora' in statement:
            speak("ahora son las " + datetime.now().strftime("%I") + " y " + datetime.now().strftime("%M"))
            time.sleep(3)
            continue

        elif 'baterÃ­a' in statement:
            speak("Queda un total de " + str(psutil.sensors_battery().percent) + " porciento de baterÃ­a")
            time.sleep(3)
            continue

        elif "cpu" in statement or "sepe" in statement or "procesador" in statement:
            speak("El CPU esta al " + str(psutil.cpu_percent()) + " porciento de su capacidad mÃ¡xima")
            time.sleep(3)
            continue

        elif "cambia de pantalla" in statement:
            speak("cambiando pantalla")
            pyautogui.hotkey('alt', 'tab')
            time.sleep(3)
            continue

        elif "captura" in statement:
            speak("haciendo una captura de pantalla")
            pyautogui.hotkey('winleft', 'shiftleft', "s")
            time.sleep(3)
            continue


        elif "netflix" in statement:
            speak("iniciando netflix")
            webbrowser.open_new_tab("www.netflix.com")
            time.sleep(3)
            continue

        elif "e con tilde" in statement:
            pyperclip.copy("Ã©")
            speak("he copiado la e con tilde en tu portapapeles")
            time.sleep(3)
            continue

        elif "u con tilde" in statement:
            pyperclip.copy("Ãº")
            speak("he copiado la u con tilde en tu portapapeles")
            time.sleep(3)
            continue

        elif "o con tilde" in statement:
            pyperclip.copy("Ã³")
            speak("he copiado la o con tilde en tu portapapeles")
            time.sleep(3)
            continue

        elif "i con tilde" in statement or "y con tilde" in statement:
            pyperclip.copy("Ã­")
            speak("he copiado la i con tilde en tu portapapeles")
            time.sleep(3)
            continue

        elif "a con tilde" in statement:
            pyperclip.copy("Ã¡")
            speak("he copiado la a con tilde en tu portapapeles")
            time.sleep(3)
            continue

        elif "apaga" in statement and "volumen" in statement:
            pyautogui.press("volumemute")
            continue

        elif "prende" in statement and "volumen" in statement:
            pyautogui.press("volumeup")
            continue

        elif 'clima' in statement or "tiempo" in statement:

            city = read_city()

            if not OPENWEATHER_API_KEY:
                speak("No tengo una llave de OpenWeather configurada")
                time.sleep(3)
                continue

            if not city:
                speak("No hay una ciudad configurada")
                time.sleep(3)
                continue

            try:
                response = requests.get(
                    "https://api.openweathermap.org/data/2.5/weather",
                    params={"q": city, "appid": OPENWEATHER_API_KEY},
                    timeout=8,
                )
                weather_data = response.json()
                if response.status_code != 200 or "main" not in weather_data:
                    speak("Error. La ciudad ingresada no es valida para OpenWeather")
                    time.sleep(3)
                    continue
                temperature = round(weather_data["main"]["temp"] - 273.15, 2)
                humidity = weather_data["main"]["humidity"]

                speak(f"La temperatura es de {temperature} grados y la humedad es de {humidity} por ciento")

            except:
                speak("Error. La ciudad ingresada no es valida")

            time.sleep(3)
            continue

if __name__ == '__main__':
    ensure_runtime_structure()
    set_run_on_startup(bool(get_config_value(("ui", "run_on_startup"), False)))
    launch_setup_if_needed()
    panel = ControlPanel()
    CONTROL_PANEL = panel
    run_startup_keyword_actions()
    assistant_thread = threading.Thread(target=assistant_loop, daemon=True)
    assistant_thread.start()
    panel.root.after(50, panel.show)
    panel.root.mainloop()
