"""
Speech-to-Cursor: Windows spraakherkenning die gesproken tekst typt waar je cursor staat.

Gebruik:
  - Ctrl+Spatie ingedrukt houden : luisteren (push-to-talk)
  - Ctrl+Shift+Q                 : afsluiten
  - Overlay-venster: sluitknop (X) of versleep naar andere positie
"""

import ctypes
import ctypes.wintypes
import json
import os
import queue
import sys
import threading
import tkinter as tk
import tkinter.messagebox as messagebox
import urllib.request
import winsound
import zipfile

import pystray
import sounddevice as sd
from PIL import Image, ImageDraw
from vosk import KaldiRecognizer, Model

# --- Configuration ---
APP_NAME = "Speech-to-Cursor"
APP_DATA_DIR = os.path.join(os.environ.get("LOCALAPPDATA", "."), APP_NAME)
MODEL_PATH = os.path.join(APP_DATA_DIR, "model")
MODEL_URL = "https://alphacephei.com/vosk/models/vosk-model-nl-spraakherkenning-0.6.zip"
SAMPLE_RATE = 16000

# Virtual key codes
VK_CONTROL = 0x11
VK_SPACE = 0x20
VK_SHIFT = 0x10
VK_Q = 0x51

# --- Global state ---
audio_queue = queue.Queue()
listening = False
recognizer = None
model = None
tray_icon = None
overlay = None

# --- Low-level keyboard hook via Windows API ---
# GetAsyncKeyState doesn't reliably detect all keys from background processes.
# A WH_KEYBOARD_LL hook with its own message pump is the correct approach.
WH_KEYBOARD_LL = 13
WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101
WM_SYSKEYDOWN = 0x0104
WM_SYSKEYUP = 0x0105

HOOKPROC = ctypes.CFUNCTYPE(
    ctypes.c_long,
    ctypes.c_int,
    ctypes.wintypes.WPARAM,
    ctypes.wintypes.LPARAM,
)

class KBDLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [
        ("vkCode", ctypes.wintypes.DWORD),
        ("scanCode", ctypes.wintypes.DWORD),
        ("flags", ctypes.wintypes.DWORD),
        ("time", ctypes.wintypes.DWORD),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
    ]

# Properly declare Win32 function signatures for 64-bit compatibility
_user32 = ctypes.WinDLL("user32", use_last_error=True)
_kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

_user32.SetWindowsHookExW.argtypes = [
    ctypes.c_int, HOOKPROC, ctypes.wintypes.HMODULE, ctypes.wintypes.DWORD
]
_user32.SetWindowsHookExW.restype = ctypes.wintypes.HHOOK

_user32.CallNextHookEx.argtypes = [
    ctypes.wintypes.HHOOK, ctypes.c_int, ctypes.wintypes.WPARAM, ctypes.wintypes.LPARAM
]
_user32.CallNextHookEx.restype = ctypes.c_long

_user32.UnhookWindowsHookEx.argtypes = [ctypes.wintypes.HHOOK]
_user32.UnhookWindowsHookEx.restype = ctypes.wintypes.BOOL

_user32.GetMessageW.argtypes = [
    ctypes.POINTER(ctypes.wintypes.MSG), ctypes.wintypes.HWND,
    ctypes.c_uint, ctypes.c_uint
]
_user32.GetMessageW.restype = ctypes.wintypes.BOOL

_user32.TranslateMessage.argtypes = [ctypes.POINTER(ctypes.wintypes.MSG)]
_user32.DispatchMessageW.argtypes = [ctypes.POINTER(ctypes.wintypes.MSG)]

_user32.PostThreadMessageW.argtypes = [
    ctypes.wintypes.DWORD, ctypes.c_uint, ctypes.wintypes.WPARAM, ctypes.wintypes.LPARAM
]

_user32.GetAsyncKeyState.argtypes = [ctypes.c_int]
_user32.GetAsyncKeyState.restype = ctypes.c_short

_kernel32.GetModuleHandleW.argtypes = [ctypes.wintypes.LPCWSTR]
_kernel32.GetModuleHandleW.restype = ctypes.wintypes.HMODULE

_kernel32.GetCurrentThreadId.restype = ctypes.wintypes.DWORD

# Clipboard-related functions (64-bit safe)
_kernel32.GlobalAlloc.argtypes = [ctypes.c_uint, ctypes.c_size_t]
_kernel32.GlobalAlloc.restype = ctypes.c_void_p
_kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
_kernel32.GlobalLock.restype = ctypes.c_void_p
_kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
_kernel32.GlobalUnlock.restype = ctypes.wintypes.BOOL
_user32.OpenClipboard.argtypes = [ctypes.wintypes.HWND]
_user32.EmptyClipboard.argtypes = []
_user32.SetClipboardData.argtypes = [ctypes.c_uint, ctypes.c_void_p]
_user32.SetClipboardData.restype = ctypes.c_void_p
_user32.CloseClipboard.argtypes = []

_ctrl_held = False
_space_held = False
_ptt_active = False
_ptt_event = threading.Event()  # signals the worker thread when state changes
_hook_handle = None
_hook_proc = None
_hook_thread_id = None
_quit_requested = False


def _ll_keyboard_proc(nCode, wParam, lParam):
    """Low-level keyboard hook — MUST return fast or Windows kills the hook."""
    global _ctrl_held, _space_held, _ptt_active, _quit_requested
    if nCode >= 0:
        kb = ctypes.cast(lParam, ctypes.POINTER(KBDLLHOOKSTRUCT)).contents
        vk = kb.vkCode
        is_down = wParam in (WM_KEYDOWN, WM_SYSKEYDOWN)

        if vk == VK_CONTROL or vk == 0xA2 or vk == 0xA3:
            _ctrl_held = is_down
        elif vk == VK_SPACE:
            _space_held = is_down
        elif is_down and vk == VK_Q and _ctrl_held:
            if _user32.GetAsyncKeyState(VK_SHIFT) & 0x8000:
                _quit_requested = True
                _ptt_event.set()

        # Update active state and wake worker thread — no heavy work here
        new_active = _ctrl_held and _space_held
        if new_active != _ptt_active:
            _ptt_active = new_active
            _ptt_event.set()

    return _user32.CallNextHookEx(None, nCode, wParam, lParam)


def _ptt_worker():
    """Worker thread that reacts to push-to-talk state changes."""
    was_active = False
    while True:
        _ptt_event.wait()
        _ptt_event.clear()

        if _quit_requested:
            quit_app()
            return

        active = _ptt_active
        if active and not was_active:
            start_listening()
        elif not active and was_active:
            stop_listening()
        was_active = active


def _start_keyboard_hook():
    """Install a system-wide low-level keyboard hook with its own message pump."""
    global _hook_proc, _hook_handle, _hook_thread_id

    # Start worker thread that handles start/stop listening
    worker = threading.Thread(target=_ptt_worker, daemon=True)
    worker.start()

    def _hook_thread():
        global _hook_proc, _hook_handle, _hook_thread_id
        _hook_thread_id = _kernel32.GetCurrentThreadId()
        _hook_proc = HOOKPROC(_ll_keyboard_proc)
        h_mod = _kernel32.GetModuleHandleW("user32.dll")
        _hook_handle = _user32.SetWindowsHookExW(
            WH_KEYBOARD_LL, _hook_proc, h_mod, 0
        )
        if not _hook_handle:
            err = ctypes.get_last_error()
            print(f"[hook] SetWindowsHookExW FAILED: {err}", flush=True)
            return
        print("[hook] Keyboard hook installed OK", flush=True)
        msg = ctypes.wintypes.MSG()
        while _user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
            _user32.TranslateMessage(ctypes.byref(msg))
            _user32.DispatchMessageW(ctypes.byref(msg))

    t = threading.Thread(target=_hook_thread, daemon=True)
    t.start()


# --- Overlay indicator window ---
class OverlayIndicator:
    """Always-on-top floating window that shows listening status."""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Speech")
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.attributes("-alpha", 0.85)

        screen_w = self.root.winfo_screenwidth()
        self.root.geometry(f"200x40+{screen_w - 220}+10")

        self.frame = tk.Frame(self.root, bg="#555555")
        self.frame.pack(fill="both", expand=True)

        self.label = tk.Label(
            self.frame,
            text="  MIC UIT",
            font=("Segoe UI", 14, "bold"),
            fg="white",
            bg="#555555",
            anchor="w",
            padx=10,
        )
        self.label.pack(side="left", fill="both", expand=True)

        self.close_btn = tk.Label(
            self.frame,
            text=" X ",
            font=("Segoe UI", 12, "bold"),
            fg="white",
            bg="#555555",
            cursor="hand2",
        )
        self.close_btn.pack(side="right", fill="y")
        self.close_btn.bind("<Button-1>", lambda e: quit_app())

        self.label.bind("<Button-1>", self._start_drag)
        self.label.bind("<B1-Motion>", self._on_drag)
        self._drag_x = 0
        self._drag_y = 0

    def _start_drag(self, event):
        self._drag_x = event.x
        self._drag_y = event.y

    def _on_drag(self, event):
        x = self.root.winfo_x() + event.x - self._drag_x
        y = self.root.winfo_y() + event.y - self._drag_y
        self.root.geometry(f"+{x}+{y}")

    def set_listening(self, active: bool):
        self.root.after(0, self._update, active)

    def _update(self, active: bool):
        bg = "#1B8C3A" if active else "#555555"
        text = "  LUISTERT..." if active else "  MIC UIT"
        self.frame.config(bg=bg)
        self.label.config(text=text, bg=bg)
        self.close_btn.config(bg=bg)

    def run(self):
        self.root.mainloop()

    def stop(self):
        self.root.after(0, self.root.destroy)


# --- Tray icon ---
def create_icon(active: bool) -> Image.Image:
    img = Image.new("RGB", (64, 64), color=(40, 40, 40))
    draw = ImageDraw.Draw(img)
    fill = (0, 200, 80) if active else (128, 128, 128)
    draw.ellipse([12, 12, 52, 52], fill=fill)
    return img


# --- Audio ---
def audio_callback(indata, frames, time_info, status):
    if listening:
        audio_queue.put(bytes(indata))


def recognition_loop():
    global listening, recognizer

    while True:
        data = audio_queue.get()
        if data is None:
            break
        if not listening:
            continue

        if recognizer.AcceptWaveform(data):
            result = json.loads(recognizer.Result())
            text = result.get("text", "").strip()
            if text:
                type_text(text)


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", ctypes.c_long),
        ("dy", ctypes.c_long),
        ("mouseData", ctypes.wintypes.DWORD),
        ("dwFlags", ctypes.wintypes.DWORD),
        ("time", ctypes.wintypes.DWORD),
        ("dwExtraInfo", ctypes.c_size_t),
    ]


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", ctypes.wintypes.WORD),
        ("wScan", ctypes.wintypes.WORD),
        ("dwFlags", ctypes.wintypes.DWORD),
        ("time", ctypes.wintypes.DWORD),
        ("dwExtraInfo", ctypes.c_size_t),
    ]


class HARDWAREINPUT(ctypes.Structure):
    _fields_ = [
        ("uMsg", ctypes.wintypes.DWORD),
        ("wParamL", ctypes.wintypes.WORD),
        ("wParamH", ctypes.wintypes.WORD),
    ]


class _INPUT_UNION(ctypes.Union):
    _fields_ = [("mi", MOUSEINPUT), ("ki", KEYBDINPUT), ("hi", HARDWAREINPUT)]


class INPUT(ctypes.Structure):
    _fields_ = [("type", ctypes.wintypes.DWORD), ("u", _INPUT_UNION)]


_user32.SendInput.argtypes = [ctypes.c_uint, ctypes.POINTER(INPUT), ctypes.c_int]
_user32.SendInput.restype = ctypes.c_uint

INPUT_KEYBOARD = 1
KEYEVENTF_UNICODE = 0x0004
KEYEVENTF_KEYUP = 0x0002


def type_text(text: str):
    """Type text using SendInput with KEYEVENTF_UNICODE — works in any app."""
    chars = text + " "
    inputs = []
    for ch in chars:
        code = ord(ch)
        down = INPUT()
        down.type = INPUT_KEYBOARD
        down.u.ki.wVk = 0
        down.u.ki.wScan = code
        down.u.ki.dwFlags = KEYEVENTF_UNICODE
        down.u.ki.dwExtraInfo = 0
        inputs.append(down)
        up = INPUT()
        up.type = INPUT_KEYBOARD
        up.u.ki.wVk = 0
        up.u.ki.wScan = code
        up.u.ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP
        up.u.ki.dwExtraInfo = 0
        inputs.append(up)

    n = len(inputs)
    arr = (INPUT * n)(*inputs)
    _user32.SendInput(n, arr, ctypes.sizeof(INPUT))


def play_beep(on: bool):
    freq = 800 if on else 400
    threading.Thread(
        target=winsound.Beep, args=(freq, 150), daemon=True
    ).start()


def start_listening():
    global listening
    if listening:
        return
    listening = True
    print("Luisteren: AAN", flush=True)
    play_beep(True)
    if tray_icon:
        tray_icon.icon = create_icon(True)
    if overlay:
        overlay.set_listening(True)


def stop_listening():
    global listening
    if not listening:
        return
    listening = False
    print("Luisteren: UIT", flush=True)
    play_beep(False)
    # Flush remaining audio from recognizer
    if recognizer:
        result = json.loads(recognizer.FinalResult())
        text = result.get("text", "").strip()
        if text:
            type_text(text)
    if tray_icon:
        tray_icon.icon = create_icon(False)
    if overlay:
        overlay.set_listening(False)


def quit_app():
    global listening
    listening = False
    audio_queue.put(None)
    if _hook_handle:
        _user32.UnhookWindowsHookEx(_hook_handle)
    if tray_icon:
        tray_icon.stop()
    if overlay:
        overlay.stop()


def on_quit(icon, item):
    quit_app()


def ensure_model():
    if os.path.exists(MODEL_PATH):
        return

    os.makedirs(APP_DATA_DIR, exist_ok=True)

    win = tk.Tk()
    win.title("Speech-to-Cursor \u2014 Model downloaden")
    win.geometry("450x100")
    win.resizable(False, False)
    tk.Label(win, text="Nederlands spraakmodel downloaden (~860 MB)...",
             font=("Segoe UI", 11)).pack(pady=(15, 5))
    progress_var = tk.StringVar(value="0%")
    tk.Label(win, textvariable=progress_var, font=("Segoe UI", 10)).pack()

    zip_path = MODEL_PATH + ".zip"
    error = [None]

    def do_download():
        try:
            def reporthook(block_num, block_size, total_size):
                downloaded = block_num * block_size
                if total_size > 0:
                    pct = min(100, downloaded * 100 // total_size)
                    mb = downloaded / (1024 * 1024)
                    total_mb = total_size / (1024 * 1024)
                    progress_var.set(f"{mb:.0f} / {total_mb:.0f} MB ({pct}%)")

            urllib.request.urlretrieve(MODEL_URL, zip_path, reporthook)

            progress_var.set("Uitpakken...")
            with zipfile.ZipFile(zip_path, "r") as zf:
                top = zf.namelist()[0].split("/")[0]
                zf.extractall(APP_DATA_DIR)
                extracted = os.path.join(APP_DATA_DIR, top)
                if extracted != MODEL_PATH:
                    os.rename(extracted, MODEL_PATH)

            os.remove(zip_path)
        except Exception as e:
            error[0] = str(e)
        finally:
            win.after(0, win.destroy)

    threading.Thread(target=do_download, daemon=True).start()
    win.mainloop()

    if error[0]:
        messagebox.showerror("Fout", f"Model downloaden mislukt:\n{error[0]}")
        sys.exit(1)


def main():
    global model, recognizer, tray_icon, overlay

    ensure_model()

    print("Model laden (dit kan even duren bij het grote model)...")
    model = Model(MODEL_PATH)
    recognizer = KaldiRecognizer(model, SAMPLE_RATE)
    print("Model geladen!")

    rec_thread = threading.Thread(target=recognition_loop, daemon=True)
    rec_thread.start()

    stream = sd.RawInputStream(
        samplerate=SAMPLE_RATE,
        blocksize=8000,
        dtype="int16",
        channels=1,
        callback=audio_callback,
    )
    stream.start()

    # Install system-wide keyboard hook (runs in its own thread with message pump)
    _start_keyboard_hook()
    print("Houd Ctrl+Spatie ingedrukt om te spreken (push-to-talk).")
    print("Druk op Ctrl+Shift+Q om af te sluiten.")

    overlay = OverlayIndicator()

    tray_icon = pystray.Icon(
        "speech",
        create_icon(False),
        "Speech-to-Cursor",
        menu=pystray.Menu(
            pystray.MenuItem("Afsluiten", on_quit),
        ),
    )
    tray_thread = threading.Thread(target=tray_icon.run, daemon=True)
    tray_thread.start()

    overlay.run()

    stream.stop()
    stream.close()
    print("Afgesloten.")


if __name__ == "__main__":
    main()
