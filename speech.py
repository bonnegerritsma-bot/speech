"""
Speech-to-Cursor: Windows spraakherkenning die gesproken tekst typt waar je cursor staat.

Gebruik:
  - Ctrl+Shift+S  : start/stop luisteren
  - Ctrl+Shift+Q  : afsluiten
  - Overlay-venster: sluitknop (X) of versleep naar andere positie
"""

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

import keyboard
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
HOTKEY = "ctrl+shift+s"
QUIT_HOTKEY = "ctrl+shift+q"

# --- Global state ---
audio_queue = queue.Queue()
listening = False
recognizer = None
model = None
tray_icon = None
overlay = None


# --- Overlay indicator window ---
class OverlayIndicator:
    """Always-on-top floating window that shows listening status."""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Speech")
        self.root.overrideredirect(True)  # No window border
        self.root.attributes("-topmost", True)  # Always on top
        self.root.attributes("-alpha", 0.85)  # Slightly transparent

        # Position in top-right corner of screen
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

        # Allow dragging the overlay
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
        """Update overlay to show listening state. Thread-safe."""
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
    """Create a simple tray icon — green when listening, grey when idle."""
    img = Image.new("RGB", (64, 64), color=(40, 40, 40))
    draw = ImageDraw.Draw(img)
    fill = (0, 200, 80) if active else (128, 128, 128)
    draw.ellipse([12, 12, 52, 52], fill=fill)
    return img


# --- Audio ---
def audio_callback(indata, frames, time_info, status):
    """Called by sounddevice for each audio block."""
    if listening:
        audio_queue.put(bytes(indata))


def recognition_loop():
    """Continuously process audio and type recognised text."""
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


def type_text(text: str):
    """Type text at the current cursor position using keyboard simulation."""
    keyboard.write(text + " ")


def play_beep(on: bool):
    """Play a short beep — high tone for on, low tone for off."""
    freq = 800 if on else 400
    threading.Thread(
        target=winsound.Beep, args=(freq, 150), daemon=True
    ).start()


def toggle_listening():
    """Toggle speech recognition on/off."""
    global listening
    listening = not listening
    status = "AAN" if listening else "UIT"
    print(f"Luisteren: {status}")

    play_beep(listening)

    if tray_icon:
        tray_icon.icon = create_icon(listening)
    if overlay:
        overlay.set_listening(listening)


def quit_app():
    """Quit the application from any context."""
    global listening
    listening = False
    audio_queue.put(None)
    if tray_icon:
        tray_icon.stop()
    if overlay:
        overlay.stop()


def on_quit(icon, item):
    """Quit from tray menu."""
    quit_app()


def ensure_model():
    """Download the speech model if not present, with a progress window."""
    if os.path.exists(MODEL_PATH):
        return

    os.makedirs(APP_DATA_DIR, exist_ok=True)

    # Show progress window
    win = tk.Tk()
    win.title("Speech-to-Cursor — Model downloaden")
    win.geometry("450x100")
    win.resizable(False, False)
    tk.Label(win, text="Nederlands spraakmodel downloaden (~860 MB)...",
             font=("Segoe UI", 11)).pack(pady=(15, 5))
    progress_var = tk.StringVar(value="0%")
    progress_label = tk.Label(win, textvariable=progress_var, font=("Segoe UI", 10))
    progress_label.pack()

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

    # Start recognition thread
    rec_thread = threading.Thread(target=recognition_loop, daemon=True)
    rec_thread.start()

    # Start audio stream
    stream = sd.RawInputStream(
        samplerate=SAMPLE_RATE,
        blocksize=8000,
        dtype="int16",
        channels=1,
        callback=audio_callback,
    )
    stream.start()

    # Register hotkeys
    keyboard.add_hotkey(HOTKEY, toggle_listening)
    keyboard.add_hotkey(QUIT_HOTKEY, quit_app)
    print(f"Druk op {HOTKEY.upper()} om luisteren te starten/stoppen.")
    print(f"Druk op {QUIT_HOTKEY.upper()} om af te sluiten.")

    # Create overlay indicator
    overlay = OverlayIndicator()

    # System tray icon (runs in background thread)
    tray_icon = pystray.Icon(
        "speech",
        create_icon(False),
        "Speech-to-Cursor",
        menu=pystray.Menu(
            pystray.MenuItem("Luisteren aan/uit", lambda icon, item: toggle_listening()),
            pystray.MenuItem("Afsluiten", on_quit),
        ),
    )
    tray_thread = threading.Thread(target=tray_icon.run, daemon=True)
    tray_thread.start()

    # Overlay runs on main thread (tkinter requires main thread on Windows)
    overlay.run()

    # Cleanup after overlay closes
    stream.stop()
    stream.close()
    keyboard.remove_hotkey(HOTKEY)
    keyboard.remove_hotkey(QUIT_HOTKEY)
    print("Afgesloten.")


if __name__ == "__main__":
    main()
