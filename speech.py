"""
Speech-to-Cursor: Windows spraakherkenning die gesproken tekst typt waar je cursor staat.

Gebruik:
  - Ctrl+Shift+S  : start/stop luisteren
  - Systeem-tray icoon: rechtsklik om af te sluiten
"""

import json
import os
import queue
import sys
import threading

import keyboard
import pystray
import sounddevice as sd
from PIL import Image, ImageDraw
from vosk import KaldiRecognizer, Model

# --- Configuration ---
MODEL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "model")
SAMPLE_RATE = 16000
HOTKEY = "ctrl+shift+s"

# --- Global state ---
audio_queue = queue.Queue()
listening = False
recognizer = None
model = None
tray_icon = None


def create_icon(active: bool) -> Image.Image:
    """Create a simple tray icon — green when listening, grey when idle."""
    img = Image.new("RGB", (64, 64), color=(40, 40, 40))
    draw = ImageDraw.Draw(img)
    fill = (0, 200, 80) if active else (128, 128, 128)
    draw.ellipse([12, 12, 52, 52], fill=fill)
    return img


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
        else:
            # Partial results — we ignore these for now
            pass


def type_text(text: str):
    """Type text at the current cursor position using keyboard simulation."""
    # Add a space after each recognised phrase so words don't stick together
    keyboard.write(text + " ")


def toggle_listening():
    """Toggle speech recognition on/off."""
    global listening
    listening = not listening
    status = "AAN" if listening else "UIT"
    print(f"Luisteren: {status}")
    if tray_icon:
        tray_icon.icon = create_icon(listening)


def on_quit(icon, item):
    """Quit the application."""
    global listening
    listening = False
    audio_queue.put(None)  # Signal recognition thread to stop
    icon.stop()


def main():
    global model, recognizer, tray_icon

    # Check model exists
    if not os.path.exists(MODEL_PATH):
        print("Spraakmodel niet gevonden. Voer eerst uit:  python download_model.py")
        sys.exit(1)

    print("Model laden...")
    model = Model(MODEL_PATH)
    recognizer = KaldiRecognizer(model, SAMPLE_RATE)

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

    # Register hotkey
    keyboard.add_hotkey(HOTKEY, toggle_listening)
    print(f"Druk op {HOTKEY.upper()} om luisteren te starten/stoppen.")

    # System tray icon
    tray_icon = pystray.Icon(
        "speech",
        create_icon(False),
        "Speech-to-Cursor",
        menu=pystray.Menu(
            pystray.MenuItem("Luisteren aan/uit", lambda icon, item: toggle_listening()),
            pystray.MenuItem("Afsluiten", on_quit),
        ),
    )

    # pystray.run() blocks — it runs the event loop
    tray_icon.run()

    # Cleanup after tray icon stops
    stream.stop()
    stream.close()
    keyboard.remove_hotkey(HOTKEY)
    print("Afgesloten.")


if __name__ == "__main__":
    main()
