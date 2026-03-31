"""Download the Vosk speech recognition model for Dutch."""

import os
import sys
import urllib.request
import zipfile

MODEL_URL = "https://alphacephei.com/vosk/models/vosk-model-small-nl-0.22.zip"
MODEL_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "model")


def download_model():
    if os.path.exists(MODEL_DIR):
        print(f"Model directory already exists: {MODEL_DIR}")
        print("Delete it first if you want to re-download.")
        return

    zip_path = MODEL_DIR + ".zip"
    print(f"Downloading Dutch speech model from {MODEL_URL} ...")
    urllib.request.urlretrieve(MODEL_URL, zip_path, _progress)
    print("\nExtracting...")

    with zipfile.ZipFile(zip_path, "r") as zf:
        # The zip contains a top-level folder; extract and rename to "model"
        top = zf.namelist()[0].split("/")[0]
        zf.extractall(os.path.dirname(MODEL_DIR))
        os.rename(os.path.join(os.path.dirname(MODEL_DIR), top), MODEL_DIR)

    os.remove(zip_path)
    print(f"Model ready at {MODEL_DIR}")


def _progress(block_num, block_size, total_size):
    downloaded = block_num * block_size
    if total_size > 0:
        pct = min(100, downloaded * 100 // total_size)
        mb = downloaded / (1024 * 1024)
        total_mb = total_size / (1024 * 1024)
        sys.stdout.write(f"\r  {mb:.1f} / {total_mb:.1f} MB ({pct}%)")
        sys.stdout.flush()


if __name__ == "__main__":
    download_model()
