# Speech-to-Cursor

Windows spraakherkenning die gesproken tekst typt op de plek waar je cursor staat.
Gebruikt [faster-whisper](https://github.com/SYSTRAN/faster-whisper) met het `small` model, geoptimaliseerd voor Nederlands.

## Installatie

```bash
pip install -r requirements.txt
```

Het Whisper-model wordt automatisch gedownload bij de eerste keer opstarten.

## Gebruik

```bash
python speech.py
```

- Start de applicatie — er verschijnt een overlay-venster en een icoon in de systeem-tray
- Houd **Ctrl+Spatie** ingedrukt om te spreken (push-to-talk)
- Laat los — de herkende tekst wordt getypt waar je cursor staat
- **Ctrl+Shift+Q** om af te sluiten (of klik op de X in de overlay)
