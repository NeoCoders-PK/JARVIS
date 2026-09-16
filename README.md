# J.A.R.V.I.S. Desktop

A desktop assistant backed by Groq's API, with a real HTML/CSS/SVG
UI (smooth glowing rotating rings, animated via CSS) rendered inside a
native window via `pywebview`. JARVIS can also open apps/websites and
read/write files for you, always asking for confirmation first.

The UI is genuinely a web page (`web/index.html` + `web/style.css` +
`web/app.js`) - Python does everything behind the scenes (Groq calls,
actions, voice, mic) and talks to the page over pywebview's JS bridge, the
same idea as the shop site but running inside its own window instead of a
browser tab.

## 1. Install Python

You need Python 3.9+ on Windows. Get it from python.org if you don't have
it (check "Add Python to PATH" during install).

## 2. Install dependencies

Open a terminal (Command Prompt / PowerShell) in this folder and run:

```
pip install requests pyttsx3 SpeechRecognition pywebview
```

```
pip install -r requirements.txt
```

The microphone (voice input) feature additionally needs `sounddevice` and
`numpy`:

```
pip install sounddevice numpy
```

Both ship precompiled wheels for every current Python version (no
compiler needed), unlike `pyaudio`, which this project intentionally
avoids - `pyaudio` only has precompiled Windows wheels up to Python
3.13, so on newer Pythons pip tries to build it from source and fails
without Microsoft's C++ Build Tools installed.

**If sounddevice won't install, that's fine** — the app detects this and
just hides the microphone button. You can still type to JARVIS and hear
it speak back.

`pywebview` on Windows uses the Edge WebView2 runtime, which ships with
Windows 10/11 by default. If the window fails to open with a WebView2
error, install it from https://developer.microsoft.com/microsoft-edge/webview2/.

## 3. Get a free Groq API key

1. Go to https://console.groq.com/keys and sign in (no payment card
   needed for the free tier).
2. Create an API key from the dashboard.
3. Either:
   - **Let the app ask you**: just run it (step 4) - a modal will ask for
     the key on first launch and save it to `config.json` automatically, or
   - **Set it up yourself**: copy `config.example.json` to `config.json`
     in this folder, and paste your key in place of
     `PASTE_YOUR_GROQ_API_KEY_HERE`:
     ```json
     {
       "api_key": "gsk_...your real key...",
       "model": "llama-3.3-70b-versatile"
     }
     ```
     `config.json` is never sent anywhere except to Groq's API with
     your requests. Free tier: 30 requests/min, 6,000 tokens/min,
     14,400 requests/day, shared across all Groq models. Check
     https://console.groq.com/docs/rate-limits for current numbers.

## 4. Run it

```
python main.py
```

If you haven't set up `config.json` yourself (step 3), a modal in the
window will ask for your Groq API key the first time you run it and
save it to `config.json` so you won't be asked again.

## What you get

- A real HTML/CSS/SVG HUD: three independently-rotating glowing rings +
  a pulsing core, using SVG `feGaussianBlur` for the glow and CSS
  `@keyframes` for smooth rotation - genuinely fluid, unlike the earlier
  Tkinter version.
- Text-to-speech (offline, via Windows SAPI through `pyttsx3`) — JARVIS
  reads its replies aloud, and the HUD pulses in sync with each spoken
  word (Python calls `pulseHUD()` in the page per word).
- Optional push-to-talk microphone input (a free speech-recognition
  backend via the `SpeechRecognition` library) if `sounddevice` is installed.
- Conversation memory within a session (sent as chat history to the model
  each turn) so JARVIS remembers context as you talk.
- **Actions with confirmation**: JARVIS can open a URL, launch an app/file,
  list a folder, read a file, or write/overwrite a file — including
  Word (`.docx`), Excel (`.xlsx`), PDF, and **PowerPoint (`.pptx`)**
  presentations built from notes you give it — but every single one pops
  up a plain-English "JARVIS wants to do something" Yes/No modal first.
  Nothing runs without you clicking Yes. Rich formats are read out as
  plain text and written back as brand-new documents (default template
  only) - not an in-place edit preserving the original's design.

## Files

| File                     | Purpose                                                     |
|--------------------------|---------------------------------------------------------------|
| `main.py`                | pywebview window + `Api` class (all Python-side logic)      |
| `web/index.html`         | Page structure: HUD, chat log, input bar, modals             |
| `web/style.css`          | HUD glow/rotation animations, chat + modal styling            |
| `web/app.js`             | UI wiring; exposes functions Python calls via `evaluate_js`  |
| `voice.py`               | Text-to-speech engine + per-word "vibrate" callback           |
| `stt.py`                 | Optional microphone speech-to-text                            |
| `assistant.py`           | Groq API call, incl. function calling                          |
| `actions.py`             | The actions JARVIS can propose (open/read/write/list)         |
| `presentation.py`        | Designed .pptx generation (theme, layouts, charts, images)    |
| `config.example.json`    | Template - copy to `config.json` and add your key             |
| `config.json`            | Created on first run (or by you) — holds your API key + model |

## How JS and Python talk to each other

- **JS → Python**: `window.pywebview.api.<method>(...)` calls a method on
  the `Api` class in `main.py` (e.g. `send_message`, `confirm_response`).
- **Python → JS**: `self.window.evaluate_js("someFunction(...)")` calls a
  global function defined in `app.js` (e.g. `addMessage`, `pulseHUD`,
  `showConfirm`). `json.dumps()` is used to safely escape any text going
  into that JS string.

## Presentations

Ask JARVIS for a deck ("make a presentation about X", "turn these notes
into slides") and it uses `presentation.py`, which produces a designed
deck rather than plain bullets:

- **Dark JARVIS theme** — teal on near-black, matching the app.
- **Nine layouts** — title, section divider, bullets, stat callouts,
  two-column compare, numbered steps, chart, full-bleed image, quote.
- **Native charts** — bar, column, line, pie, area, doughnut. These stay
  editable in PowerPoint (not flattened to images).
- **Slide transitions** — fade, push, wipe, morph, split.
- **Stock photos** from the web (optional, see below).

### Optional: real photos

Without a key, image slides render a styled placeholder panel. To pull
real photos, get a free API key at https://www.pexels.com/api/ and add it
to `config.json`:

```json
{
  "api_key": "gsk_...",
  "model": "openai/gpt-oss-120b",
  "pexels_api_key": "your_pexels_key"
}
```

Pexels photos are free to use commercially and don't require attribution,
but check their current license if the decks go somewhere public.

### Implementation notes

`python-pptx` has no API for slide transitions, so `presentation.py`
injects the transition XML directly into each slide. This is the standard
workaround and works in PowerPoint, but it's writing to an unsupported
surface — if a future library or PowerPoint release changes how that XML
parses, transitions could stop applying. The deck content itself is never
at risk; you'd just lose the transitions.

Per-element entrance animations (text flying in, etc.) are deliberately
not attempted — that XML is far more fragile, and getting it wrong makes
PowerPoint reject the whole file rather than degrading gracefully.

## Customizing

- **Colors / HUD shape**: edit the CSS variables at the top of
  `web/style.css` (`--glow`, `--glow-mid`, `--glow-dim`) and the ring radii
  / stroke-dasharray values in `web/index.html`'s `<svg>`.
- **Animation speed**: the `spin-slow` / `spin-rev` / `spin-fast`
  `animation-duration`s in `web/style.css`.
- **Voice**: change `rate` in `VoiceEngine(rate=...)` (words per minute),
  or pick a different installed Windows voice in `voice.py`.
- **Personality**: edit `SYSTEM_PROMPT` in `main.py`.
- **Wake word / always-listening**: not included by default (keeps the
  app simple and avoids accidental API usage) — the mic button is
  push-to-talk. Ask if you'd like a continuous "Hey JARVIS" wake-word mode
  added; it needs an extra offline wake-word library (e.g. `openwakeword`).

## Building a standalone .exe

You don't need Python installed to *run* JARVIS if you package it first:

1. Install PyInstaller (on top of everything in `requirements.txt`):
   ```
   pip install pyinstaller
   ```
2. Run `build_exe.bat` from inside this folder (or the equivalent
   command by hand - see the .bat file for what it runs).
3. Find `JARVIS.exe` in the new `dist\` folder. Run it directly - no
   Python needed on the machine you copy it to.

Two things specific to this project's setup:
- `config.json` is created **next to the .exe**, not inside it - so you
  can move/copy the .exe around and it'll remember your API key as long
  as `config.json` travels with it (or gets recreated on first run
  there).
- `build_exe.bat` already includes `--collect-all`/`--hidden-import`
  flags for pyttsx3, pywebview, pptx, sounddevice, reportlab, and
  certifi - each works around that specific library loading something
  at runtime (a driver, a backend, a template file, a compiled binary,
  font data, a CA bundle) that PyInstaller's static analysis can't see
  on its own. If you ever add a new dependency that behaves similarly,
  the same pattern applies: `--collect-all <package>` is the blunt but
  reliable fix.
