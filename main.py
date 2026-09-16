"""
main.py - JARVIS desktop assistant.

Run with:  python main.py

On first run it asks for your Groq API key (get a free one at
https://console.groq.com/keys) and saves it to config.json next to
this file.

The window is a pywebview-hosted view of web/index.html - real HTML/CSS/SVG
for smooth animated graphics, with this file doing everything Python-side
(Groq calls, actions, voice, mic) exactly as before. JS and Python talk
over pywebview's built-in bridge:
  JS -> Python:  window.pywebview.api.<method>(...)   (the Api class below)
  Python -> JS:  window.evaluate_js("someJsFunction(...)")
"""
import json
import os
import sys
import threading

import webview

from assistant import ask_groq
from voice import VoiceEngine
import actions

try:
    from stt import listen_once
    STT_AVAILABLE = True
except ImportError as e:
    print(f"[main.py] Mic input disabled - couldn't import stt.py: {e!r}")
    print("[main.py] This almost always means sounddevice or numpy isn't installed. Try:")
    print("[main.py]   pip install sounddevice numpy")
    STT_AVAILABLE = False

if getattr(sys, "frozen", False):
    # Running as a PyInstaller-built .exe. sys._MEIPASS is a temporary
    # folder that's re-extracted (and wiped) on every launch, so it's
    # only safe for read-only bundled assets like web/index.html -
    # config.json needs to live somewhere that actually persists, i.e.
    # next to the .exe itself.
    BASE_DIR = os.path.dirname(sys.executable)
    BUNDLE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    BUNDLE_DIR = BASE_DIR

CONFIG_FILE = os.path.join(BASE_DIR, "config.json")
INDEX_HTML = os.path.join(BUNDLE_DIR, "web", "index.html")

# GPT-OSS 120B is Groq's recommended replacement for the now-deprecated
# llama-3.3-70b-versatile (decommissioned Aug 16, 2026), and is on the
# free tier with solid tool-calling support. Check
# https://console.groq.com/docs/models for the current model list, and
# https://console.groq.com/docs/deprecations if this one ever 404s too.
DEFAULT_MODEL = "openai/gpt-oss-120b"

SYSTEM_PROMPT = (
    "You are JARVIS, a calm, witty, highly capable personal AI assistant. "
    "Keep replies conversational and reasonably concise, since they will be "
    "read aloud by a text-to-speech engine. Address the user naturally.\n\n"
    "You have tools to open apps/URLs and to read, list, or write files on "
    "the user's computer. Every tool call is shown to the user as a plain-"
    "English confirmation prompt before it runs, and only executes if they "
    "approve - so it's fine to propose an action whenever it's clearly what "
    "the user asked for. If the user declines an action, don't retry it; "
    "acknowledge it and ask what they'd like instead.\n\n"
    "write_file replaces a file's ENTIRE content - it does not patch just "
    "one line or section. So whenever the user asks you to edit, update, "
    "add to, or change part of an EXISTING file, you must call read_file "
    "on it first, then call write_file with the complete file content "
    "(the original text plus your change merged in) - never with just the "
    "changed part alone, or you will delete the rest of the file. Only "
    "skip the read step when creating a brand new file that doesn't exist "
    "yet.\n\n"
    "For .docx and .pdf files, write_file's content is plain text, one "
    "paragraph/line per line - it creates a brand new document, so tell "
    "the user this loses the original's formatting/images if they're "
    "editing an existing rich document, not just plain text. For .xlsx, "
    "content must be comma-separated values (one row per line, matching "
    "what read_file returns for that file).\n\n"
    "PRESENTATIONS: use create_presentation, NOT write_file, whenever the "
    "user wants slides, a deck, or a presentation - including 'turn these "
    "notes into a presentation' or 'make slides about X'. It produces a "
    "designed deck (dark teal theme, varied layouts, native charts, stock "
    "photos, slide transitions) instead of plain default-template bullets. "
    "Put real effort into the deck structure: vary the layouts across "
    "slides, open with a title slide, use section dividers, turn any "
    "numbers into a chart or stat callout rather than bullets, give slides "
    "an image search phrase, and keep on-slide text short with the detail "
    "in speaker notes. Only fall back to write_file for .pptx if "
    "create_presentation fails."
)

MAX_TOOL_ROUNDS = 6  # safety cap on function-call round-trips per message


class Api:
    """
    Every public method here is callable from JS as
    window.pywebview.api.<method_name>(...). This class owns all the
    Python-side state and logic; app.js is purely presentation.
    """

    def __init__(self):
        self.window = None  # set in main() after the window is created
        self.config_data = self._load_config()
        self.history = [{"role": "system", "content": SYSTEM_PROMPT}]

        self._confirm_event = threading.Event()
        self._confirm_result = False
        self._page_loaded = threading.Event()
        self._mic_stop_event = threading.Event()
        self._mic_listening = False

        self.voice = VoiceEngine(
            on_word=lambda: self._js("pulseHUD()"),
            on_start=lambda: self._js("setStatus('SPEAKING')"),
            on_end=lambda: self._js("setStatus('STANDBY')"),
        )

    # ---- called once the window exists ---------------------------------

    def _mark_page_loaded(self):
        """
        Bound to window.events.loaded. Must stay trivial/non-blocking -
        it runs ON the GUI thread, and anything here that calls
        evaluate_js() would deadlock (that's what caused the window to
        freeze earlier). It only flips a flag; the real startup work
        happens in run() below, on its own thread.
        """
        self._page_loaded.set()

    def run(self):
        """
        Passed to webview.start(func=...), so this runs on its own thread
        once the GUI loop is up - safe to call evaluate_js() from here,
        unlike from the loaded event directly (deadlock) or immediately on
        thread start (window/page may not have actually finished loading
        yet - a fixed sleep() guessed at this and wasn't reliable).
        """
        if not self._page_loaded.wait(timeout=15):
            print("[main.py] WARNING: page 'loaded' event never fired within "
                  "15s - proceeding anyway, but evaluate_js calls may fail.")

        self._js(f"setMicState({'idle' if STT_AVAILABLE else 'hidden'!r})")
        if not self.config_data.get("api_key", "").strip():
            self._js("showApiKeyPrompt()")
            return  # save_api_key() (called from JS) continues startup
        self._greet()

    def _greet(self):
        self._add_message("jarvis", "Online. How can I help, sir?")

    # ---- config ----------------------------------------------------------

    def _load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r") as f:
                    return json.load(f)
            except json.JSONDecodeError:
                return {}
        return {}

    def _save_config(self):
        with open(CONFIG_FILE, "w") as f:
            json.dump(self.config_data, f, indent=2)

    def save_api_key(self, key):
        """Called from JS when the user submits the API-key modal."""
        if key and key.strip():
            self.config_data["api_key"] = key.strip()
            self.config_data.setdefault("model", DEFAULT_MODEL)
            self._save_config()
            self._greet()
        else:
            self._js("showApiKeyPrompt()")
        return True

    # ---- JS bridge helpers -----------------------------------------------

    def _js(self, code):
        if not self.window:
            return
        try:
            self.window.evaluate_js(code)
        except Exception as e:
            print(f"[main.py] evaluate_js failed (usually harmless/transient): {e!r}")

    def _add_message(self, sender, text, speak_text=None):
        # json.dumps gives us safe JS-string escaping for free.
        self._js(f"addMessage({json.dumps(sender)}, {json.dumps(text)})")
        # Read every non-user message aloud, no matter what triggered it -
        # errors, "didn't catch that", action prompts, everything. Pass
        # speak_text to say something different aloud than what's shown in
        # the chat bubble (e.g. a fuller sentence for an action prompt);
        # otherwise it speaks `text` verbatim.
        if sender != "user":
            self.voice.speak(speak_text if speak_text is not None else text)

    # ---- methods exposed to JS --------------------------------------------

    def send_message(self, text):
        threading.Thread(target=self._process, args=(text,), daemon=True).start()
        return True

    def toggle_mic(self):
        """
        Single entry point for the mic button, called on every click
        regardless of what the button currently looks like. Python's own
        self._mic_listening flag decides start vs. stop - not the button's
        CSS class - because setMicState() calls into JS from a background
        thread can silently fail on some WebView2 setups (visible as
        'CoreWebView2Controller members can only be accessed from the UI
        thread' in the console), which would desync the visual state from
        what's actually happening and made stopping unreliable.
        """
        if not STT_AVAILABLE:
            return False
        if self._mic_listening:
            self._mic_stop_event.set()
        else:
            # Barge-in: pressing the mic while JARVIS is talking cuts it
            # off immediately instead of making you wait for it to finish.
            self.voice.interrupt()
            self._mic_listening = True
            self._mic_stop_event.clear()
            threading.Thread(target=self._mic_thread, daemon=True).start()
        return True

    def interrupt(self):
        """Stop JARVIS talking right now, without starting the mic."""
        self.voice.interrupt()
        return True

    def confirm_response(self, approved: bool):
        """Called by app.js when the user clicks Yes/No on the confirm modal."""
        self._confirm_result = bool(approved)
        self._confirm_event.set()
        return True

    # ---- background work --------------------------------------------------

    def _mic_thread(self):
        self._js("setMicState('recording')")
        self._js("setStatus('LISTENING')")
        text = listen_once(self._mic_stop_event)
        self._mic_listening = False
        self._js("setMicState('idle')")
        if text:
            self._add_message("user", text)
            self._js("setStatus('PROCESSING')")
            self._process(text)
        else:
            self._js("setStatus('STANDBY')")
            self._add_message("jarvis", "(didn't catch that)")

    def _confirm_action(self, prompt_text):
        """
        Blocks the calling (background) thread until the user answers the
        Yes/No modal in the page. confirm_response() (called from JS) sets
        the Event this is waiting on.
        """
        self._confirm_event.clear()
        self._js(f"showConfirm({json.dumps(prompt_text)})")
        self._confirm_event.wait()
        return self._confirm_result

    def _process(self, text):
        self._js("setInputEnabled(false)")
        self._js("setStatus('PROCESSING')")
        self.history.append({"role": "user", "content": text})

        for _ in range(MAX_TOOL_ROUNDS):
            try:
                result = ask_groq(
                    self.history,
                    self.config_data["api_key"],
                    self.config_data.get("model", DEFAULT_MODEL),
                    function_declarations=actions.FUNCTION_DECLARATIONS,
                )
            except Exception as e:
                reply = f"I hit an error reaching Groq: {e}"
                self.history.append({"role": "assistant", "content": reply})
                self._add_message("jarvis", reply)
                self._js("setInputEnabled(true)")
                return

            if result["type"] == "text":
                reply = result["text"]
                self.history.append({"role": "assistant", "content": reply})
                self._add_message("jarvis", reply)
                self._js("setInputEnabled(true)")
                return

            # result["type"] == "function_calls" - OpenAI-style APIs (Groq
            # included) can ask for more than one tool call in a single
            # turn, so this appends ONE assistant message listing every
            # call, then must follow it with exactly one "tool" response
            # per call before the next request - the API 400s otherwise.
            calls = result["calls"]
            self.history.append({
                "role": "assistant",
                "content": None,
                "tool_calls": [
                    {
                        "id": c["id"],
                        "type": "function",
                        "function": {"name": c["name"], "arguments": json.dumps(c["args"])},
                    }
                    for c in calls
                ],
            })

            for c in calls:
                name, args = c["name"], c["args"]
                prompt_text = actions.confirm_text_for(name, args)
                self._add_message("action", f"wants to: {name} - confirm in the popup", speak_text=prompt_text)
                approved = self._confirm_action(prompt_text)

                if approved:
                    outcome = actions.run_action(name, args)
                else:
                    outcome = {"error": "The user declined this action."}

                self.history.append({
                    "role": "tool",
                    "tool_call_id": c["id"],
                    "content": json.dumps(outcome),
                })
            # loop again so the model can react to the outcome(s) / propose the next step

        reply = "That turned into more steps than I'm comfortable auto-chaining - let's pause there."
        self.history.append({"role": "assistant", "content": reply})
        self._add_message("jarvis", reply)
        self._js("setInputEnabled(true)")


def main():
    api = Api()
    window = webview.create_window(
        "J.A.R.V.I.S.",
        INDEX_HTML,
        js_api=api,
        width=1000,
        height=700,
        min_size=(760, 520),
        background_color="#020a12",
    )
    api.window = window
    window.events.loaded += api._mark_page_loaded
    # gui='edgechromium' forces the modern Edge WebView2 content renderer
    # (better CSS support than letting it silently choose). The earlier
    # "window.native.AccessibilityObject.Bounds.Empty.Empty..." crash
    # turned out to be unrelated to renderer choice - it was evaluate_js()
    # getting called before the page had genuinely finished loading. That's
    # fixed by waiting on window.events.loaded (see _mark_page_loaded/run)
    # rather than guessing with a fixed delay.
    webview.start(api.run, gui="edgechromium", debug=False)


if __name__ == "__main__":
    main()
