"""
voice.py - Text-to-speech using pyttsx3 (offline, uses Windows SAPI5).

Runs on a single dedicated worker thread with a queue, so pyttsx3 (which is
not safe to call concurrently from multiple threads) stays happy. Fires
on_word() for every word spoken - the HUD uses this to "vibrate".

pyttsx3's SAPI5 driver on Windows has a well-known bug where reusing one
engine instance across multiple say()/runAndWait() calls works for the
FIRST utterance and then silently produces no audio afterward (this is
why a greeting can speak fine but every reply after it stays silent).
The documented workaround - and what this does - is to create a fresh
engine per utterance instead of reusing one for the app's whole
lifetime. Adds a small ~100-200ms overhead per line spoken, which isn't
noticeable for a voice assistant.

interrupt() lets JARVIS be talked over ("barge-in"): it stops whatever
is currently playing and drops anything still queued behind it, so a
new mic press or typed message doesn't have to wait for JARVIS to
finish a long reply first.
"""
import queue
import threading

import pyttsx3


class VoiceEngine:
    def __init__(self, on_word=None, on_start=None, on_end=None, rate=178):
        self.on_word = on_word
        self.on_start = on_start
        self.on_end = on_end
        self.rate = rate
        self._voice_id = None  # resolved once during the startup probe below

        self._q = queue.Queue()
        self._lock = threading.Lock()
        self._current_engine = None  # the engine actively mid-runAndWait(), if any

        self._thread = threading.Thread(target=self._worker, daemon=True)
        self._thread.start()

    def _make_engine(self):
        engine = pyttsx3.init()
        engine.setProperty("rate", self.rate)
        if self._voice_id:
            engine.setProperty("voice", self._voice_id)
        engine.connect(
            "started-word",
            lambda name, location, length: self.on_word() if self.on_word else None,
        )
        return engine

    def _worker(self):
        # One-time diagnostic pass: confirm pyttsx3 can init at all and pick
        # a preferred English voice - but this particular engine instance is
        # discarded rather than reused for actual speaking (see docstring).
        try:
            probe = pyttsx3.init()
            voices = probe.getProperty("voices")
            if not voices:
                print("[voice.py] WARNING: pyttsx3 found no installed voices. "
                      "TTS will silently produce no audio. On Windows, check "
                      "Settings > Time & Language > Speech, or reinstall a "
                      "SAPI5 voice.")
            for v in voices:
                name = (v.name or "").lower()
                if "english" in name or "en_" in (v.id or "").lower():
                    self._voice_id = v.id
                    break
            del probe
        except Exception as e:
            print(f"[voice.py] FATAL: pyttsx3 failed to initialize: {e!r}")
            print("[voice.py] JARVIS will run without voice output. "
                  "Try: pip install --upgrade pyttsx3  (Windows also needs "
                  "pypiwin32: pip install pypiwin32)")
            # Drain the queue forever so callers calling .speak() don't hang,
            # but nothing will ever be spoken.
            while True:
                if self._q.get() is None:
                    return

        while True:
            text = self._q.get()
            if text is None:
                break
            try:
                engine = self._make_engine()
                with self._lock:
                    self._current_engine = engine
                if self.on_start:
                    self.on_start()
                engine.say(text)
                engine.runAndWait()
                engine.stop()
                del engine
            except Exception as e:
                print(f"[voice.py] ERROR while speaking: {e!r}")
            finally:
                with self._lock:
                    self._current_engine = None
                if self.on_end:
                    self.on_end()

    def speak(self, text: str):
        """Queue text to be spoken. Non-blocking."""
        if text:
            self._q.put(text)

    def interrupt(self):
        """
        Barge-in: stops whatever's currently playing and discards
        anything still queued behind it, so JARVIS goes silent right
        away instead of finishing its current line (or its whole
        backlog) first.
        """
        # Drop everything waiting to be spoken.
        try:
            while True:
                self._q.get_nowait()
        except queue.Empty:
            pass

        with self._lock:
            engine = self._current_engine
        if engine is not None:
            try:
                engine.stop()
            except Exception as e:
                print(f"[voice.py] interrupt() couldn't stop the active engine: {e!r}")

    def stop(self):
        self._q.put(None)
