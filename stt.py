"""
stt.py - Optional microphone input via `sounddevice` (records audio) +
`SpeechRecognition`'s recognize_google() (transcribes it) - both free.

This intentionally avoids the `pyaudio` package. PyAudio only ships
precompiled Windows wheels up to Python 3.13; on 3.14+ pip falls back
to compiling it from source, which needs Microsoft's C++ Build Tools
and fails without them. `sounddevice` loads its audio backend
(PortAudio) at runtime via `cffi` instead of compiling against a
specific Python version, so its wheels work on any Python release with
nothing to compile.

If sounddevice isn't installed, main.py catches the ImportError and
simply hides the mic button - typing still works fine.

Recording streams in small (~100ms) chunks rather than one fixed-length
blocking call, so a `stop_event` can interrupt it almost immediately -
this is what lets main.py stop listening the moment you press the mic
button again, instead of waiting out the full phrase_time_limit.
"""
import time

import numpy as np
import sounddevice as sd
import speech_recognition as sr

SAMPLE_RATE = 16000  # Hz - what Google's recognizer expects
SAMPLE_WIDTH = 2  # bytes per sample (16-bit PCM)
CHUNK_SAMPLES = SAMPLE_RATE // 10  # ~100ms per chunk - how quickly stop_event is noticed


def listen_once(stop_event, timeout=5, phrase_time_limit=30):
    """
    Records from the default microphone until either stop_event is set
    (the mic button was pressed again) or phrase_time_limit seconds
    pass, whichever comes first. Returns the recognized text, or None
    if nothing usable was heard. `timeout` is unused (kept for
    interface compatibility) since recording starts immediately.
    """
    print(f"[stt.py] Listening (press mic again to stop, or up to {phrase_time_limit}s)...")
    chunks = []
    try:
        with sd.InputStream(samplerate=SAMPLE_RATE, channels=1, dtype="int16") as stream:
            start = time.monotonic()
            while not stop_event.is_set() and (time.monotonic() - start) < phrase_time_limit:
                data, _ = stream.read(CHUNK_SAMPLES)
                chunks.append(data.copy())
    except Exception as e:
        print(f"[stt.py] FATAL: couldn't open a microphone: {e!r}")
        print("[stt.py] Check Windows Settings > Privacy > Microphone access, "
              "and that a default input device is set.")
        return None

    if not chunks:
        print("[stt.py] Stopped before anything was recorded.")
        return None

    recording = np.concatenate(chunks)
    audio = sr.AudioData(recording.tobytes(), SAMPLE_RATE, SAMPLE_WIDTH)
    r = sr.Recognizer()
    try:
        return r.recognize_google(audio)
    except sr.UnknownValueError:
        print("[stt.py] Heard audio but couldn't understand any words.")
        return None
    except sr.RequestError as e:
        print(f"[stt.py] FATAL: speech-recognition API request failed: {e!r}")
        print("[stt.py] This needs an internet connection - check you're online.")
        return None
