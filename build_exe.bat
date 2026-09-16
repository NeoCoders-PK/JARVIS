@echo off
REM Builds JARVIS into a single-file Windows .exe using PyInstaller.
REM Run this from inside the jarvis/ folder, after installing all
REM requirements.txt packages (and pyinstaller itself).
REM
REM The --collect-all / --hidden-import flags below aren't optional
REM extras - each one works around a specific, well-documented
REM PyInstaller blind spot in one of this project's dependencies:
REM
REM   pyttsx3     picks its platform driver (sapi5 on Windows) via a
REM               dynamic importlib call, which PyInstaller's static
REM               analysis can't see - without --collect-all, the exe
REM               runs but is silently mute (or crashes when speaking).
REM   pywebview   picks its GUI backend (edgechromium on Windows) the
REM               same dynamic way.
REM   pptx        Presentation() loads a bundled default.pptx template
REM               FILE at runtime - not a Python import, so PyInstaller
REM               never bundles it unless told to. Without this, every
REM               create_presentation call fails as soon as it's frozen.
REM   sounddevice ships a compiled PortAudio binary as package data,
REM               loaded via cffi at runtime rather than a normal import.
REM   reportlab   bundles its own font metrics as data files, needed for
REM               .pdf writing.
REM   certifi     requests' HTTPS calls (to Groq, Pexels) need certifi's
REM               CA bundle file, which is data, not code - without it
REM               you get SSL verify errors calling out from the .exe.

pyinstaller --noconfirm --onefile --windowed --name JARVIS ^
    --add-data "web;web" ^
    --collect-all pyttsx3 ^
    --collect-all pywebview ^
    --collect-all pptx ^
    --collect-all sounddevice ^
    --collect-data reportlab ^
    --collect-data certifi ^
    --hidden-import pyttsx3.drivers ^
    --hidden-import pyttsx3.drivers.sapi5 ^
    main.py

echo.
echo Done. Find JARVIS.exe in the dist\ folder.
echo First run there will ask for your Groq API key and save it as
echo config.json next to the .exe - keep the two together.
