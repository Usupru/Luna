<h1>🌙 Luna Assistant</h1>

Luna is a Windows desktop voice assistant focused on fast actions, automation, and customization.
It listens for hotkeys, recognizes Spanish voice commands, and runs custom actions on your own PC.

<img width="220" alt="Luna logo" src="LunaLogo.png" />

<h2>🚀 Features</h2>

<ul>
<li>🎙️ Offline speech recognition with Vosk (Spanish model)</li>
<li>🗣️ Speech synthesis with Piper TTS (pyttsx3 fallback)</li>
<li>🎮 Custom keyword actions (open apps, run commands, fixed/random replies)</li>
<li>🎵 Spotify playback control (play, pause, next, previous, volume)</li>
<li>🌦️ Weather queries using OpenWeather API</li>
<li>🖥️ Windows actions (shutdown, restart, screenshot, Alt+Tab)</li>
<li>🧰 Built-in control panel and setup wizard</li>
</ul>

<h2>📦 Installation</h2>

<h3>Option 1: Windows installer (.exe wizard)</h3>

If you downloaded the installer build:
1. Run the installer `.exe`.
2. Follow the installation wizard.
3. Launch Luna from the Start Menu or desktop shortcut.

<h3>Option 2: From source (git clone)</h3>

Make sure you are on Windows and have Python 3.11+ installed.

Clone the repository:
````
git clone <TU_REPO_URL>
cd Luna
````

Create and activate a virtual environment:
````
python -m venv venv
venv\Scripts\activate
````

Install dependencies:
````
pip install piper-tts vosk keyboard psutil pyaudio pyautogui pyperclip pyttsx3 requests wikipedia winotify spotipy sounddevice numpy pystray pillow pywin32
````

<h3>Required assets</h3>

Be aware that Luna expects these folders (or equivalent paths via environment variables):
````
Sonidos/
Modelos/
Imagenes/
````
Therefore, it is adviced for anyone looking to modify the source code, to download these folders via the windows installer.

Common files:
- `Sonidos/IniciarSound.wav`
- Piper voice model `.onnx` (for example `es_AR-daniela-high.onnx`)
- Spanish Vosk model folder (for example `vosk-model-small-es-0.42`)

<h2>⚙️ Usage</h2>

Start the app:
````
python main.py
````

On first launch, complete the setup wizard:
- Assistant name and voice model
- Optional custom keyword actions
- Optional OpenWeather and Spotify credentials

Default hotkeys:
- `Ctrl + |` -> listen for a voice command
- `Ctrl + Shift + |` -> reopen control panel

<h2>🔧 Configuration</h2>

Main runtime config file:
````
data/config/app_config.json
````

Useful environment variables:
- `LUNA_DATA_DIR`
- `LUNA_ASSETS_DIR`
- `LUNA_SOUNDS_DIR`
- `LUNA_MODELS_DIR`
- `VOSK_MODEL_PATH`
- `OPENWEATHER_API_KEY`
- `SPOTIPY_CLIENT_ID`
- `SPOTIPY_CLIENT_SECRET`
- `SPOTIPY_REDIRECT_URI`
- `SPOTIPY_CACHE_PATH`

<h2>🧠 Notes</h2>

<ul>
<li>This project is Windows-focused (uses <code>winreg</code>, <code>win32clipboard</code>, startup registry integration, and toast notifications).</li>
<li>Some commands require external services: Spotify Web API credentials and an OpenWeather API key.</li>
<li>If Piper/Vosk models are missing, speech recognition or TTS will be limited.</li>
<li>Depending on settings, Luna can keep running in the background after the panel is closed.</li>
</ul>
