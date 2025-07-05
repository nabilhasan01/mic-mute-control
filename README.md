# Microphone Mute Control

A lightweight Windows application to toggle microphone mute status with a hotkey, displaying a 48x48 red microphone icon (with white slash) in the top-middle of the screen when muted. Features a system tray icon, GUI for hotkey and overlay customization, external mute detection, and customizable sound feedback with bundled default sounds. Available in two implementations: Tkinter (`mic_state_controller_tkinter.py`) and PyQt6 (`mic_state_controller_pyqt.py`).

## Features
- **Toggle Mute**: Mute/unmute the default microphone using a customizable hotkey (default: `Ctrl+Alt+M`).
- **Overlay Icon**: Displays a 48x48 vector-based red microphone icon with a white slash (no shadow) in the top-middle of the screen when muted, with 70% transparency for minimal obstruction.
- **System Tray**: Shows a tray icon (red "M" for muted, green "U" for unmuted) with a right-click menu to toggle mute, show GUI, or exit.
- **Hotkey Customization**: Set a new hotkey via the GUI (e.g., `Ctrl+Shift+M`, `Pause`) with real-time capture.
- **Overlay Customization**: Adjust overlay position (e.g., Top Mid, Bottom Right), size (32x32, 48x48, 64x64), margin (0-50 pixels), and opacity (0.1-1.0) via the GUI.
- **External Mute Detection**: Polls every 100ms (PyQt6) or 500ms (Tkinter) to detect mute/unmute changes from other sources (e.g., keyboard mute key).
- **Sound Feedback**: Play custom WAV files or bundled default sounds (`_mute.wav` for mute, `_unmute.wav` for unmute) on mute/unmute (configurable via GUI). Falls back to a default beep if custom sounds fail (Tkinter only).
- **Start Minimized**: Launches directly to the system tray, with the GUI hidden until requested.
- **Start with Windows**: Option to add the application to Windows startup for automatic launching.
- **Cross-Resolution Support**: Overlay dynamically centers based on screen resolution (e.g., `x=936` for 1920x1080).
- **Configuration Persistence**: Saves overlay settings (position, size, margin, opacity), sound file paths, and hotkey to `~/.mic_mute_app/config.json` for both script and executable, ensuring settings persist across restarts.
- **PyQt6 Enhancements** (in `mic_state_controller_pyqt.py`):
  - Threaded hotkey capture for non-blocking GUI.
  - Debounced toggle mechanism to prevent rapid successive toggles.
  - Mutex-based synchronization for safe audio operations.
  - Retry logic for audio device initialization and mute toggling.

## Requirements
- Windows 10/11 (64-bit).
- Python 3.11+ (for development; not needed for the `.exe`).
- `libcairo-2.dll` (included in the `resource` folder of the repository).

## Installation (Development)
1. Clone the repository:
   ```bash
   git clone https://github.com/nabilhasan01/mic-mute-control.git
   cd mic-mute-control
   ```
2. Install dependencies for either Tkinter or PyQt6 version:
   - For Tkinter (`mic_state_controller_tkinter.py`):
     ```bash
     pip install pycaw comtypes psutil pywin32 keyboard pystray Pillow pyinstaller cairosvg cairocffi pygame
     ```
   - For PyQt6 (`mic_state_controller_pyqt.py`):
     ```bash
     pip install pycaw comtypes psutil pywin32 keyboard pystray Pillow pyinstaller cairosvg cairocffi pygame PyQt6
     ```
3. Verify that `libcairo-2.dll` is present in the `resource` folder (included in the repository).

## Usage
1. Run the desired script (preferably as administrator for hotkey support):
   - For Tkinter:
     ```bash
     python mic_state_controller_tkinter.py
     ```
   - For PyQt6:
     ```bash
     python mic_state_controller_pyqt.py
     ```
   - The app starts minimized to the system tray if configured.
2. **Toggle Mute**:
   - Press `Ctrl+Al t+M` (default hotkey).
   - Or right-click the tray icon and select "Toggle Mute".
   - A 48x48 red mic icon with a white slash appears in the configured position when muted.
   - Default bundled sounds (`_mute.wav` for mute, `_unmute.wav` for unmute) or custom WAV sounds play on mute/unmute transitions.
3. **Customize Hotkey**:
   - Right-click tray icon, select "Show Window".
   - Click "Set Hotkey", then press a key combination (e.g., `Ctrl+Shift+M`).
4. **Customize Overlay**:
   - Open the GUI, adjust "Position" (e.g., Top Mid), "Size" (e.g., 48x48), "Margin" (e.g., Digest, or "Opacity" (e.g., 0.7).
   - Changes apply instantly and are saved to `~/.mic_mute_app/config.json`.
5. **Configure Sound**:
   - Open the GUI, click "Browse" to select custom WAV files for mute and unmute sounds, then click "Apply".
   - Default sounds (`_mute.wav`, `_unmute.wav`) are used if no custom files are selected.
6. **Configure Startup**:
   - Enable "Start Minimized to Tray" to launch directly to the system tray.
   - Enable "Start with Windows" to add the app to Windows startup.
7. **Exit**:
   - Right-click tray icon, select "Exit".

## Building the Executable
To create a standalone `.exe` with bundled default sound files (`_mute.wav`, `_unmute.wav`), `config.json`, and administrator privileges:

1. Ensure `config.json` exists in the project directory with default settings (e.g., `{"overlay_position": "Top Mid", "overlay_size": "48x48", "overlay_margin": 10, "overlay_opacity": 0.7, "mute_sound_file": "resource\\_mute.wav", "unmute_sound_file": "resource\\_unmute.wav", "start_minimized": false, "start_with_windows": false, "hotkey": "ctrl+alt+m"}`).
2. Run the provided batch script to build the executable for either version:
   - For Tkinter:
     ```bash
     build_mic_mute_exe.bat tkinter
     ```
   - For PyQt6:
     ```bash
     build_mic_mute_exe.bat pyqt
     ```
   - The script creates a manifest file to ensure the executable requests administrator privileges and runs `PyInstaller` with the necessary parameters.
   - The script includes all required dependencies and resources (`libcairo-2.dll`, `_mute.wav`, `_unmute.wav`, `config.json`).
   - Settings are saved to `~/.mic_mute_app/config.json` to persist across restarts.
   - If `comtypes.gen` is needed (check for `pycaw` errors), modify the batch script to include:
     ```bash
     --add-data "C:\Users\YourUsername\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.11_qbz5n2kfra8p0\LocalCache\local-packages\Python311\site-packages\comtypes\gen;comtypes\gen"
     ```
3. Find `mic_mute_tkinter.exe` or `mic_mute_pyqt.exe` in the `dist` folder.
4. Share the `.exe` (run as administrator on other PCs).

### Manual PyInstaller Command (Alternative)
If you prefer to build manually, use the appropriate command:
- For Tkinter:
  ```bash
  pyinstaller --onefile --windowed ^
    --hidden-import=pycaw ^
    --hidden-import=comtypes ^
    --hidden-import=pywin32 ^
    --hidden-import=pycaw.utils ^
    --hidden-import=pycaw.constants ^
    --hidden-import=PIL.ImageTk ^
    --hidden-import=PIL.ImageFilter ^
    --hidden-import=cairosvg ^
    --hidden-import=cairocffi ^
    --hidden-import=pygame ^
    --add-binary "resource\libcairo-2.dll;resource" ^
    --add-data "resource\_mute.wav;resource" ^
    --add-data "resource\_unmute.wav;resource" ^
    --add-data "config.json;." ^
    --manifest mic_mute_tkinter.exe.manifest ^
    --name mic_mute_tkinter ^
    mic_state_controller_tkinter.py
  ```
- For PyQt6:
  ```bash
  pyinstaller --onefile --windowed ^
    --hidden-import=pycaw ^
    --hidden-import=comtypes ^
    --hidden-import=pywin32 ^
    --hidden-import=pycaw.utils ^
    --hidden-import=pycaw.constants ^
    --hidden-import=PyQt6.QtSvg ^
    --hidden-import=pygame ^
    --add-binary "resource\libcairo-2.dll;resource" ^
    --add-data "resource\_mute.wav;resource" ^
    --add-data "resource\_unmute.wav;resource" ^
    --add-data "config.json;." ^
    --manifest mic_mute_pyqt.exe.manifest ^
    --name mic_mute_pyqt ^
    mic_state_controller_pyqt.py
  ```
- For Windows, use semicolons (`;`) as separators in `--add-binary` and `--add-data`. For Unix-like systems, use colons (`:`).
- Ensure a manifest file (`mic_mute_tkinter.exe.manifest` or `mic_mute_pyqt.exe.manifest`) is created to enforce administrator privileges, as shown in the batch script.

## Troubleshooting
- **`.exe` Fails**:
  - Run as administrator to ensure hotkey and audio control functionality.
  - Ensure the microphone is the default recording device (Windows Sound settings).
  - Verify `libcairo-2.dll` is included in the `resource` folder and bundled with the `.exe`.
- **Settings Not Saved**:
  - Check that `~/.mic_mute_app/config.json` exists and is writable:
    ```bash
    dir %USERPROFILE%\.mic_mute_app
    ```
  - Ensure the executable is run as administrator, as writing to `~/.mic_mute_app` may fail without elevated permissions.
  - Verify that `config.json` is included in the executable (bundled in the root directory via `--add-data "config.json;."`).
  - Check console logs for `Error saving config` or `Error loading config` messages.
- **Icon Issues**:
  - Check console for "SVG converted to PNG, mode: RGBA" (Tkinter) or "Overlay created" (PyQt6).
  - Ensure `libcairo-2.dll` is in the `resource` folder and bundled with the `.exe`.
  - If vector icon fails, a fallback RGBA icon is used (Tkinter only).
- **Hotkey Issues**:
  - Run as administrator for special keys (e.g., `Pause`).
  - Reset hotkey via GUI if conflicts occur.
  - For PyQt6, ensure hotkey capture is not interrupted by other applications.
- **Overlay Not Showing**:
  - Check console for "Overlay shown: Muted" or "Overlay hidden: Unmuted".
  - Ensure no conflicting apps block the overlay.
  - Verify overlay position, size, and opacity settings in `~/.mic_mute_app/config.json`.
- **Sound Not Playing**:
  - Ensure bundled WAV files (`_mute.wav`, `_unmute.wav`) are accessible in the executable's temporary directory or custom WAV files are valid and accessible.
  - Check console for sound loading errors (e.g., "Error loading mute sound").
  - Reapply sounds via GUI or select new WAV files.
  - If no valid sound is loaded, the app falls back to a default beep (Tkinter) or skips playback (PyQt6).
- **PyQt6-Specific Issues**:
  - If the GUI is unresponsive during hotkey capture, ensure no other applications are intercepting keyboard input.
  - Check for "Toggle request ignored" logs to diagnose rapid toggle issues.

## Contributing
Fork the repository, make changes, and submit a pull request. Issues and feature requests are welcome!