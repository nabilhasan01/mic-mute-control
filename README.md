# Microphone Mute Control

A lightweight Windows application to toggle microphone mute status with a hotkey, displaying a 48x48 red microphone icon (with white slash) in the top-middle of the screen when muted. Features a system tray icon, GUI for hotkey and overlay customization, external mute detection, and customizable sound feedback with bundled default sounds.

## Features
- **Toggle Mute**: Mute/unmute the default microphone using a customizable hotkey (default: `Ctrl+Alt+M`).
- **Overlay Icon**: Displays a 48x48 vector-based red microphone icon with a white slash (no shadow) in the top-middle of the screen when muted, with 70% transparency for minimal obstruction.
- **System Tray**: Shows a tray icon (red "M" for muted, green "U" for unmuted) with a right-click menu to toggle mute, show GUI, or exit.
- **Hotkey Customization**: Set a new hotkey via the GUI (e.g., `Ctrl+Shift+M`, `Pause`) with real-time capture.
- **Overlay Customization**: Adjust overlay position (e.g., Top Mid, Bottom Right), size (32x32, 48x48, 64x64), and margin (0-50 pixels) via the GUI.
- **External Mute Detection**: Polls every 500ms to detect mute/unmute changes from other sources (e.g., keyboard mute key).
- **Sound Feedback**: Play custom WAV files or bundled default sounds (`_mute.wav` for mute, `_unmute.wav` for unmute) on mute/unmute (configurable via GUI). Falls back to a default beep if custom sounds fail.
- **Start Minimized**: Launches directly to the system tray, with the GUI hidden until requested.
- **Cross-Resolution Support**: Overlay dynamically centers based on screen resolution (e.g., `x=936` for 1920x1080).
- **Configuration Persistence**: Saves overlay settings (position, size, margin, opacity) and sound file paths to `~/.mic_mute_app/config.json` for both the script and executable, ensuring settings persist across restarts.

## Requirements
- Windows 10/11 (64-bit).
- Python 3.11+ (for development; not needed for the `.exe`).
- `libcairo-2.dll` (included in the `resource` folder for development or bundled with the `.exe`).

## Installation (Development)
1. Clone the repository:
   ```bash
   git clone https://github.com/nabilhasan01/mic-mute-control.git
   cd mic-mute-control
   ```
2. Install dependencies:
   ```bash
   pip install pycaw comtypes psutil pywin32 keyboard pystray Pillow pyinstaller cairosvg cairocffi pygame
   ```
3. Ensure `libcairo-2.dll` is in the `resource` folder:
   - Download from [GTK+ for Windows](https://github.com/tschoonj/gtkmm-winbuild/releases) (`bin/libcairo-2.dll`).
   - Or copy from `C:\Program Files\GTK3-Runtime Win64\bin` to `resource\libcairo-2.dll`.
   ```bash
   copy "C:\Program Files\GTK3-Runtime Win64\bin\libcairo-2.dll" resource\
   ```

## Usage
1. Run the script (preferably as administrator for hotkey support):
   ```bash
   python mute_mic_app.py
   ```
   - The app starts minimized to the system tray.
2. **Toggle Mute**:
   - Press `Ctrl+Alt+M` (default hotkey).
   - Or right-click the tray icon and select "Toggle Mute".
   - A 48x48 red mic icon with a white slash appears in the configured position when muted.
   - Default bundled sounds (`_mute.wav` for mute, `_unmute.wav` for unmute) or custom WAV sounds play on mute/unmute transitions.
3. **Customize Hotkey**:
   - Right-click tray icon, select "Show Window".
   - Click "Set Hotkey", then press a key combination (e.g., `Ctrl+Shift+M`).
4. **Customize Overlay**:
   - Open the GUI, adjust "Position" (e.g., Top Mid), "Size" (e.g., 48x48), "Margin" (e.g., 10), or "Opacity" (e.g., 0.7).
   - Changes apply instantly and are saved to `~/.mic_mute_app/config.json`.
5. **Configure Sound**:
   - Open the GUI, click "Browse" to select custom WAV files for mute and unmute sounds, then click "Apply".
   - Default sounds (`_mute.wav`, `_unmute.wav`)58 are used if no custom files are selected.
6. **Exit**:
   - Right-click tray icon, select "Exit".

## Building the Executable
To create a standalone `.exe` with bundled default sound files (`_mute.wav`, `_unmute.wav`), `config.json`, and administrator privileges:

1. Ensure `config.json` exists in the project directory with default settings (e.g., `{"overlay_position": "Top Mid", "overlay_size": "48x48", "overlay_margin": 10, "overlay_opacity": 0.7, "mute_sound_file": "resource\\_mute.wav", "unmute_sound_file": "resource\\_unmute.wav"}`).
2. Run the provided batch script to build the executable:
   ```bash
   build_mic_mute_exe.bat
   ```
   - This script creates a manifest file to ensure the executable requests administrator privileges and runs `PyInstaller` with the necessary parameters.
   - The script includes all required dependencies and resources (`libcairo-2.dll`, `_mute.wav`, `_unmute.wav`, `config.json`).
   - Settings are saved to `~/.mic_mute_app/config.json` to persist across restarts.
   - If `comtypes.gen` is needed (check for `pycaw` errors), modify the batch script to include:
     ```bash
     --add-data "C:\Users\YourUsername\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.11_qbz5n2kfra8p0\LocalCache\local-packages\Python311\site-packages\comtypes\gen;comtypes\gen"
     ```
3. Find `mute_mic_app.exe` in the `dist` folder.
4. Share the `.exe` (run as administrator on other PCs).

### Manual PyInstaller Command (Alternative)
If you prefer to build manually, use the following command:
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
  --manifest mute_mic_app.exe.manifest ^
  mute_mic_app Specialization.py
```
- For Windows, use semicolons (`;`) as separators in `--add-binary` and `--add-data`. For Unix-like systems, use colons (`:`).
- Ensure a manifest file (`mute_mic_app.exe.manifest`) is created to enforce administrator privileges, as shown in the batch script.

## Troubleshooting
- **`.exe` Fails**:
  - Run as administrator to ensure hotkey and audio control functionality.
  - Ensure the microphone is the default recording device (Windows Sound settings).
  - Verify `libcairo-2.dll` is bundled with the `.exe` (included in the `resource` folder).
- **Settings Not Saved**:
  - Check that `~/.mic_mute_app/config.json` exists and is writable:
    ```bash
    dir %USERPROFILE%\.mic_mute_app
    ```
  - Ensure the executable is run as administrator, as writing to `~/.mic_mute_app` may fail without elevated permissions.
  - Verify that `config.json` is included in the executable (bundled in the root directory via `--add-data "config.json;."`).
  - Check console logs for `Error saving config` or `Error loading config` messages.
- **Icon Issues**:
  - Check console for "SVG converted to PNG, mode: RGBA".
  - Ensure `libcairo-2.dll` is in the `resource` folder for development or bundled with the `.exe`.
  - If vector icon fails, a fallback RGBA icon is used.
- **Hotkey Issues**:
  - Run as administrator for special keys (e.g., `Pause`).
  - Reset hotkey via GUI if conflicts occur.
- **Overlay Not Showing**:
  - Check console for "Overlay shown: Muted".
  - Ensure no conflicting apps block the overlay.
  - Verify overlay position, size, and opacity settings in `~/.mic_mute_app/config.json`.
- **Sound Not Playing**:
  - Ensure bundled WAV files (`_mute.wav`, `_unmute.wav`) are accessible in the executable's temporary directory or custom WAV files are valid and accessible.
  - Check console for sound loading errors (e.g., "Error loading mute sound").
  - Reapply sounds via GUI or select new WAV files.
  - If no valid sound is loaded, the app falls back to a default beep.

## Contributing
Fork the repository, make changes, and submit a pull request. Issues and feature requests are welcome!