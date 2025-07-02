

# Microphone Mute Control

A lightweight Windows application to toggle microphone mute status with a hotkey, displaying a 48x48 red microphone icon (with white slash) in the top-middle of the screen when muted. Features a system tray icon, GUI for hotkey and overlay customization, and external mute detection.

## Features
- **Toggle Mute**: Mute/unmute the default microphone using a customizable hotkey (default: `Ctrl+Alt+M`).
- **Overlay Icon**: Displays a 48x48 vector-based red microphone icon with a white slash (no shadow) in the top-middle of the screen when muted, with 70% transparency for minimal obstruction.
- **System Tray**: Shows a tray icon (red "M" for muted, green "U" for unmuted) with a right-click menu to toggle mute, show GUI, or exit.
- **Hotkey Customization**: Set a new hotkey via the GUI (e.g., `Ctrl+Shift+M`, `Pause`) with real-time capture.
- **Overlay Customization**: Adjust overlay position (e.g., Top Mid, Bottom Right), size (32x32, 48x48, 64x64), and margin (0-50 pixels) via the GUI.
- **External Mute Detection**: Polls every 500ms to detect mute/unmute changes from other sources (e.g., keyboard mute key).
- **Start Minimized**: Launches directly to the system tray, with the GUI hidden until requested.
- **Cross-Resolution Support**: Overlay dynamically centers based on screen resolution (e.g., `x=936` for 1920x1080).
- **Configuration Persistence**: Saves overlay settings (position, size, margin) to `config.json`.

## Requirements
- Windows 10/11 (64-bit).
- Python 3.11+ (for development; not needed for the `.exe`).
- `libcairo-2.dll` (included in the `.exe` or via GTK+ for development).

## Installation (Development)
1. Clone the repository:
   ```bash
   git clone https://github.com/nabilhasan01/mic-mute-control.git
   cd mic-mute-control
   ```
2. Install dependencies:
   ```bash
   pip install pycaw comtypes psutil pywin32 keyboard pystray Pillow pyinstaller cairosvg cairocffi
   ```
3. Place `libcairo-2.dll` in the project directory (`mic-mute-control`):
   - Download from [GTK+ for Windows](https://github.com/tschoonj/gtkmm-winbuild/releases) (`bin/libcairo-2.dll`).
   - Or copy from `C:\Program Files\GTK3-Runtime Win64\bin`.
   ```bash
   copy "C:\Program Files\GTK3-Runtime Win64\bin\libcairo-2.dll" .
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
3. **Customize Hotkey**:
   - Right-click tray icon, select "Show Window".
   - Click "Set Hotkey", then press a key combination (e.g., `Ctrl+Shift+M`).
4. **Customize Overlay**:
   - Open the GUI, adjust "Position" (e.g., Top Mid), "Size" (e.g., 48x48), or "Margin" (e.g., 10).
   - Changes apply instantly and are saved to `config.json`.
5. **Exit**:
   - Right-click tray icon, select "Exit".

## Building the Executable
To create a standalone `.exe`:
1. Run `PyInstaller`:
   ```bash
   pyinstaller --onefile --windowed --hidden-import=pycaw --hidden-import=comtypes --hidden-import=pywin32 --hidden-import=pycaw.utils --hidden-import=pycaw.constants --hidden-import=PIL.ImageTk --hidden-import=PIL.ImageFilter --hidden-import=cairosvg --hidden-import=cairocffi --add-binary "libcairo-2.dll;." mute_mic_app.py
   ```
   - If `comtypes.gen` is needed (check for `pycaw` errors):
     ```bash
     pyinstaller --onefile --windowed --hidden-import=pycaw --hidden-import=comtypes --hidden-import=pywin32 --hidden-import=pycaw.utils --hidden-import=pycaw.constants --hidden-import=PIL.ImageTk --hidden-import=PIL.ImageFilter --hidden-import=cairosvg --hidden-import=cairocffi --add-binary "libcairo-2.dll;." --add-data "C:\Users\YourUsername\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.11_qbz5n2kfra8p0\LocalCache\local-packages\Python311\site-packages\comtypes\gen;comtypes\gen" mute_mic_app.py
     ```
2. Find `mute_mic_app.exe` in the `dist` folder.
3. Share the `.exe` (run as administrator on other PCs).

## Troubleshooting
- **`.exe` Fails**:
  - Run as administrator.
  - Ensure the microphone is the default recording device (Windows Sound settings).
  - Verify `libcairo-2.dll` is in the `.exe` folder.
- **Icon Issues**:
  - Check console for "SVG converted to PNG, mode: RGBA".
  - Ensure `libcairo-2.dll` is present.
  - If vector icon fails, a fallback RGBA icon is used.
- **Hotkey Issues**:
  - Run as administrator for special keys (e.g., `Pause`).
  - Reset hotkey via GUI if conflicts occur.
- **Overlay Not Showing**:
  - Check console for "Overlay shown: Muted".
  - Ensure no conflicting apps block the overlay.
  - Verify overlay position and size settings.

## Contributing
Fork the repository, make changes, and submit a pull request. Issues and feature requests are welcome!