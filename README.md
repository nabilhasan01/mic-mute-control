# Microphone Mute Control

A lightweight Windows application to toggle microphone mute status with a hotkey, displaying a 48x48 red microphone icon (with white slash) in the top-middle of the screen when muted. Features a system tray icon, GUI for hotkey customization, and external mute detection.

## Features
- **Toggle Mute**: Mute/unmute the default microphone using a customizable hotkey (default: `Ctrl+Alt+M`).
- **Overlay Icon**: Shows a 48x48 vector-based red microphone icon with a white slash (no shadow) in the top-middle of the screen when muted, with transparency for minimal obstruction.
- **System Tray**: Displays a tray icon (red "M" for muted, green "U" for unmuted) with a right-click menu to toggle mute, show GUI, or exit.
- **Hotkey Customization**: Set a new hotkey via the GUI (e.g., `Ctrl+Shift+M`, `Pause`).
- **External Mute Detection**: Detects mute/unmute changes from other sources (e.g., keyboard mute key).
- **Start Minimized**: Launches directly to the system tray, with the GUI hidden until requested.
- **Cross-Resolution Support**: Overlay dynamically centers on any screen resolution (e.g., `x=936` for 1920x1080).

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
3. Place `libcairo-2.dll` in the `mic-mute-control\resource` directory:
   - Download from [GTK+ for Windows](https://github.com/tschoonj/gtkmm-winbuild/releases) (`bin/libcairo-2.dll`).
   - Or copy from `C:\Program Files\GTK3-Runtime Win64\bin`.
   ```bash
   copy "C:\Program Files\GTK3-Runtime Win64\bin\libcairo-2.dll" mic-mute-control\resource\libcairo-2.dll
   ```
4. (Optional) Create a `resource` folder and add a fallback PNG icon (`red-radio-microphone-14641.png`) for errors.

## Usage
1. Run the script (preferably as administrator for hotkey support):
   ```bash
   python mic_mute_app.py
   ```
   - The app starts minimized to the system tray.
2. **Toggle Mute**:
   - Press `Ctrl+Alt+M` (default hotkey).
   - Or right-click the tray icon and select "Toggle Mute".
   - A 48x48 red mic icon appears top-middle when muted.
3. **Customize Hotkey**:
   - Right-click tray icon, select "Show Window".
   - Click "Set Hotkey (Press Keys)", then press a key combination (e.g., `Ctrl+Shift+M`).
4. **Exit**:
   - Right-click tray icon, select "Exit".

## Building the Executable
To create a standalone `.exe` for sharing:
1. Run `PyInstaller`:
   ```bash
   pyinstaller --onefile --windowed --hidden-import=pycaw --hidden-import=comtypes --hidden-import=pywin32 --hidden-import=pycaw.utils --hidden-import=pycaw.constants --hidden-import=PIL.ImageTk --hidden-import=PIL.ImageFilter --hidden-import=cairosvg --hidden-import=cairocffi --add-binary "resource\libcairo-2.dll;resource" mic_mute_app.py
   ```
   - If `comtypes.gen` is needed (check for `pycaw` errors):
     ```bash
     pyinstaller --onefile --windowed --hidden-import=pycaw --hidden-import=comtypes --hidden-import=pywin32 --hidden-import=pycaw.utils --hidden-import=pycaw.constants --hidden-import=PIL.ImageTk --hidden-import=PIL.ImageFilter --hidden-import=cairosvg --hidden-import=cairocffi --add-binary "resource\libcairo-2.dll;resource" --add-data "C:\Users\YourUsername\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.11_qbz5n2kfra8p0\LocalCache\local-packages\Python311\site-packages\comtypes\gen;comtypes\gen" mic_mute_app.py
     ```
2. Find `mic_mute_app.exe` in the `dist` folder.
3. Share the `.exe` (run as administrator on other PCs).

## Troubleshooting
- **`.exe` Fails**:
  - Run as administrator.
  - Ensure the microphone is the default recording device (Windows Sound settings).
  - Check for missing `libcairo-2.dll` in the `resource` folder with the `.exe`.
- **Icon Looks Bad**:
  - Verify console output: "SVG converted to PNG, mode: RGBA".
  - Ensure `libcairo-2.dll` is in the `mic-mute-control\resource` directory.
- **Hotkey Issues**:
  - Run as administrator for special keys (e.g., `Pause`).
  - Set a new hotkey via the GUI.
- **Overlay Not Showing**:
  - Check console for "Overlay shown: Muted".
  - Ensure no conflicting apps block the overlay.

## Contributing
Fork the repository, make changes, and submit a pull request. Issues and feature requests are welcome!