@echo off
net session >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo This script requires administrative privileges. Requesting elevation...
    powershell -Command "Start-Process cmd -ArgumentList '/c cd /d %CD% && %0 %1' -Verb RunAs"
    exit /b
)

set "BUILD_TYPE=%1"

if "%BUILD_TYPE%"=="" (
    echo Error: Please specify the build type: 'tkinter' or 'pyqt'
    echo Usage: %0 [tkinter^|pyqt]
    pause
    exit /b 1
)

echo Building Microphone Mute Control executable with admin privileges for %BUILD_TYPE%...

if "%BUILD_TYPE%"=="tkinter" (
    set "EXE_NAME=Microphone Mute Control Tkinter"
    set "SCRIPT_NAME=mic_state_controller_tkinter.py"
) else if "%BUILD_TYPE%"=="pyqt" (
    set "EXE_NAME=Microphone Mute Control PyQt"
    set "SCRIPT_NAME=mic_state_controller_pyqt.py"
) else (
    echo Error: Invalid build type '%BUILD_TYPE%'. Use 'tkinter' or 'pyqt'.
    pause
    exit /b 1
)

set "ICON_FILE=resource\icon.ico"
if exist "%ICON_FILE%" (
    set "ICON_OPTION=--icon=%ICON_FILE%"
) else (
    echo Warning: %ICON_FILE% not found. Building without an icon, which may affect --uac-admin reliability.
    set "ICON_OPTION="
)

if "%BUILD_TYPE%"=="tkinter" (
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
      --add-data "resource\mute_icon.ico;resource" ^
      --add-data "resource\icon.ico;resource" ^
      --add-data "config.json;." ^
      --uac-admin ^
      %ICON_OPTION% ^
      --name "%EXE_NAME%" ^
      %SCRIPT_NAME%
) else (
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
      --add-data "resource\mute_icon.ico;resource" ^
      --add-data "resource\icon.ico;resource" ^
      --add-data "config.json;." ^
      --uac-admin ^
      %ICON_OPTION% ^
      --name "%EXE_NAME%" ^
      %SCRIPT_NAME%
)

if %ERRORLEVEL% neq 0 (
    echo Error: PyInstaller build failed.
    pause
    exit /b 1
)

pause