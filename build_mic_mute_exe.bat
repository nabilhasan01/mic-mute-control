@echo off
set "BUILD_TYPE=%1"

if "%BUILD_TYPE%"=="" (
    echo Error: Please specify the build type: 'tkinter' or 'pyqt'
    echo Usage: %0 [tkinter^|pyqt]
    pause
    exit /b 1
)

echo Building Microphone Mute Control executable with admin privileges for %BUILD_TYPE%...

:: Set executable and manifest names based on build type
if "%BUILD_TYPE%"=="tkinter" (
    set "EXE_NAME=Microphone Mute Control tkinter"
    set "SCRIPT_NAME=mic_state_controller_tkinter.py"
) else if "%BUILD_TYPE%"=="pyqt" (
    set "EXE_NAME=Microphone Mute Control pyQt"
    set "SCRIPT_NAME=mic_state_controller_pyqt.py"
) else (
    echo Error: Invalid build type '%BUILD_TYPE%'. Use 'tkinter' or 'pyqt'.
    pause
    exit /b 1
)

:: Create manifest file
echo ^<^?xml version="1.0" encoding="UTF-8" standalone="yes"^?^> > "%EXE_NAME%.exe.manifest"
echo ^<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0"^> >> "%EXE_NAME%.exe.manifest"
echo   ^<trustInfo xmlns="urn:schemas-microsoft-com:asm.v3"^> >> "%EXE_NAME%.exe.manifest"
echo     ^<security^> >> "%EXE_NAME%.exe.manifest"
echo       ^<requestedPrivileges^> >> "%EXE_NAME%.exe.manifest"
echo         ^<requestedExecutionLevel level="requireAdministrator" uiAccess="false"/^> >> "%EXE_NAME%.exe.manifest"
echo       ^</requestedPrivileges^> >> "%EXE_NAME%.exe.manifest"
echo     ^</security^> >> "%EXE_NAME%.exe.manifest"
echo   ^</trustInfo^> >> "%EXE_NAME%.exe.manifest"
echo ^</assembly^> >> "%EXE_NAME%.exe.manifest"

:: Run PyInstaller
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
      --add-data "config.json;." ^
      --manifest "%EXE_NAME%.exe.manifest" ^
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
      --add-data "config.json;." ^
      --manifest "%EXE_NAME%.exe.manifest" ^
      --name "%EXE_NAME%" ^
      %SCRIPT_NAME%
)

:: Clean up manifest file
del "%EXE_NAME%.exe.manifest"

echo Build complete. Executable is in the 'dist' folder.
pause