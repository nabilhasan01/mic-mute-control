@echo off
echo Building Microphone Mute Control executable with admin privileges...

:: Create manifest file
echo ^<^?xml version="1.0" encoding="UTF-8" standalone="yes"^?^> > mute_mic_app.exe.manifest
echo ^<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0"^> >> mute_mic_app.exe.manifest
echo   ^<trustInfo xmlns="urn:schemas-microsoft-com:asm.v3"^> >> mute_mic_app.exe.manifest
echo     ^<security^> >> mute_mic_app.exe.manifest
echo       ^<requestedPrivileges^> >> mute_mic_app.exe.manifest
echo         ^<requestedExecutionLevel level="requireAdministrator" uiAccess="false"/^> >> mute_mic_app.exe.manifest
echo       ^</requestedPrivileges^> >> mute_mic_app.exe.manifest
echo     ^</security^> >> mute_mic_app.exe.manifest
echo   ^</trustInfo^> >> mute_mic_app.exe.manifest
echo ^</assembly^> >> mute_mic_app.exe.manifest

:: Run PyInstaller
pyinstaller --onefile --windowed ^
  --hidden-import=pycaw ^
  --hidden-import=comtypes ^
  --hidden-import=pywin32 ^
  --hidden-import=pycaw.utils ^
  --hidden-import=pycaw.constants ^
  --hidden-import=PyQt6 ^
  --hidden-import=pygame ^
  --add-binary "resource\libcairo-2.dll;resource" ^
  --add-data "resource\_mute.wav;resource" ^
  --add-data "resource\_unmute.wav;resource" ^
  --add-data "config.json;." ^
  --manifest mute_mic_app.exe.manifest ^
  mute_mic_app.py

:: Clean up manifest file
del mute_mic_app.exe.manifest

echo Build complete. Executable is in the 'dist' folder.
pause