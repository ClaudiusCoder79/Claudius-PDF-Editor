@echo off
cd /d "%~dp0"
call build_exe.bat
if not exist "dist\Claudius PDF Editor.exe" (
  echo The program was not built - see the messages above.
  pause
  exit /b 1
)
set ISCC=
for %%P in ("%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" "%ProgramFiles%\Inno Setup 6\ISCC.exe" "%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe") do if exist %%P set ISCC=%%~P
if "%ISCC%"=="" (
  echo.
  echo Inno Setup is needed to make the installer. Download it free from https://jrsoftware.org/isdl.php
  echo Install it, then run build_installer.bat again.
  pause
  exit /b 1
)
"%ISCC%" installer.iss
echo.
echo Done. Your installer is in the "installer" folder: Claudius-PDF-Editor-Setup-3.3.exe
pause
