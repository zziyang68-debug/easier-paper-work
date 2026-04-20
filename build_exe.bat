@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

set "PYINSTALLER_EXE=%USERPROFILE%\AppData\Roaming\Python\Python313\Scripts\pyinstaller.exe"

if not exist "%PYINSTALLER_EXE%" set "PYINSTALLER_EXE=%APPDATA%\Python\Python313\Scripts\pyinstaller.exe"

if exist "%PYINSTALLER_EXE%" goto run_local

where pyinstaller >nul 2>nul
if %errorlevel%==0 (
    pyinstaller --noconfirm --clean --distpath .. --workpath "%SCRIPT_DIR%build" --specpath "%SCRIPT_DIR%" "TextCompareTool.spec"
    goto end
)

python -m PyInstaller --noconfirm --clean --distpath .. --workpath "%SCRIPT_DIR%build" --specpath "%SCRIPT_DIR%" "TextCompareTool.spec"
goto end

:run_local
"%PYINSTALLER_EXE%" --noconfirm --clean --distpath .. --workpath "%SCRIPT_DIR%build" --specpath "%SCRIPT_DIR%" "TextCompareTool.spec"

:end
