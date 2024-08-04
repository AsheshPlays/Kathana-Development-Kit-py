@echo off
setlocal

set SCRIPT_NAME=venv_creator.py
set EXE_NAME=venv_creator
set ICON_FILE=asheshicon.ico
set UPX_PATH=B:\pyModules\upx-4.2.4-win64\upx.exe

echo Removing old build directories...
if exist build rmdir /S /Q build
if exist dist rmdir /S /Q dist

echo Creating standalone executable using PyInstaller...
pyinstaller --onefile --windowed --add-data "asheshicon.ico;." --add-data "asheshdevkitbanner.png;." --icon=%ICON_FILE% --name=%EXE_NAME% %SCRIPT_NAME%
if %ERRORLEVEL% neq 0 (
    echo Error: PyInstaller failed to create the executable.
    pause
    exit /b %ERRORLEVEL%
)

if not exist dist\%EXE_NAME%.exe (
    echo Error: The executable file was not created.
    pause
    exit /b 1
)

echo Optimizing executable...
%UPX_PATH% --best --lzma dist\%EXE_NAME%.exe
if %ERRORLEVEL% neq 0 (
    echo Error: UPX failed to optimize the executable.
    pause
    exit /b %ERRORLEVEL%
)

echo Build complete. Executable is located in the 'dist' directory.
pause
