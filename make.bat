@echo off

where pyinstaller >nul 2>&1
if %ERRORLEVEL% neq 0 (
    python -m pip install pyinstaller
)

if exist .\build (
    rd /s /q .\build
)

xcopy /E /I /H /Y .\source_files .\build

cd .\build

python -m PyInstaller main.py
xcopy /E /I /H /Y  .\assets .\dist\assets
xcopy /E /I /H /Y  .\assets .\dist\assets
python -m PyInstaller --name WolfWire --onefile --windowed --icon=.\icon.ico main.py

cd ..

"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" .\setup_compile_script.iss