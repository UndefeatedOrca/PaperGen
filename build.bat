@echo off
setlocal

set "ROOT=%~dp0"
set "PYTHON=%ROOT%.venv-build\Scripts\python.exe"

if not exist "%PYTHON%" set "PYTHON=python"

echo Checking the Python build environment...
"%PYTHON%" -c "import PyInstaller, docx" >nul 2>&1
if errorlevel 1 (
    echo.
    echo Missing build dependencies.
    echo Install them with:
    echo   %ROOT%.venv-build\Scripts\python.exe -m pip install -r "%ROOT%requirements-build.txt"
    echo.
    exit /b 1
)

echo Building PaperGen with PyInstaller...
"%PYTHON%" -m PyInstaller --noconfirm --clean "%ROOT%PaperGen.spec"
if errorlevel 1 (
    echo PyInstaller failed.
    exit /b 1
)

set "ISCC="
if exist "%ProgramFiles(x86)%\Inno Setup 7\ISCC.exe" set "ISCC=%ProgramFiles(x86)%\Inno Setup 7\ISCC.exe"
if not defined ISCC if exist "%ProgramFiles%\Inno Setup 7\ISCC.exe" set "ISCC=%ProgramFiles%\Inno Setup 7\ISCC.exe"
if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not defined ISCC if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if not defined ISCC for /f "delims=" %%I in ('where ISCC.exe 2^>nul') do if not defined ISCC set "ISCC=%%I"

if not defined ISCC (
    echo.
    echo Could not find Inno Setup's ISCC.exe.
    echo Add its folder to PATH or update build.bat with its installation path.
    exit /b 1
)

echo Building the installer with Inno Setup...
"%ISCC%" "%ROOT%installer\PaperGen.iss"
if errorlevel 1 (
    echo Inno Setup failed.
    exit /b 1
)

echo.
echo Build complete:
echo   %ROOT%dist\installer\PaperGenSetup.exe
exit /b 0
