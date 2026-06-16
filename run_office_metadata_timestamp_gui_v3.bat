@echo off
mode con: cols=112 lines=32
color A1
title Office_Metadata_Timestamp_GUI_v3
setlocal EnableExtensions EnableDelayedExpansion

echo +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
echo Office Metadata Timestamp GUI v3
echo Copyright 2026 // NGHIEN CUU THUOC // RnD PHARMA PLUS // WWW.NGHIENCUUTHUOC.COM
echo +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
echo.

REM ==========================================================================================
REM Portable Python finder
REM You may manually set PY_EXE here if needed, for example:
REM set "PY_EXE=D:\PharmApp\Python312\python.exe"
REM ==========================================================================================
set "PY_EXE="

for %%V in (312) do (
    for %%D in ("%~dp0Python%%V\python.exe" "%~dp0..\Python%%V\python.exe" "%~dp0..\..\Python%%V\python.exe" "%~dp0..\..\..\Python%%V\python.exe" "%~dp0..\..\..\..\Python%%V\python.exe" "%~dp0..\..\..\..\..\Python%%V\python.exe" "%~dp0..\..\..\..\..\..\Python%%V\python.exe") do (
        if not defined PY_EXE if exist "%%~fD" set "PY_EXE=%%~fD"
    )
)

if not defined PY_EXE (
    echo ERROR: Portable Python was not found.
    echo Checked Python312 in current and parent folders.
    echo.
    pause
    exit /b 1
)

set "SCRIPT=%~dp0office_metadata_timestamp_gui_v3.py"
if not exist "%SCRIPT%" (
    echo ERROR: Python script was not found:
    echo %SCRIPT%
    echo.
    pause
    exit /b 1
)

echo Python:
echo "%PY_EXE%"
echo.
echo Script:
echo "%SCRIPT%"
echo.

REM Optional dependency for old Office files: .doc, .xls, .ppt
REM Modern .docx/.xlsx/.pptx files do not need olefile.
echo Checking optional dependency: olefile
"%PY_EXE%" -c "import olefile" >nul 2>nul
if errorlevel 1 (
    echo.
    echo WARNING: olefile is not installed.
    echo Legacy .doc/.xls/.ppt metadata may not be readable until you install it.
    echo.
    set "INSTALL_OLEFILE="
    set /p INSTALL_OLEFILE=Install olefile now using pip? [y/N]: 
    if /i "!INSTALL_OLEFILE!"=="y" (
        "%PY_EXE%" -m pip install olefile
    )
)

echo.
echo Starting GUI v3...
echo.

"%PY_EXE%" "%SCRIPT%"

echo.
echo GUI was closed.
pause
endlocal
