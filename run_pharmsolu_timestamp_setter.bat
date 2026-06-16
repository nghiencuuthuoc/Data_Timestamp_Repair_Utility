@echo off
mode con: cols=110 lines=35
color A1
title PharmSolu Timestamp Setter
setlocal EnableExtensions EnableDelayedExpansion

echo +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
echo PharmSolu Timestamp Setter
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

set "SCRIPT=%~dp0pharmsolu_timestamp_setter.py"
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

REM ==========================================================================================
REM User inputs
REM ==========================================================================================
set "ROOT_PATH="
set /p ROOT_PATH=Enter root folder [D:\PharmSolu]: 
if not defined ROOT_PATH set "ROOT_PATH=D:\PharmSolu"

echo.
echo Choose timestamp mode:
echo   1 - Manual year; preserve existing month/day/time
echo       Example: 2029-05-14 14:24:00 becomes 2015-05-14 14:24:00
echo   2 - Exact date/time for every selected item
echo       Example: 2015-01-01 00:00:00
echo   3 - Detect year from folder/file path; optional fallback year
echo       Example path contains: CHINESE PHARMACOPOEIA 2015 CMSP
echo.

set "MODE_CHOICE="
set /p MODE_CHOICE=Enter choice [1]: 
if not defined MODE_CHOICE set "MODE_CHOICE=1"

set "MODE=year"
set "MODE_EXTRA="

if "%MODE_CHOICE%"=="1" (
    set "MODE=year"
    set "TARGET_YEAR="
    set /p TARGET_YEAR=Enter target year, for example 2015: 
    if not defined TARGET_YEAR (
        echo ERROR: Target year is required.
        pause
        exit /b 1
    )
    set "MODE_EXTRA=--year !TARGET_YEAR!"
)

if "%MODE_CHOICE%"=="2" (
    set "MODE=exact"
    set "TARGET_DT="
    set /p TARGET_DT=Enter exact date/time [2015-01-01 00:00:00]: 
    if not defined TARGET_DT set "TARGET_DT=2015-01-01 00:00:00"
    set "MODE_EXTRA=--datetime "!TARGET_DT!""
)

if "%MODE_CHOICE%"=="3" (
    set "MODE=path-year"
    set "FALLBACK_YEAR="
    set /p FALLBACK_YEAR=Fallback year if no year is found in path [press Enter to skip]: 
    if defined FALLBACK_YEAR set "MODE_EXTRA=--year !FALLBACK_YEAR!"
)

echo.
set "SCOPE_ARG=--suspicious-only"
set "SUSPICIOUS_ONLY="
set /p SUSPICIOUS_ONLY=Process suspicious timestamps only? [Y/n]: 
if /i "%SUSPICIOUS_ONLY%"=="n" set "SCOPE_ARG=--all-items"

set "DIR_ARG="
set "INCLUDE_DIRS="
set /p INCLUDE_DIRS=Also process folder timestamps? [y/N]: 
if /i "%INCLUDE_DIRS%"=="y" set "DIR_ARG=--include-dirs"

set "LIMIT_ARG="
set "LIMIT_VAL="
set /p LIMIT_VAL=Limit selected items, 0 means no limit [0]: 
if not defined LIMIT_VAL set "LIMIT_VAL=0"
if not "%LIMIT_VAL%"=="0" set "LIMIT_ARG=--limit %LIMIT_VAL%"

set "APPLY_ARG="
set "APPLY_NOW="
echo.
echo IMPORTANT: Default is dry-run. No timestamps are changed unless you type Y below.
set /p APPLY_NOW=Apply changes now? [y/N]: 
if /i "%APPLY_NOW%"=="y" set "APPLY_ARG=--apply"

echo.
echo ==========================================================================================
echo Running timestamp scan...
echo ==========================================================================================
echo.

"%PY_EXE%" "%SCRIPT%" --root "%ROOT_PATH%" --mode "%MODE%" %MODE_EXTRA% %SCOPE_ARG% %DIR_ARG% %LIMIT_ARG% %APPLY_ARG% --log-dir "%~dp0timestamp_setter_logs"

echo.
echo ==========================================================================================
echo Done. Check the CSV log folder:
echo "%~dp0timestamp_setter_logs"
echo ==========================================================================================
echo.
pause
endlocal
