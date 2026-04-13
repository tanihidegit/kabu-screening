@echo off
title Japan Stock Selector
set SCRIPT_DIR=%~dp0
set SCRIPT_PATH=%~dp0japan_stock_realtime.py
set PYTHON_EXE=
echo.
echo  === Japan Stock Selector ===
echo.
echo  Script folder: %SCRIPT_DIR%
echo.
if not exist "%SCRIPT_PATH%" (
    echo ERROR: japan_stock_realtime.py not found in same folder.
    pause
    exit /b 1
)
echo  Script found: OK
echo.
if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" ( set PYTHON_EXE="%LOCALAPPDATA%\Programs\Python\Python313\python.exe" & goto RUN )
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" ( set PYTHON_EXE="%LOCALAPPDATA%\Programs\Python\Python312\python.exe" & goto RUN )
if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" ( set PYTHON_EXE="%LOCALAPPDATA%\Programs\Python\Python311\python.exe" & goto RUN )
if exist "C:\Python313\python.exe" ( set PYTHON_EXE="C:\Python313\python.exe" & goto RUN )
where py >nul 2>&1 && ( set PYTHON_EXE=py & goto RUN )
where python >nul 2>&1 && ( set PYTHON_EXE=python & goto RUN )
echo ERROR: Python not found. Install from https://www.python.org
pause
exit /b 1
:RUN
echo  Python: %PYTHON_EXE%
echo.
echo  Installing libraries...
%PYTHON_EXE% -m pip install yfinance openpyxl tqdm --quiet
echo.
echo  Starting... (5-10 min)
echo.
cd /d "%SCRIPT_DIR%"
%PYTHON_EXE% "%SCRIPT_PATH%"
echo.
echo  === Done! Excel file saved in: %SCRIPT_DIR% ===
echo.
pause
