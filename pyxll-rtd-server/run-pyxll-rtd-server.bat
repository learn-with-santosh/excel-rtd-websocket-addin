@echo off
setlocal

set "PROJECT_DIR=%~dp0"
set "ROOT_DIR=%PROJECT_DIR%.."
set "PYTHON=%PROJECT_DIR%.venv\Scripts\python.exe"

cd /d "%PROJECT_DIR%"

if not exist "%PYTHON%" (
    echo Creating Python virtual environment...
    py -3 -m venv .venv
    if errorlevel 1 goto :error
)

"%PYTHON%" -m pip install --upgrade pip
if errorlevel 1 goto :error

"%PYTHON%" -m pip install -e .
if errorlevel 1 goto :error

if not exist "%ROOT_DIR%\mock-ws-server\node_modules" (
    echo Installing mock WebSocket server dependencies...
    npm --prefix "%ROOT_DIR%\mock-ws-server" install
    if errorlevel 1 goto :error
)

start "Mock WebSocket Server" cmd /k "cd /d "%ROOT_DIR%\mock-ws-server" && npm start"

echo.
echo PyXLL RTD package is installed in:
echo   %PROJECT_DIR%.venv
echo.
echo Configure PyXLL with:
echo   %PROJECT_DIR%pyxll.cfg
echo.
echo Use in Excel: =WEBSOCKET_RTD("ACC.NS","LAST")
pause
endlocal
exit /b 0

:error
echo.
echo Setup failed. Review the error above.
pause
endlocal
exit /b 1
