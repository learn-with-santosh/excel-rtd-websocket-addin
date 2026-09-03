@echo off
setlocal

set "ROOT=%~dp0"

start "Excel RTD Add-in" cmd /k "cd /d "%ROOT%addin" && npm run start"
start "Mock WebSocket Server" cmd /k "cd /d "%ROOT%mock-ws-server" && npm run start"

endlocal
