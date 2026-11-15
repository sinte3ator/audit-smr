@echo off
title Audit SMR
cd /d "%~dp0"
echo.
echo  Audit SMR - starting...
echo  Folder: %CD%
echo.
where node >nul 2>nul
if errorlevel 1 (
  echo  ERROR: Node.js not found. Install it from https://nodejs.org
  echo.
  pause
  exit /b 1
)
if not exist "start.mjs" (
  echo  ERROR: file start.mjs not found next to this bat file.
  echo  Put start.mjs into the same folder as START.bat
  echo.
  pause
  exit /b 1
)
node start.mjs
echo.
echo  Server stopped. You can close this window.
pause