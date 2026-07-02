@echo off
cd /d "%~dp0"
pwsh -ExecutionPolicy Bypass -File "%~dp0update.ps1"
