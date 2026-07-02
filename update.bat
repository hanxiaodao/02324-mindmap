@echo off
cd /d "%~dp0"
pwsh -ExecutionPolicy Bypass -File "%~dp0scripts\update.ps1"
