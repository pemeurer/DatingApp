@echo off
rem Startet die DatingApp ohne Konsolenfenster.
cd /d "%~dp0"
start "" pythonw "DatingApp.py"
