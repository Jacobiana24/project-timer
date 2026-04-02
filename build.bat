@echo off
pyinstaller --onefile --windowed --name "ProjectTimer" --collect-data customtkinter project_timer.py
pause
