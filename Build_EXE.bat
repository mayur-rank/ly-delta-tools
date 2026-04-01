@echo off
echo Building ODIN GUI Fetcher...
venv\Scripts\python.exe -m PyInstaller --onefile --noconsole --name OdinFetcher ly-copy-tread\odin_gui_logger.py
echo.
echo Build Complete! Check the 'dist' folder for 'OdinFetcher.exe'
pause
