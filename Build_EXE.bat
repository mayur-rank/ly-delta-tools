@echo off
echo Building ODIN Logger EXE...
venv\Scripts\python.exe -m PyInstaller --onefile --noconsole ly-copy-tread\odin_data_logger.py
echo.
echo Build Complete! Check the 'dist' folder for 'odin_data_logger.exe'
pause
