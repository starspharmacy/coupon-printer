@echo off

echo Checking if Python is installed...

where python >nul 2>&1

if %errorlevel%==1 (
    echo Python not found, downloading and installing...
    curl https://www.python.org/ftp/python/3.10.0/python-3.10.0-amd64.exe --output python-3.10.0-amd64.exe
    start /wait python-3.10.0-amd64.exe /quiet InstallAllUsers=1 PrependPath=1
    del python-3.10.0-amd64.exe
) else (
    echo Python found, continuing with installation...
)

echo Installing required packages...
python -m pip install -r requirements.txt

echo Running script...
python script.py

pause
