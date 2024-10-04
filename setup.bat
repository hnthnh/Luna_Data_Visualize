@echo off
REM Check for Python
python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo Python is not installed. Downloading Python...
    
    REM Download Python (change link if needed)
    set "PYTHON_INSTALLER=https://www.python.org/ftp/python/3.9.13/python-3.9.13-amd64.exe"
    set "INSTALLER_NAME=python_installer.exe"
    
    powershell -Command "Invoke-WebRequest -Uri %PYTHON_INSTALLER% -OutFile %INSTALLER_NAME%"
    
    echo Installing Python...
    start /wait "" "%INSTALLER_NAME%" /quiet InstallAllUsers=1 PrependPath=1
    
    echo Installation completed. Please restart Command Prompt to update PATH.
    exit /b
)

REM Check for pip
pip --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo pip is not installed. Installing pip...
    
    REM Download get-pip.py
    set "GET_PIP_SCRIPT=https://bootstrap.pypa.io/get-pip.py"
    set "SCRIPT_NAME=get-pip.py"
    
    powershell -Command "Invoke-WebRequest -Uri %GET_PIP_SCRIPT% -OutFile %SCRIPT_NAME%"
    
    echo Installing pip...
    python "%SCRIPT_NAME%"
    
    echo pip installation completed.
)

REM Install requirements from requirements.txt
echo Installing requirements from requirements.txt...
pip install -r "%~dp0requirements.txt"
timeout /t 5 > nul

REM Run main.py
echo Running main.py...
python "%~dp0main.py"
pause
