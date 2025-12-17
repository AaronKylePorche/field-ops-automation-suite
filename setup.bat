@echo off
chcp 65001 >nul
REM ============================================================================
REM KD Assistant - Automated Setup Script
REM ============================================================================
REM This script sets up Python, pip, and all dependencies automatically

setlocal enabledelayedexpansion
cd /d "%~dp0"

REM Prevent Python from creating __pycache__ folders (keeps root directory clean)
set PYTHONDONTWRITEBYTECODE=1

REM Colors
color 0A
cls

REM Setup logging
set LOG_FILE=_internals\setup_log.txt
(
    echo Setup started at %date% %time%
    echo.
) > "!LOG_FILE!"

echo.
echo ============================================================================
echo  KD Assistant - Setup Wizard
echo ============================================================================
echo.
echo This script will:
echo   1. Check for Python
echo   2. Check for pip
echo   3. Install Python dependencies
echo   4. Configure your settings
echo   5. Verify Python installation
echo   6. Launch the application
echo.
echo Log file: _internals\setup_log.txt
echo.
echo ============================================================================
echo.
echo Press any key to continue...
pause >nul
echo.

REM ============================================================================
REM Check if Python is installed
REM ============================================================================

echo [Step 1/6] Checking for Python...
(echo [Step 1/6] Checking for Python...) >> "!LOG_FILE!"

set PYTHON_FOUND=0

REM First check: is python in PATH?
python --version >nul 2>&1
if !ERRORLEVEL! EQU 0 (
    set PYTHON_FOUND=1
    echo OK - Python is installed and in PATH
    for /f "tokens=*" %%i in ('python --version 2^>^&1') do (
        set PYTHON_VERSION=%%i
        echo !PYTHON_VERSION! >> "!LOG_FILE!"
    )
    echo !PYTHON_VERSION!
    goto after_install
) else (
    REM Second check: Search common installation paths
    (echo Python not in PATH - searching common installation paths) >> "!LOG_FILE!"
    echo Searching for Python in common locations...

    REM Check each common path
    if exist "C:\Python312\python.exe" (
        (echo Found Python at C:\Python312\) >> "!LOG_FILE!"
        set "PATH=C:\Python312;C:\Python312\Scripts;!PATH!"
        set PYTHON_FOUND=1
    ) else (
        if exist "C:\Program Files\Python312\python.exe" (
            (echo Found Python at C:\Program Files\Python312\) >> "!LOG_FILE!"
            set "PATH=C:\Program Files\Python312;C:\Program Files\Python312\Scripts;!PATH!"
            set PYTHON_FOUND=1
        ) else (
            if exist "C:\Progra~2\Python312\python.exe" (
                (echo Found Python at C:\Progra~2\Python312\) >> "!LOG_FILE!"
                set "PATH=C:\Progra~2\Python312;C:\Progra~2\Python312\Scripts;!PATH!"
                set PYTHON_FOUND=1
            ) else (
                if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
                    (echo Found Python at %%LOCALAPPDATA%%\Programs\Python\Python312\) >> "!LOG_FILE!"
                    set "PATH=!LOCALAPPDATA!\Programs\Python\Python312;!PATH!"
                    set "PATH=!LOCALAPPDATA!\Programs\Python\Python312\Scripts;!PATH!"
                    set PYTHON_FOUND=1
                )
            )
        )
    )

    REM If found, test if python command now works
    if !PYTHON_FOUND! EQU 1 (
        python --version >nul 2>&1
        if !ERRORLEVEL! EQU 0 (
            echo Found Python! Adding to PATH...
            (echo Python found and added to current PATH) >> "!LOG_FILE!"
            for /f "tokens=*" %%i in ('python --version 2^>^&1') do (
                set PYTHON_VERSION=%%i
            )
            echo !PYTHON_VERSION!
            goto after_install
        ) else (
            set PYTHON_FOUND=0
            (echo Python path found but command failed - will reinstall) >> "!LOG_FILE!"
        )
    )
)

if !PYTHON_FOUND! EQU 0 (
    REM Safety recheck to avoid installing when Python already works
    python --version >nul 2>&1
    if !ERRORLEVEL! EQU 0 (
        echo Python detected after recheck; skipping installation.
    ) else (
        (echo WARNING: Python not found - attempting automatic installation) >> "!LOG_FILE!"
        echo.
        echo Python is not installed. Attempting automatic installation...
        echo.

    REM Try Method 1: Windows Package Manager (winget) - visible output
    winget --version >nul 2>&1
    if !ERRORLEVEL! EQU 0 (
        echo Installing Python 3.12 via Windows Package Manager...
        (echo Attempting installation via winget) >> "!LOG_FILE!"
        winget install --id Python.Python.3.12 -e --accept-package-agreements --accept-source-agreements
        if !ERRORLEVEL! EQU 0 (
            (echo Python installed successfully via winget) >> "!LOG_FILE!"
            echo.
            echo Python installed successfully via Windows Package Manager
            echo.
            echo Closing and reopening command window to refresh PATH...
            (echo Python installed via winget - restarting CMD to refresh PATH) >> "!LOG_FILE!"
            timeout /t 2 /nobreak >nul 2>&1

            REM Relaunch setup.bat in new CMD window with refreshed PATH
            start cmd /k "cd /d %~dp0 && call setup.bat"
            exit /b 0
        )
    )

    REM Method 2: Direct download and install - visible installer
    echo Downloading Python 3.12 installer from python.org...
    (echo Attempting direct download from python.org) >> "!LOG_FILE!"

    set PYTHON_INSTALLER=%TEMP%\python-3.12-installer.exe
    set PYTHON_DOWNLOAD_URL=https://www.python.org/ftp/python/3.12.0/python-3.12.0-amd64.exe

    echo Downloading... (this may take 1-2 minutes)
    powershell -Command "& {[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object Net.WebClient).DownloadFile('!PYTHON_DOWNLOAD_URL!', '!PYTHON_INSTALLER!')}" 2>nul

    if exist "!PYTHON_INSTALLER!" (
        echo.
        echo Running Python 3.12 installer...
        echo (The installation window will appear. Please follow the prompts.)
        echo (Make sure to check "Add Python to PATH" during installation.)
        echo.
        (echo Python installer downloaded, running installer GUI) >> "!LOG_FILE!"

        REM Run installer with visible GUI, PrependPath adds Python to PATH automatically
        "!PYTHON_INSTALLER!" InstallAllUsers=1 PrependPath=1 Include_test=0

        if !ERRORLEVEL! EQU 0 (
            (echo Python installed successfully via direct installer) >> "!LOG_FILE!"
            echo.
            echo Python installed successfully
            echo.
            echo Cleaning up installer...
            del "!PYTHON_INSTALLER!" 2>nul

            echo Closing and reopening command window to refresh PATH...
            (echo Python installed via direct download - restarting CMD to refresh PATH) >> "!LOG_FILE!"
            timeout /t 2 /nobreak >nul 2>&1

            REM Relaunch setup.bat in new CMD window with refreshed PATH
            start cmd /k "cd /d %~dp0 && call setup.bat"
            exit /b 0
        ) else (
            (echo Python installation failed or was cancelled - exit code: !ERRORLEVEL!) >> "!LOG_FILE!"
            echo.
            echo Python installation did not complete successfully.
            del "!PYTHON_INSTALLER!" 2>nul
            goto python_install_failed
        )
    ) else (
        (echo Failed to download Python installer) >> "!LOG_FILE!"
        echo Failed to download Python installer from python.org
        goto python_install_failed
    )
    )
)

:after_install
echo.

REM ============================================================================
REM Check if pip is available
REM ============================================================================

echo [Step 2/6] Checking for pip...
(echo [Step 2/6] Checking for pip...) >> "!LOG_FILE!"

python -m pip --version >nul 2>&1
if !ERRORLEVEL! EQU 0 (
    echo OK - pip is available
    for /f "tokens=*" %%i in ('python -m pip --version 2^>^&1') do (
        set PIP_VERSION=%%i
        echo !PIP_VERSION! >> "!LOG_FILE!"
    )
    echo !PIP_VERSION!
) else (
    (echo pip not found - attempting ensurepip) >> "!LOG_FILE!"
    echo pip not found - attempting to bootstrap with ensurepip...
    python -m ensurepip --upgrade >nul 2>&1
    python -m pip --version >nul 2>&1
    if !ERRORLEVEL! EQU 0 (
        echo OK - pip bootstrapped successfully
        for /f "tokens=*" %%i in ('python -m pip --version 2^>^&1') do (
            set PIP_VERSION=%%i
            echo !PIP_VERSION! >> "!LOG_FILE!"
        )
        echo !PIP_VERSION!
    ) else (
        (echo ERROR: pip not available after ensurepip) >> "!LOG_FILE!"
        echo.
        echo ERROR: pip is not available
        echo This usually means Python was not installed correctly.
        echo.
        pause
        exit /b 1
    )
)

echo.

REM ============================================================================
REM Install dependencies
REM ============================================================================

echo [Step 3/6] Installing dependencies...
echo This may take 2-5 minutes...
echo.
(echo [Step 3/6] Installing dependencies...) >> "!LOG_FILE!"

if exist "_internals\requirements.txt" (
    python -m pip install -r _internals\requirements.txt

    if !ERRORLEVEL! EQU 0 (
        (echo All dependencies installed successfully) >> "!LOG_FILE!"
        echo.
        echo OK - All dependencies installed
    ) else (
        (echo WARNING: Some dependencies may have failed) >> "!LOG_FILE!"
        echo.
        echo WARNING: Some dependencies may have failed to install
        echo But attempting to launch anyway...
    )
) else (
    (echo ERROR: requirements.txt not found) >> "!LOG_FILE!"
    echo.
    echo ERROR: requirements.txt not found
    echo Expected location: _internals\requirements.txt
    echo.
    pause
    exit /b 1
)

echo.

REM ============================================================================
REM Step 4: Configure Settings
REM ============================================================================

echo [Step 4/6] Configuring settings...
(echo [Step 4/6] Configuring settings...) >> "!LOG_FILE!"

set CONFIG_FILE=_internals\config\config.py
(echo Config file path: !CONFIG_FILE!) >> "!LOG_FILE!"
set SKIP_WIZARD=0

if exist "!CONFIG_FILE!" (
    (echo Config file EXISTS - checking reconfigure choice) >> "!LOG_FILE!"
    echo.
    echo Configuration file already exists.
    echo.
    echo Do you want to reconfigure your settings? (Y/N)
    set /p RECONFIGURE="Enter choice (Y/N): "

    (echo User response: !RECONFIGURE!) >> "!LOG_FILE!"

    if /i "!RECONFIGURE!"=="Y" (
        (echo User chose to reconfigure - launching ConfigEditor) >> "!LOG_FILE!"
        python "_internals\scripts\core\ConfigEditor.py"
        (echo ConfigEditor returned with code: !ERRORLEVEL!) >> "!LOG_FILE!"
        REM Wait for file to be flushed to disk before proceeding
        timeout /t 3 /nobreak >nul 2>&1
        set SKIP_WIZARD=1
    ) else (
        (echo User skipped reconfiguration) >> "!LOG_FILE!"
        echo Skipping reconfiguration
        set SKIP_WIZARD=1
    )
)

if !SKIP_WIZARD! EQU 0 (
    if not exist "!CONFIG_FILE!" (
        (echo Config file NOT FOUND - running setup wizard) >> "!LOG_FILE!"
        echo.
        echo Running Configuration Wizard...
        echo.
        python "_internals\scripts\core\ConfigEditor.py"
        (echo ConfigEditor returned with code: !ERRORLEVEL!) >> "!LOG_FILE!"
        REM Wait for file to be flushed to disk before proceeding
        timeout /t 3 /nobreak >nul 2>&1

        if !ERRORLEVEL! NEQ 0 (
            (echo ERROR: ConfigEditor failed) >> "!LOG_FILE!"
            echo.
            echo ERROR: Configuration wizard failed
            echo.
            pause
            exit /b 1
        )
    )
)

echo.

REM ============================================================================
REM Step 5: Verify Python Installation
REM ============================================================================

echo [Step 5/6] Verifying Python installation...
(echo [Step 5/6] Verifying Python installation...) >> "!LOG_FILE!"
(echo Reached verification step) >> "!LOG_FILE!"

for /f "tokens=*" %%i in ('python -c "import sys; print(sys.executable)" 2^>^&1') do set PYTHON_EXE=%%i

REM Extract directory from path
set PYTHON_DIR=!PYTHON_EXE:\python.exe=!

(echo Python location: !PYTHON_DIR!) >> "!LOG_FILE!"
echo Python location: !PYTHON_DIR!

REM PATH mutation removed (no longer necessary)
REM Users can run: python "KD Assistant.py"
REM Or double-click "KD Assistant.py" from File Explorer
(echo PATH mutation skipped - not required for launcher execution) >> "!LOG_FILE!"

echo.
echo.

REM ============================================================================
REM Summary and Launch
REM ============================================================================

echo ============================================================================
echo Setup Complete!
echo ============================================================================
echo.
echo All systems ready:
echo   - Python configured
echo   - Dependencies installed
echo   - Ready to launch
echo.
echo Log file: _internals\setup_log.txt
echo.
(echo Setup completed successfully) >> "!LOG_FILE!"

timeout /t 2 /nobreak

echo.
echo Press any key to launch KD Assistant...
pause >nul

echo [Step 6/6] Launching KD Assistant...
python "KD Assistant.py"

goto :post_launch

:python_install_failed
echo.
echo ERROR: Could not install Python automatically.
echo.
echo Please install Python manually from:
echo   https://www.python.org/downloads/
echo.
echo During installation, make sure to check "Add Python to PATH"
echo.
(echo Setup failed - user needs to install Python manually) >> "!LOG_FILE!"
pause
exit /b 1

:post_launch
endlocal
exit /b 0
