@echo off
setlocal enabledelayedexpansion

REM ============================================================================
REM Portable KD Assistant Taskbar Shortcut Creator
REM Run this ONCE to create and pin KD Assistant to your taskbar
REM ============================================================================

cls
echo.
echo ============================================================================
echo Creating KD Assistant Taskbar Shortcut
echo ============================================================================
echo.

REM Get the directory where THIS script is located (portable!)
set "SCRIPT_DIR=%~dp0"
set "PYTHON_EXE=python.exe"
set "KD_ASSISTANT_PY=!SCRIPT_DIR!KD Assistant.py"
set "ICON_PATH=!SCRIPT_DIR!_internals\data\templates\bismillah.ico"
set "SHORTCUT_NAME=KD Assistant"

REM Verify files exist
if not exist "!KD_ASSISTANT_PY!" (
    echo ERROR: KD Assistant.py not found at:
    echo !KD_ASSISTANT_PY!
    echo.
    pause
    exit /b 1
)

if not exist "!ICON_PATH!" (
    echo WARNING: Custom icon not found at:
    echo !ICON_PATH!
    echo Proceeding without custom icon...
    set "ICON_PATH="
)

echo Script directory: !SCRIPT_DIR!
echo KD Assistant.py: !KD_ASSISTANT_PY!
if not "!ICON_PATH!"=="" (
    echo Icon: !ICON_PATH!
)
echo.

REM Step 1: Verify Python is available
echo Step 1: Checking Python installation...
python --version >nul 2>&1
if !ERRORLEVEL! EQU 0 (
    echo OK - Python found in PATH
) else (
    echo ERROR: Python not found. Please install Python first.
    pause
    exit /b 1
)

REM Step 2: Get Start Menu path (user-aware, portable)
echo Step 2: Locating Start Menu folder...
set "START_MENU=%APPDATA%\Microsoft\Windows\Start Menu\Programs"

if not exist "!START_MENU!" (
    echo ERROR: Start Menu path not found:
    echo !START_MENU!
    pause
    exit /b 1
)
echo OK - Found: !START_MENU!
echo.

REM Step 3: Create shortcut using VBScript
echo Step 3: Creating shortcut...

set "SHORTCUT_PATH=!START_MENU!\!SHORTCUT_NAME!.lnk"
set "VBSCRIPT=!SCRIPT_DIR!make_shortcut_temp.vbs"

REM Build VBScript with icon support
(
    echo Set oWS = WScript.CreateObject("WScript.Shell"^)
    echo sLinkFile = "!SHORTCUT_PATH!"
    echo Set oLink = oWS.CreateShortcut(sLinkFile^)
    echo oLink.TargetPath = "!PYTHON_EXE!"
    echo oLink.Arguments = """!KD_ASSISTANT_PY!"""
    echo oLink.WorkingDirectory = "!SCRIPT_DIR!"
    echo oLink.Description = "Launch G5 Automation Tools - KD Assistant"
    if not "!ICON_PATH!"=="" (
        echo oLink.IconLocation = "!ICON_PATH!"
    )
    echo oLink.Save
    echo WScript.Echo "Shortcut created successfully"
) > "!VBSCRIPT!"

REM Execute VBScript
cscript "!VBSCRIPT!" >nul 2>&1
set "VBS_RESULT=!ERRORLEVEL!"

REM Clean up VBScript
del "!VBSCRIPT!" 2>nul

if not !VBS_RESULT! EQU 0 (
    echo ERROR: Failed to create shortcut
    pause
    exit /b 1
)

if exist "!SHORTCUT_PATH!" (
    echo OK - Shortcut created in Start Menu
) else (
    echo ERROR: Shortcut file was not created
    pause
    exit /b 1
)

echo.

REM Step 4: Pin to taskbar via PowerShell
echo Step 4: Pinning to taskbar...

REM Use PowerShell to pin the shortcut to taskbar
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { $shell = New-Object -ComObject Shell.Application; $folder = $shell.Namespace('!START_MENU!'); $item = $folder.ParseName('KD Assistant.lnk'); $verbs = $item.Verbs(); $pinVerb = $verbs | Where-Object {$_.Name -like '*Pin*' -or $_.Name -like '*pin*'} | Select-Object -First 1; if ($null -ne $pinVerb) { $pinVerb.DoIt(); Write-Host 'Pinned to taskbar'; exit 0 } else { Write-Host 'WARNING: Pin verb not found - may be unavailable on this system'; exit 1 } } catch { Write-Host ('Error: ' + $_); exit 1 }"

set "PIN_RESULT=!ERRORLEVEL!"

echo.
echo ============================================================================
echo Setup Results
echo ============================================================================
echo.
echo Shortcut created at:
echo   !SHORTCUT_PATH!
echo.

if !PIN_RESULT! EQU 0 (
    echo Status: SUCCESS
    echo Your shortcut has been created and pinned to the taskbar!
    echo.
    echo You can now launch KD Assistant directly from your taskbar.
) else (
    echo Status: PARTIAL SUCCESS
    echo The shortcut was created successfully in your Start Menu.
    echo.
    echo To complete pinning to taskbar:
    echo   1. Press Windows key to open Start Menu
    echo   2. Search for "KD Assistant"
    echo   3. Right-click it
    echo   4. Select "Pin to taskbar"
)

echo.
echo Press any key to continue...
pause >nul

endlocal
exit /b 0
