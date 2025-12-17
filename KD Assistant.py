"""
Bismillah - KD Assistant Launcher
==================================
Main application launcher that reads configuration from config/config.py

This launcher:
1. Reads all script definitions from config.py
2. Displays a menu of enabled scripts
3. Runs selected scripts with proper error handling
4. Launches background services in separate windows
"""

import subprocess
import os
import sys
import time
import io
from pathlib import Path

# Ensure sys is available for typing effect in closing
import sys as _sys

# Set up UTF-8 output for Windows console with line buffering (for immediate rendering)
if sys.platform == 'win32':
    import msvcrt  # noqa: F401
    import ctypes
    try:
        ctypes.windll.kernel32.SetConsoleCP(65001)
        ctypes.windll.kernel32.SetConsoleOutputCP(65001)
    except Exception:
        pass
    try:
        sys.stdout.reconfigure(encoding='utf-8', line_buffering=True)
        sys.stderr.reconfigure(encoding='utf-8', line_buffering=True)
    except Exception:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', line_buffering=True)
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', line_buffering=True)

# Prevent Python from creating __pycache__ folders (keeps root directory clean)
os.environ["PYTHONDONTWRITEBYTECODE"] = "1"

# Add parent directories to path for imports
# Bismillah.py is now in the root directory
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(BASE_DIR, "_internals", "config"))

try:
    import config
except ImportError:
    print("ERROR: Could not find config.py")
    print(f"Expected location: {os.path.join(BASE_DIR, '_internals', 'config', 'config.py')}")
    sys.exit(1)

# Set up console window
os.system('title 🟢 KD Assistant Launcher')
os.system('color 0A')  # Green text on black background

# ============================================================================
# SPLASH SCREEN
# ============================================================================

def show_bismillah_ascii():
    """Display Bismillah ASCII art"""
    bismillah_art = [
        r".---------------------------------------------------------------------.",
        r"|                                                                     |",
        r"|  ██████╗ ██╗███████╗███╗   ███╗██╗██╗     ██╗      █████╗ ██╗  ██╗  |",
        r"|  ██╔══██╗██║██╔════╝████╗ ████║██║██║     ██║     ██╔══██╗██║  ██║  |",
        r"|  ██████╔╝██║███████╗██╔████╔██║██║██║     ██║     ███████║███████║  |",
        r"|  ██╔══██╗██║╚════██║██║╚██╔╝██║██║██║     ██║     ██╔══██║██╔══██║  |",
        r"|  ██████╔╝██║███████║██║ ╚═╝ ██║██║███████╗███████╗██║  ██║██║  ██║  |",
        r"|  ╚═════╝ ╚═╝╚══════╝╚═╝     ╚═╝╚═╝╚══════╝╚══════╝╚═╝  ╚═╝╚═╝  ╚═╝  |",
        r"|                                                                     |",
        r"'---------------------------------------------------------------------'",
    ]

    for line in bismillah_art:
        print(line, flush=True)

    print()
    # Shorter pause to keep transition snappy
    time.sleep(0.3)


def show_binary_quote():
    """Display binary quote character by character (typing effect)"""
    binary_text = """00100010 01000001 01101110 01100100  01110100 01101000 01100101 01101110
01100010 01100101 01101001 01101110 01100111
01100001 01101101 01101111 01101110 01100111
01110100 01101000 01101111 01110011 01100101  01110111 01101000 01101111
01100010 01100101 01101100 01101001 01100101 01110110 01100101 01100100
01100001 01101110 01100100
01100001 01100100 01110110 01101001 01110011 01100101 01100100
01101111 01101110 01100101
01100001 01101110 01101111 01110100 01101000 01100101 01110010  01110100 01101111
01110000 01100001 01110100 01101001 01100101 01101110 01100011 01100101
01100001 01101110 01100100
01100001 01100100 01110110 01101001 01110011 01100101 01100100
01101111 01101110 01100101
01100001 01101110 01101111 01110100 01101000 01100101 01110010  01110100 01101111
01100011 01101111 01101101 01110000 01100001 01110011 01110011 01101001 01101111 01101110
00101110 00100010  00101101  01000001 01101100 00101101 01000001 01110011 01110010
00110001 00110000 00110011 00111010 00110011"""

    print()
    for char in binary_text:
        sys.stdout.write(char)
        sys.stdout.flush()
        time.sleep(0.0008)  # Very fast character by character typing effect
    print()


def show_splash():
    """Display artistic splash screen on startup with ASCII art and binary quote"""
    # Console configured at import for UTF-8 + line buffering

    os.system('cls')
    show_bismillah_ascii()
    time.sleep(0.2)
    show_binary_quote()
    print("\n" * 8)
    time.sleep(1.5)
    os.system('cls')


def show_closing():
    """Display closing/farewell screen when user quits"""
    os.system('cls')
    print("\n" * 7)

    # Top decorative star
    print("✦".center(60))
    time.sleep(0.1)
    print()

    # TYPE OUT THE QUOTE (these lines type character-by-character)
    quote_lines = [
        "So indeed, with hardship comes ease.",
        "Indeed, with hardship comes ease.",
    ]

    for line in quote_lines:
        centered_line = line.center(60)
        for char in centered_line:
            sys.stdout.write(char)
            sys.stdout.flush()
            time.sleep(0.02)
        print()
        time.sleep(0.3)

    # INSTANT LINES (appear one after another)
    print()
    print("~ Surah Ash-Sharh 94:5-6 ~".center(60))
    time.sleep(0.5)  # Pause between verse reference and blessing
    print()
    print("Peace and Blessings be upon you.".center(60))

    # Bottom decorative star
    print()
    sys.stdout.write("✦".center(60))
    sys.stdout.write("\n")
    sys.stdout.flush()

    print("\n" * 7)
    time.sleep(1.5)

    os.system('cls')

# ============================================================================
# MENU DISPLAY
# ============================================================================

def show_menu():
    """Display the main menu of available scripts"""
    os.system('color 0A')
    print("\n" + "="*60)
    print(f"  🟢 KD Assistant")
    print(f"  User: {config.USER_NAME}")
    print("="*60 + "\n")

    # Build menu order (preserve numeric and alphabetic order)
    enabled_scripts = config.get_enabled_scripts()

    # Sort: numbers first (1-9), then 10+, then letters
    def sort_key(item):
        key = item[0]
        if key.isdigit():
            return (0, int(key))
        else:
            return (1, key)

    menu_items = sorted(enabled_scripts.items(), key=sort_key)

    for key, script_info in menu_items:
        desc = script_info.get("name", "Unknown")
        print(f"  [{key}] {desc}")

    print()

# ============================================================================
# SCRIPT EXECUTION
# ============================================================================

def run_script(choice):
    """Execute a script or perform menu action"""
    enabled_scripts = config.get_enabled_scripts()

    if choice not in enabled_scripts:
        print("⚠️  Invalid choice. Try again.")
        return

    script_info = enabled_scripts[choice]
    desc = script_info.get("name", "Script")

    # ===== QUIT OPTION =====
    if choice == "q":
        show_closing()
        return True

    # ===== BACKGROUND SERVICES (like option 10 - Claim Watcher Suite) =====
    if "background_scripts" in script_info:
        print(f"\n🔄 Launching {desc}...")
        for script_path in script_info["background_scripts"]:
            try:
                if os.path.exists(script_path):
                    script_name = os.path.splitext(os.path.basename(script_path))[0]
                    title = f'{script_name} — KD Assistant'
                    # Use /k to keep windows open; need proper quoting for paths with spaces
                    # Format: cmd /k ""python.exe" "script.py""
                    cmd = f'start "{title}" cmd /k ""{sys.executable}" "{script_path}""'
                    subprocess.run(cmd, shell=True)
                    print(f"  ✓ Launched: {script_name}")
                else:
                    print(f"⚠️  Script not found: {script_path}")
            except Exception as e:
                print(f"⚠️  Failed to launch {script_path}: {e}")

        print(f"\n🚀 {desc} launched in {len(script_info['background_scripts'])} separate windows.")
        print("   Windows will stay open so you can monitor their status.")
        return False

    # ===== NORMAL SCRIPTS (run in same window, return to menu) =====
    if "path" in script_info and script_info["path"]:
        script_path = script_info["path"]

        if not os.path.exists(script_path):
            print(f"\n❌ Script not found: {script_path}")
            print("Make sure all scripts are in the correct locations.")
            return False

        print(f"\n🔄 Running: {desc}...\n")
        print("-" * 60)

        try:
            if script_path.lower().endswith(".py"):
                subprocess.run(
                    [sys.executable, script_path],
                    check=True,
                    cwd=BASE_DIR
                )
            elif script_path.lower().endswith(".bat"):
                subprocess.run(
                    script_path,
                    check=True,
                    cwd=BASE_DIR,
                    shell=True
                )
            else:
                subprocess.run(
                    [script_path],
                    check=True,
                    cwd=BASE_DIR
                )

            print("-" * 60)
            print(f"✅ {desc} completed successfully.")

        except subprocess.CalledProcessError as e:
            print("-" * 60)
            print(f"\n❌ Script exited with error code {e.returncode}.")

        except FileNotFoundError:
            print("-" * 60)
            print(f"\n❌ Script not found: {script_path}")

        except Exception as e:
            print("-" * 60)
            print(f"\n❌ Failed to run: {e}")

        return False

    return False

# ============================================================================
# MAIN LOOP
# ============================================================================

def main():
    """Main application loop"""

    # Show splash screen on startup
    show_splash()

    # Validate configuration on startup
    if config.DEBUG:
        config.print_config_info()
        missing = config.validate_paths()
        if missing:
            print("\n⚠️  WARNING: Some script files are missing!")
            for m in missing:
                print(f"  - {m}")
            input("\nPress Enter to continue anyway...")

    while True:
        show_menu()
        choice = input("Select an option: ").strip().lower()

        if run_script(choice):
            # run_script returns True if user wants to quit
            break

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n👋 Launcher interrupted. Exiting.")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        input("Press Enter to exit...")
        sys.exit(1)
