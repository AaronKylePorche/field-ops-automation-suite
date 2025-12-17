"""
Claim Watcher Suite - Unified Console
======================================
Launches all background monitoring services in a single console window.

This unified wrapper manages:
1. Outlook monitoring (detects when Outlook starts/stops)
2. Email_Scanner - Launched when Outlook opens, stopped when Outlook closes
3. Ticket_Reader - Processes tickets from the queue (always running)
4. keep_awake - Prevents system sleep (always running)

All services output to a single window with prefixed messages to identify
which service generated each line.
"""

import subprocess
import sys
import os
import time
import signal
import threading
from pathlib import Path
from datetime import datetime

# Prevent __pycache__ in root
os.environ["PYTHONDONTWRITEBYTECODE"] = "1"

# Get the base directory (3 levels up from this script)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))
MONITORING_DIR = os.path.dirname(os.path.abspath(__file__))
CORE_DIR = os.path.join(os.path.dirname(MONITORING_DIR), "core")

# Try to import WMI (for instant Outlook notifications)
try:
    import win32com.client
    HAS_WMI = True
except ImportError:
    HAS_WMI = False

# ============================================================================
# CONFIGURATION
# ============================================================================

# Services to always run (Ticket_Reader, keep_awake)
ALWAYS_ON_SERVICES = [
    {
        "name": "Ticket_Reader",
        "path": os.path.join(MONITORING_DIR, "Ticket_Reader.py"),
        "prefix": "[Ticket_Reader]"
    },
    {
        "name": "keep_awake",
        "path": os.path.join(MONITORING_DIR, "keep_awake.py"),
        "prefix": "[keep_awake]"
    },
]

# Email_Scanner - conditionally launched based on Outlook state
EMAIL_SCANNER_PATH = os.path.join(CORE_DIR, "Email_Scanner.py")
EMAIL_SCANNER_PREFIX = "[Email_Scanner]"

# Outlook monitoring
OUTLOOK_EXE = "OUTLOOK.EXE"
RESTART_BACKOFF_SEC = 10  # Wait before restarting after crash

# Global state
running_processes = {}
email_scanner_process = None
outlook_is_running = False
shutdown_event = threading.Event()
lock = threading.Lock()

# ============================================================================
# OUTLOOK MONITORING FUNCTIONS
# ============================================================================

def is_outlook_running():
    """Fast check using tasklist (no WMI dependency)."""
    try:
        proc = subprocess.run(
            ["tasklist", "/FI", f"IMAGENAME eq {OUTLOOK_EXE}", "/NH"],
            capture_output=True,
            text=True,
            creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
        )
        txt = (proc.stdout or "").strip()
        return OUTLOOK_EXE.upper() in txt.upper()
    except Exception:
        return False

def wmi_watchers():
    """Create WMI event subscriptions for Outlook start/stop."""
    try:
        svc = win32com.client.GetObject("winmgmts:root\\cimv2")
        startq = "SELECT * FROM Win32_ProcessStartTrace WHERE ProcessName='OUTLOOK.EXE'"
        stopq = "SELECT * FROM Win32_ProcessStopTrace WHERE ProcessName='OUTLOOK.EXE'"
        start_watcher = svc.ExecNotificationQuery(startq)
        stop_watcher = svc.ExecNotificationQuery(stopq)
        return start_watcher, stop_watcher
    except Exception:
        return None, None

def next_event_nonblocking(watcher, timeout_ms=500):
    """Return an event object or None (timeout)."""
    try:
        return watcher.NextEvent(timeout_ms)
    except Exception:
        return None

def outlook_monitor_thread_func():
    """Monitor Outlook state and manage Email_Scanner lifecycle."""
    global email_scanner_process, outlook_is_running

    print("\n" + "-"*70)
    print("[Outlook] Initializing Outlook monitor...")

    # Set up WMI watchers if available
    start_wmi, stop_wmi = None, None
    if HAS_WMI:
        try:
            start_wmi, stop_wmi = wmi_watchers()
            with lock:
                print("[Outlook] WMI event watchers initialized (instant notifications)")
        except Exception as e:
            with lock:
                print(f"[Outlook] WMI not available, falling back to polling: {e}")

    # Initial state check
    outlook_is_running = is_outlook_running()
    with lock:
        print(f"[Outlook] Initial state: {'RUNNING' if outlook_is_running else 'NOT RUNNING'}")
        if outlook_is_running:
            print("[Outlook] Launching Email_Scanner...")

    if outlook_is_running:
        email_scanner_process = launch_email_scanner()

    print("-"*70 + "\n")

    try:
        while not shutdown_event.is_set():
            # Check WMI events (instant, high priority)
            if stop_wmi:
                evt = next_event_nonblocking(stop_wmi, timeout_ms=250)
                if evt is not None:
                    with lock:
                        print("[Outlook] STOP detected (WMI event)")
                    outlook_is_running = False
                    stop_email_scanner()

            if start_wmi:
                evt = next_event_nonblocking(start_wmi, timeout_ms=250)
                if evt is not None:
                    with lock:
                        print("[Outlook] START detected (WMI event)")
                    outlook_is_running = True
                    if email_scanner_process is None:
                        email_scanner_process = launch_email_scanner()

            # Polling safety net (always active, even when WMI available)
            # WMI events can be missed due to timing or queue overflow
            now_up = is_outlook_running()
            if now_up != outlook_is_running:
                outlook_is_running = now_up
                if outlook_is_running:
                    with lock:
                        print("[Outlook] RUNNING detected (poll)")
                    if email_scanner_process is None:
                        email_scanner_process = launch_email_scanner()
                else:
                    with lock:
                        print("[Outlook] NOT RUNNING detected (poll)")
                    stop_email_scanner()

            # Health check: if Email_Scanner crashed while Outlook is up, restart it
            if email_scanner_process is not None:
                code = email_scanner_process.poll()
                if code is not None:
                    with lock:
                        print(f"[Outlook] Email_Scanner exited with code {code}")
                    email_scanner_process = None

                    if outlook_is_running:
                        with lock:
                            print(f"[Outlook] Restarting Email_Scanner in {RESTART_BACKOFF_SEC}s...")
                        time.sleep(RESTART_BACKOFF_SEC)

                        # Check Outlook still up after backoff
                        if is_outlook_running():
                            email_scanner_process = launch_email_scanner()
                        else:
                            with lock:
                                print("[Outlook] Email_Scanner restart skipped - Outlook went down")

            time.sleep(1.0)

    except Exception as e:
        with lock:
            print(f"[Outlook] Monitor error: {e}")
    finally:
        stop_email_scanner()
        with lock:
            print("[Outlook] Monitor stopped")

def launch_email_scanner():
    """Launch Email_Scanner as a subprocess."""
    try:
        if not os.path.exists(EMAIL_SCANNER_PATH):
            with lock:
                print(f"{EMAIL_SCANNER_PREFIX} ERROR: Script not found: {EMAIL_SCANNER_PATH}")
            return None

        # Set UTF-8 encoding for subprocess to handle emojis/unicode
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"

        process = subprocess.Popen(
            [sys.executable, EMAIL_SCANNER_PATH],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            cwd=CORE_DIR,
            text=True,
            encoding='utf-8',
            errors='replace',
            bufsize=1,
            env=env,
            creationflags=subprocess.CREATE_NEW_PROCESS_GROUP if sys.platform == "win32" else 0
        )

        with lock:
            print(f"{EMAIL_SCANNER_PREFIX} Service started (PID: {process.pid})")

        # Start output thread
        output_thread = threading.Thread(
            target=read_output,
            args=(process, "Email_Scanner", EMAIL_SCANNER_PREFIX),
            daemon=True
        )
        output_thread.start()

        return process

    except Exception as e:
        with lock:
            print(f"{EMAIL_SCANNER_PREFIX} ERROR: Failed to launch: {e}")
        return None

def stop_email_scanner():
    """Stop Email_Scanner gracefully."""
    global email_scanner_process

    if email_scanner_process is None:
        return

    try:
        if email_scanner_process.poll() is None:
            with lock:
                print(f"{EMAIL_SCANNER_PREFIX} Stopping...")

            # Terminate the process (SIGTERM is graceful on Windows)
            email_scanner_process.terminate()
            try:
                email_scanner_process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                pass

            # Force kill if still running
            if email_scanner_process.poll() is None:
                with lock:
                    print(f"{EMAIL_SCANNER_PREFIX} Force killing...")
                email_scanner_process.kill()

            with lock:
                print(f"{EMAIL_SCANNER_PREFIX} Stopped")

    except Exception as e:
        with lock:
            print(f"{EMAIL_SCANNER_PREFIX} Error stopping: {e}")
    finally:
        email_scanner_process = None

# ============================================================================
# SERVICE MANAGEMENT
# ============================================================================

def get_timestamp():
    """Return current time in human-readable format: MM/DD/YYYY H:MMAm/Pm"""
    return datetime.now().strftime("%m/%d/%Y %I:%M%p")

def read_output(process, service_name, prefix):
    """Read output from a process and display with prefix.

    Handles output with proper formatting:
    - Each line from subprocess is prefixed with [service_name]
    - Blank lines from subprocess are preserved to maintain structure
    - Double spacing between output blocks for clarity
    """
    try:
        prev_was_blank = False

        while True:
            line = process.stdout.readline()
            if not line:
                break

            line = line.rstrip('\n\r')

            # Handle blank lines from subprocess (preserve for structure)
            if not line:
                # Add blank line only if previous wasn't blank (avoid multiple blanks)
                if not prev_was_blank:
                    with lock:
                        print()  # Single blank line from subprocess output
                        sys.stdout.flush()
                    prev_was_blank = True
            else:
                # Regular output line
                with lock:
                    print(f"[{get_timestamp()}] {prefix} {line}")
                    sys.stdout.flush()
                prev_was_blank = False

        # Add double blank line at end of subprocess output for visual separation
        with lock:
            print()
            print()
            sys.stdout.flush()

    except Exception as e:
        with lock:
            print(f"{prefix} ERROR reading output: {e}")

def launch_service(service):
    """Launch a service as a subprocess."""
    name = service["name"]
    path = service["path"]
    prefix = service["prefix"]

    try:
        if not os.path.exists(path):
            print(f"{prefix} ERROR: Script not found: {path}")
            return None

        # Set UTF-8 encoding for subprocess to handle unicode/emojis
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"

        process = subprocess.Popen(
            [sys.executable, path],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            cwd=BASE_DIR,
            text=True,
            encoding='utf-8',
            errors='replace',
            bufsize=1,
            env=env
        )

        # Start output thread
        output_thread = threading.Thread(
            target=read_output,
            args=(process, name, prefix),
            daemon=True
        )
        output_thread.start()

        with lock:
            print(f"{prefix} Service started (PID: {process.pid})")

        return process

    except Exception as e:
        print(f"{prefix} ERROR: Failed to launch: {e}")
        return None

def shutdown_all_services():
    """Gracefully shut down all services."""
    with lock:
        print("\n" + "-"*70)
        print("⏹️  Shutting down all services...")
        print("-"*70)

    # Signal outlook monitor to stop
    shutdown_event.set()

    # Stop Email_Scanner
    stop_email_scanner()

    # Terminate other services
    for name, info in list(running_processes.items()):
        process = info["process"]
        prefix = info["prefix"]

        try:
            process.terminate()
            try:
                process.wait(timeout=5)
                with lock:
                    print(f"{prefix} Gracefully stopped")
            except subprocess.TimeoutExpired:
                process.kill()
                with lock:
                    print(f"{prefix} Force killed")
        except Exception as e:
            with lock:
                print(f"{prefix} Error stopping: {e}")

    with lock:
        print("-"*70)
        print("✅ All services stopped")

# ============================================================================
# MAIN
# ============================================================================

def main():
    """Main function: launch all services and monitor them."""

    print("\n" + "="*70)
    print("  🟢 CLAIM WATCHER SUITE - UNIFIED CONSOLE")
    print("="*70)
    print("\nLaunching background monitoring services in a single window...")
    print("Each service output is prefixed with [service_name] for identification.\n")
    print("Output Key:")
    print("  [Email_Scanner]   → Detects new claim emails in Outlook")
    print("  [Ticket_Reader]   → Processes claim tickets and runs WOP22 analysis")
    print("  [keep_awake]      → Prevents system sleep")
    print("  [Outlook]         → Monitors Outlook start/stop events\n")
    print("-"*70)

    # Launch always-on services
    for service in ALWAYS_ON_SERVICES:
        process = launch_service(service)
        if process:
            running_processes[service["name"]] = {
                "process": process,
                "prefix": service["prefix"]
            }

    # Start Outlook monitor thread
    outlook_thread = threading.Thread(target=outlook_monitor_thread_func, daemon=False)
    outlook_thread.start()

    if not running_processes:
        print("\n❌ Failed to launch core services. Exiting.")
        shutdown_event.set()
        outlook_thread.join(timeout=5)
        sys.exit(1)

    print(f"\n✅ All services launched!")
    print("   All subprocess output appears below with [service_name] prefix.")
    print("   Double blank lines separate different service outputs for clarity.")
    print("   Press Ctrl+C to stop all services.\n")

    # Main loop - monitor health and auto-restart failed services
    try:
        while not shutdown_event.is_set():
            # Check if Outlook monitor thread died unexpectedly
            if not outlook_thread.is_alive():
                with lock:
                    print("\n" + "!"*70)
                    print("[SYSTEM] ⚠️  Outlook monitor thread died! Restarting...")
                    print("!"*70 + "\n")

                # Restart the monitor thread
                outlook_thread = threading.Thread(target=outlook_monitor_thread_func, daemon=False)
                outlook_thread.start()
                time.sleep(2)  # Brief pause before resuming

            # Check if always-on services died unexpectedly
            for name in list(running_processes.keys()):
                info = running_processes[name]
                if info["process"].poll() is not None:
                    with lock:
                        print(f"\n[SYSTEM] ⚠️  {name} died! Restarting...")

                    # Find the service config and restart
                    service_config = next((s for s in ALWAYS_ON_SERVICES if s["name"] == name), None)
                    if service_config:
                        process = launch_service(service_config)
                        if process:
                            running_processes[name] = {
                                "process": process,
                                "prefix": service_config["prefix"]
                            }
                        else:
                            with lock:
                                print(f"[SYSTEM] ❌ Failed to restart {name}")

            time.sleep(0.5)

    except KeyboardInterrupt:
        pass
    finally:
        shutdown_all_services()
        outlook_thread.join(timeout=10)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
