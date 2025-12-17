# supervisor1.py
# Watches OUTLOOK.EXE with WMI start/stop events and manages your email watcher.
# - If Outlook closes: immediately stops the watcher.
# - If Outlook starts (or restarts): (re)launches the watcher.
# - If the watcher crashes while Outlook is running: restarts it after a short backoff.
#
# Requirements: pywin32  (pip install pywin32)
# Run in the same folder as your watcher scripts.

import os
import sys
import time
import signal
import subprocess
from pathlib import Path
from datetime import datetime

try:
    import win32com.client
except Exception:
    print("pywin32 is required. Install with:\n  python -m pip install pywin32")
    sys.exit(1)

# -------- CONFIG ------------------------------------------------------------
# Pick which watcher to supervise. If both present, prefer Email_Scanner.py.
# Look in scripts/core folder (relative to scripts/monitoring)
BASE_DIR = Path(__file__).resolve().parent.parent / "core"
CANDIDATES = ["Email_Scanner.py", "InboxTrigger_WOP.py"]
for _name in CANDIDATES:
    if (BASE_DIR / _name).exists():
        WATCHER_SCRIPT = (BASE_DIR / _name)
        break
else:
    print("ERROR: No watcher script found (expected Email_Scanner.py or InboxTrigger_WOP.py in scripts/core/).")
    sys.exit(1)

PYTHON_EXE = sys.executable  # use current interpreter
RESTART_BACKOFF_SEC = 10     # wait before restarting after crash
OUTLOOK = "OUTLOOK.EXE"
# ---------------------------------------------------------------------------


def set_console_title(title: str):
    if os.name == "nt":
        os.system(f"title {title}")

def print_banner(name: str, description: str):
    print("=" * 70)
    print(f"{name}")
    print("=" * 70)
    print(description.strip())
    print("=" * 70 + "\n")


def get_timestamp():
    """Return current time in human-readable format: MM/DD/YYYY H:MMAm/Pm"""
    return datetime.now().strftime("%m/%d/%Y %I:%M%p")


def is_outlook_running():
    """Fast check using `tasklist` (no extra deps)."""
    try:
        # /FI filter makes it fast; /NH removes header; stdout as text
        proc = subprocess.run(
            ["tasklist", "/FI", f"IMAGENAME eq {OUTLOOK}", "/NH"],
            capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW
        )
        txt = (proc.stdout or "").strip()
        # When not found, tasklist still returns a line with 'INFO:' text
        return OUTLOOK in txt.upper()
    except Exception:
        return False

def launch_watcher():
    """Start the watcher in its own console window."""
    try:
        print(f"[{get_timestamp()}] -> Launching watcher: {WATCHER_SCRIPT.name}")
        # New console so this supervisor stays readable
        CREATE_NEW_CONSOLE = 0x00000010
        p = subprocess.Popen(
            [PYTHON_EXE, str(WATCHER_SCRIPT)],
            cwd=str(WATCHER_SCRIPT.parent),
            creationflags=CREATE_NEW_CONSOLE
        )
        return p
    except Exception as e:
        print(f"[{get_timestamp()}] ! Failed to start watcher:", e)
        return None

def stop_watcher(p):
    """Terminate watcher gracefully; force-kill if needed."""
    if not p:
        return
    try:
        if p.poll() is None:
            print(f"[{get_timestamp()}] -> Stopping watcher...")
            # Try gentle first
            try:
                if os.name == "nt":
                    # Send CTRL_BREAK to the watcher console, fallback to terminate
                    p.send_signal(signal.CTRL_BREAK_EVENT)
                    time.sleep(2)
            except Exception:
                pass
            # Then terminate
            p.terminate()
            try:
                p.wait(timeout=5)
            except Exception:
                pass
            if p.poll() is None:
                print(f"[{get_timestamp()}] ! Watcher didn't exit - killing.")
                p.kill()
        else:
            # already exited
            pass
    except Exception as e:
        print(f"[{get_timestamp()}] ! Error stopping watcher:", e)

def wmi_watchers():
    """Create WMI event subscriptions for Outlook start/stop."""
    svc = win32com.client.GetObject("winmgmts:root\\cimv2")
    startq = "SELECT * FROM Win32_ProcessStartTrace WHERE ProcessName='OUTLOOK.EXE'"
    stopq  = "SELECT * FROM Win32_ProcessStopTrace  WHERE ProcessName='OUTLOOK.EXE'"
    start_watcher = svc.ExecNotificationQuery(startq)
    stop_watcher  = svc.ExecNotificationQuery(stopq)
    return start_watcher, stop_watcher

def next_event_nonblocking(watcher, timeout_ms=500):
    """Return an event object or None (timeout)."""
    try:
        # pywin32 allows NextEvent with timeout in ms
        return watcher.NextEvent(timeout_ms)
    except Exception:
        return None

def main():
    print(f"[{get_timestamp()}] Starting Outlook session manager...")


    watcher_proc = None

    # Set up WMI event listeners (instant notifications on start/stop)
    try:
        start_wmi, stop_wmi = wmi_watchers()
    except Exception as e:
        print(f"[{get_timestamp()}] ! Failed to create WMI watchers. Falling back to periodic checks.", e)
        start_wmi = stop_wmi = None

    # Initial state sync
    outlook_up = is_outlook_running()
    print(f"[{get_timestamp()}] Initial Outlook state: {'RUNNING' if outlook_up else 'NOT RUNNING'}")
    if outlook_up and watcher_proc is None:
        watcher_proc = launch_watcher()

    try:
        while True:
            # 1) Handle WMI events (near-instant)
            if stop_wmi:
                evt = next_event_nonblocking(stop_wmi, timeout_ms=250)
                if evt is not None:
                    print(f"[{get_timestamp()}] ! Outlook STOP detected (WMI).")
                    # Stop watcher immediately
                    stop_watcher(watcher_proc)
                    watcher_proc = None
                    outlook_up = False

            if start_wmi:
                evt = next_event_nonblocking(start_wmi, timeout_ms=250)
                if evt is not None:
                    print(f"[{get_timestamp()}] OK Outlook START detected (WMI).")
                    outlook_up = True
                    # (Re)launch watcher when Outlook is up
                    if watcher_proc is None:
                        watcher_proc = launch_watcher()

            # 2) If no WMI, do a light periodic check (1s loop)
            if not start_wmi or not stop_wmi:
                now_up = is_outlook_running()
                if now_up != outlook_up:
                    outlook_up = now_up
                    if outlook_up:
                        print(f"[{get_timestamp()}] OK Outlook RUNNING (poll).")
                        if watcher_proc is None:
                            watcher_proc = launch_watcher()
                    else:
                        print(f"[{get_timestamp()}] ! Outlook NOT RUNNING (poll).")
                        stop_watcher(watcher_proc)
                        watcher_proc = None

            # 3) Watcher health: if it crashed while Outlook is up, restart after backoff
            if watcher_proc is not None:
                code = watcher_proc.poll()
                if code is not None:
                    print(f"[{get_timestamp()}] ! Watcher exited with code {code}.")
                    watcher_proc = None
                    if outlook_up:
                        print(f"[{get_timestamp()}] -> Restarting watcher in {RESTART_BACKOFF_SEC}s (Outlook is up)...")
                        time.sleep(RESTART_BACKOFF_SEC)
                        # Check Outlook still up after backoff
                        if is_outlook_running():
                            watcher_proc = launch_watcher()
                        else:
                            print(f"[{get_timestamp()}] ...Skipped restart - Outlook went down.")

            time.sleep(1.0)

    except KeyboardInterrupt:
        print(f"\n[{get_timestamp()}] Stopping supervisor...")
    finally:
        stop_watcher(watcher_proc)
        print(f"[{get_timestamp()}] Bye.")

if __name__ == "__main__":
    set_console_title("Outlook Session Manager")
    print_banner(
        "Outlook Session Manager",
        """
        Watches for OUTLOOK.EXE start/stop events and manages your Outlook Monitor.
        Makes sure Outlook Monitor is always looking at the most current outlook session. Restarts the monitor if outlook is closed and reopened.
        """
    )
    main()
