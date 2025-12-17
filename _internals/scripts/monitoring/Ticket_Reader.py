# WOP_runner_console.py
# Persistent console that reuses a single window to run WOP on demand.
# It watches a simple ticket queue directory and runs WOP once per ticket.
#
# Run this in its own console (keep it open):
#   python WOP_runner_console.py
#
# The watcher will drop a ticket file in ./queue/ to trigger a run.

import time
import sys
import subprocess
from pathlib import Path
from datetime import datetime

# Get queue folder from config (same as Email_Scanner uses)
import os
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent / "config"))
try:
    import config
    QUEUE_DIR = Path(config.OUTLOOK_SETTINGS["queue_folder"])
except (ImportError, KeyError):
    # Fallback if config import fails
    BASE_DIR = Path(__file__).resolve().parent.parent
    QUEUE_DIR = BASE_DIR / "queue"

QUEUE_DIR.mkdir(exist_ok=True)

def set_console_title(title: str):
    """Set the console window title for clarity."""
    if os.name == "nt":
        os.system(f"title {title}")

def print_banner(name: str, description: str):
    """Print a consistent header banner at startup."""
    print("=" * 70)
    print(f"{name}")
    print("=" * 70)
    print(description.strip())
    print("=" * 70 + "\n")


def get_timestamp():
    """Return current time in human-readable format: MM/DD/YYYY H:MMAm/Pm"""
    return datetime.now().strftime("%m/%d/%Y %I:%M%p")


def find_wop():
    """Try WOP/WOP22.py first, then WOP/WOP.py. Return Path or None."""
    scripts_dir = Path(__file__).resolve().parent.parent
    p1 = (scripts_dir / "WOP" / "WOP22.py").resolve()
    p2 = (scripts_dir / "WOP" / "WOP.py").resolve()
    if p1.exists():
        return p1
    if p2.exists():
        return p2
    return None

def run_wop_once():
    wop_path = find_wop()
    if not wop_path:
        print(f"[{get_timestamp()}] ! WOP script not found (tried WOP22.py and WOP.py). Put it under scripts/WOP/")
        return 1
    print(f"[{get_timestamp()}] -> Running WOP: {wop_path.name}")
    try:
        proc = subprocess.run([sys.executable, str(wop_path)], cwd=str(wop_path.parent))
        print(f"[{get_timestamp()}] -> WOP exit code: {proc.returncode}")
        return proc.returncode
    except Exception as e:
        print(f"[{get_timestamp()}] ! Failed to run WOP:", e)
        return 1

def main():
    print(f"[{get_timestamp()}] === WOP Runner Console ===")
    print(f"[{get_timestamp()}] Watching queue: {QUEUE_DIR}")
    print(f"[{get_timestamp()}] Leave this window open; each ticket will run WOP once.\n")
    while True:
        try:
            tickets = sorted(QUEUE_DIR.glob("wop_ticket_*.txt"))
            if tickets:
                ticket = tickets[0]
                print(f"[{get_timestamp()}] * Ticket detected: {ticket.name}")
                # Remove first to avoid double-run if WOP crashes
                try:
                    ticket.unlink(missing_ok=True)
                except Exception:
                    pass
                run_wop_once()
            time.sleep(1.0)
        except KeyboardInterrupt:
            print(f"\n[{get_timestamp()}] Stopping runner. Bye.")
            break
        except Exception as e:
            print(f"[{get_timestamp()}] ! Runner loop error:", e)
            time.sleep(2.0)

if __name__ == "__main__":
    set_console_title("Ticket Reader - Processed Claim Output")
    print_banner(
        "Ticket Reader",
        """
        Scans the ticket queue directory to process any new claims.
        - Reads new tickets from ./queue/
        - Launches WOP22.py once per ticket and prints the exit code
        - Removes processed tickets to avoid duplicates
        """
    )
    main()
