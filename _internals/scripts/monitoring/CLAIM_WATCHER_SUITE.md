# Claim Watcher Suite - Complete Guide

## Overview

The **Claim Watcher Suite** is a set of three background services that run continuously to monitor emails, process claims, and keep your system awake.

```
┌─────────────────────────────────────────────────────────┐
│ Claim Watcher Suite (Runs in 3 Separate Windows)        │
├─────────────────────────────────────────────────────────┤
│ 1. Supervisor (🔍)      - Manages email watcher         │
│ 2. Ticket Reader (📨)   - Processes claim tickets       │
│ 3. Keep Awake (⏰)      - Prevents system sleep         │
└─────────────────────────────────────────────────────────┘
```

---

## The Three Services

### 1. Supervisor (supervisor1.py) 🔍
**What it does:**
- Monitors Outlook.exe status
- Launches the email watcher when Outlook opens
- Stops the email watcher when Outlook closes
- Restarts the email watcher if it crashes
- Logs all events

**Why you need it:**
- Ensures email scanning only runs when Outlook is open
- Automatically recovers from crashes
- Prevents resource waste

**How to run:**
```bash
python scripts\monitoring\supervisor1.py
```

**Output:**
- Logs to: `data/output/supervisor_log.txt`
- Each action is timestamped and logged

---

### 2. Ticket Reader (Ticket_Reader.py) 📨
**What it does:**
- Monitors `scripts/queue/` folder for ticket files
- When a new ticket appears (wop_ticket_*.txt), processes it
- Launches WOP22.py or WOP.py to handle the claim
- Removes processed tickets to avoid duplicates
- Logs all ticket processing

**Why you need it:**
- Processes claims automatically as they arrive
- Works with Supervisor to create a ticket queue system
- Prevents duplicate processing

**How to run:**
```bash
python scripts\monitoring\Ticket_Reader.py
```

**Folder Structure:**
```
scripts/
├── queue/                    ← Ticket files appear here
│   ├── wop_ticket_001.txt   ← Processed by Ticket Reader
│   ├── wop_ticket_002.txt
│   └── wop_ticket_003.txt
├── WOP/
│   ├── WOP22.py            ← Launched by Ticket Reader
│   └── WOP.py
└── Ticket_Reader.py
```

**Output:**
- Logs to: `data/output/ticket_reader_log.txt`

---

### 3. Keep Awake (keep_awake.py) ⏰
**What it does:**
- Prevents Windows from sleeping
- Keeps display from turning off
- Minimal resource usage (refreshes every 30 seconds)
- Simple exit with Ctrl+C

**Why you need it:**
- Ensures background services keep running overnight
- No interruptions from sleep/hibernation
- Perfect for unattended processing

**How to run:**
```bash
python scripts\monitoring\keep_awake.py
```

**Output:**
- Status messages in console window
- Refreshes every 30 seconds
- Ctrl+C releases hold and exits

---

## Launching the Suite

### Option A: From Bismillah Launcher (Recommended)

1. Run: `python scripts\core\Bismillah.py`
2. Select: `[10] 🟢 Start Claim Watcher Suite`
3. Three windows will open automatically
4. Leave them running in the background

### Option B: Manual Launch

Open three separate Command Prompt windows and run:

```bash
# Window 1: Supervisor
python scripts\monitoring\supervisor1.py

# Window 2: Ticket Reader
python scripts\monitoring\Ticket_Reader.py

# Window 3: Keep Awake
python scripts\monitoring\keep_awake.py
```

---

## How They Work Together

### Workflow

```
Email Arrives
    ↓
Email_Scanner.py (supervised by Supervisor)
    ↓
Drops ticket file in scripts/queue/
    ↓
Ticket_Reader.py detects it
    ↓
Launches WOP22.py to process claim
    ↓
Removes ticket file
    ↓
Waits for next ticket
```

### System Status

```
Keep Awake (always running)
├─ System awake + display on
│
Supervisor (always running)
├─ Monitors Outlook
├─ Launches/stops Email_Scanner
├─ Restarts if it crashes
│
Ticket_Reader (always running)
├─ Watches queue folder
├─ Processes tickets
├─ Logs all activity
```

---

## Configuration

### Required Scripts

For the suite to work, you need:

1. **Email_Scanner.py** (in scripts/core/ or scripts/)
   - Supervised by Supervisor
   - Generates tickets when emails arrive

2. **WOP22.py or WOP.py** (in scripts/WOP/)
   - Launched by Ticket Reader
   - Processes each claim ticket

### Setting Up

1. Copy `Email_Scanner.py` to `scripts/core/`
2. Copy `WOP22.py` or `WOP.py` to `scripts/WOP/`
3. Create queue folder (auto-created if missing):
   ```
   scripts/queue/
   ```
4. Run the suite!

### Customization

Edit each script directly to customize:

**supervisor1.py:**
- Change RESTART_BACKOFF_SEC (wait time after crash)
- Modify LOG_FILE location
- Change which email script to supervise

**Ticket_Reader.py:**
- Change QUEUE_DIR location
- Modify WOP search paths
- Adjust timeout (currently 5 minutes)

**keep_awake.py:**
- Change refresh interval (currently 30 seconds)
- Already minimal, no other customization needed

---

## Logs

All activity is logged for monitoring and troubleshooting:

### Supervisor Logs
- **File:** `data/output/supervisor_log.txt`
- **Contains:** Outlook status, watcher launch/stop, crashes
- **Example:**
  ```
  [2025-01-20 09:00:15] Supervisor started
  [2025-01-20 09:00:30] 📨 Outlook started - launching watcher
  [2025-01-20 09:00:32] ✅ Watcher launched (PID: 1234)
  [2025-01-20 09:15:45] ⚠️  Watcher crashed (exit code: 1)
  ```

### Ticket Reader Logs
- **File:** `data/output/ticket_reader_log.txt`
- **Contains:** Tickets found, WOP runs, status codes
- **Example:**
  ```
  [2025-01-20 09:15:50] * Processing ticket: wop_ticket_001.txt
  [2025-01-20 09:15:51] → Running: WOP22.py
  [2025-01-20 09:16:02] ✅ WOP22.py completed (exit code: 0)
  [2025-01-20 09:16:03]   ✓ Ticket removed
  ```

### Monitoring

Check logs periodically:
```bash
# View supervisor activity
type data\output\supervisor_log.txt

# View ticket processing
type data\output\ticket_reader_log.txt

# Or check from within Windows
# data/output/ folder in File Explorer
```

---

## Troubleshooting

### "Supervisor can't find Email_Scanner.py"

**Problem:** Script can't locate your email scanner

**Solution:**
1. Create if missing: `scripts/core/`
2. Copy your `Email_Scanner.py` there
3. Or modify supervisor1.py to look in correct location:
   ```python
   CANDIDATES = [
       BASE_DIR / "core" / "Email_Scanner.py",
       BASE_DIR / "YourFolderName" / "Email_Scanner.py",
   ]
   ```

### "Ticket Reader can't find WOP22.py"

**Problem:** Script can't locate WOP processor

**Solution:**
1. Copy your WOP script to: `scripts/WOP/`
2. Name it: `WOP22.py` or `WOP.py`
3. Check the ticket reader log for details:
   ```
   cat data/output/ticket_reader_log.txt
   ```

### "Keep Awake says Windows ctypes not available"

**Problem:** Script can't prevent sleep

**Likely cause:** Not running on Windows, or running in unusual environment

**Solution:**
- This script is Windows-only
- On non-Windows systems, just skip it
- Or comment out this script from config.py

### Services keep crashing

**Problem:** Services exit unexpectedly

**Solution:**
1. Check the logs for error messages
2. Verify all required scripts exist
3. Test Email_Scanner.py manually
4. Test WOP22.py manually
5. Check for Python errors: `python -m py_compile scripts\monitoring\supervisor1.py`

### Tickets not being processed

**Problem:** Ticket files appear but aren't processed

**Troubleshooting:**
1. Check if Ticket Reader is running
2. Verify WOP script exists at: `scripts/WOP/WOP22.py` or `scripts/WOP/WOP.py`
3. Check `data/output/ticket_reader_log.txt` for errors
4. Test WOP manually: `python scripts\WOP\WOP22.py`

---

## Starting at Startup

### Option 1: Windows Task Scheduler

1. Open Task Scheduler
2. Create Basic Task
3. Name: "Start Claim Watcher Suite"
4. Trigger: "At startup"
5. Action: Start a program
   - Program: `python`
   - Arguments: `scripts\core\Bismillah.py`
   - Start in: Your New User Package folder

6. On the last page, check: "Open the Properties dialog"
7. Under "Settings", check: "Run whether user is logged in or not"
8. Click OK

### Option 2: Batch File

Create `start_claim_watcher.bat`:
```batch
@echo off
cd /d "%~dp0"
python scripts\core\Bismillah.py
```

Then:
1. Right-click the .bat file
2. Send to → Desktop (create shortcut)
3. Right-click shortcut → Properties
4. Advanced → Check "Run as administrator"

### Option 3: Run Bismillah manually

Each time you start work:
```bash
python scripts\core\Bismillah.py
# Then select [10] to start the suite
```

---

## Performance & Resources

### System Impact

- **Supervisor:** ~15-30 MB RAM, 0-2% CPU (idle)
- **Ticket Reader:** ~15-30 MB RAM, 0-2% CPU (idle)
- **Keep Awake:** ~5-10 MB RAM, <1% CPU (refreshes every 30s)

**Total:** Minimal impact, safe to run 24/7

### Recommendations

- Run on a dedicated user account or admin account
- Consider a system tray tool if you need visibility
- Use Windows Task Scheduler for automatic startup
- Check logs weekly for issues

---

## Customization Examples

### Custom Queue Folder

**In Ticket_Reader.py, change:**
```python
QUEUE_DIR = BASE_DIR / "queue"
```

**To:**
```python
QUEUE_DIR = BASE_DIR.parent / "data" / "queue"
```

### Different WOP Location

**In Ticket_Reader.py, change:**
```python
def find_wop():
    candidates = [
        BASE_DIR / "WOP" / "WOP22.py",
        BASE_DIR / "WOP" / "WOP.py",
    ]
```

**To:**
```python
def find_wop():
    candidates = [
        Path("C:/MyScripts/WOP/WOP22.py"),  # Custom location
        BASE_DIR / "WOP" / "WOP.py",
    ]
```

### Increase Supervisor Restart Delay

**In supervisor1.py, change:**
```python
RESTART_BACKOFF_SEC = 10
```

**To:**
```python
RESTART_BACKOFF_SEC = 30  # Wait 30 seconds before restart
```

---

## Security Notes

- Services run with current user privileges
- No credentials stored in code
- Logs may contain claim information (keep private)
- Keep logs folder secure (`data/output/`)

---

## Summary

| Service | Purpose | Status Window |
|---------|---------|---------------|
| **Supervisor** | Manages email watcher | Shows Outlook monitoring |
| **Ticket Reader** | Processes claim tickets | Shows ticket activity |
| **Keep Awake** | Prevents system sleep | Shows status messages |

All three work together to create a complete **background claim processing system**.

---

## Quick Reference

```bash
# Launch from Bismillah
python scripts\core\Bismillah.py
# Then select: [10] 🟢 Start Claim Watcher Suite

# Or run individually
python scripts\monitoring\supervisor1.py
python scripts\monitoring\Ticket_Reader.py
python scripts\monitoring\keep_awake.py

# Check logs
type data\output\supervisor_log.txt
type data\output\ticket_reader_log.txt

# Stop a service
# Click the window and press Ctrl+C
```

---

**Ready to run?** Launch from Bismillah and select option [10]! 🚀
