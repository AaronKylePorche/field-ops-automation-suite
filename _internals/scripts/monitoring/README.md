# Background Monitoring Services - Claim Watcher Suite

Complete guide to the three background services that work together for automated claim processing.

## Overview

The **Claim Watcher Suite** consists of three independent services that run continuously:

1. **Supervisor** (supervisor1.py) - Manages email watcher
2. **Ticket Reader** (Ticket_Reader.py) - Processes claim queue
3. **Keep Awake** (keep_awake.py) - Prevents system sleep

Each runs in its own window and logs to `data/output/`.

## Quick Start

### Launch All Three at Once

Via Bismillah launcher:
```bash
python scripts\core\Bismillah.py
# Select: [10] 🟢 Start Claim Watcher Suite
```

### Or Launch Individually

```bash
python supervisor1.py
python Ticket_Reader.py
python keep_awake.py
```

## The Three Services

### 1. Supervisor (supervisor1.py) 🔍

**Purpose:** Manages the email watcher lifecycle

**What it does:**
- Monitors Outlook.exe
- Launches email scanner when Outlook opens
- Stops when Outlook closes
- Auto-restarts if it crashes

**Log file:** `data/output/supervisor_log.txt`

### 2. Ticket Reader (Ticket_Reader.py) 📨

**Purpose:** Processes claim tickets from queue

**What it does:**
- Watches `scripts/queue/` folder
- Detects new ticket files (wop_ticket_*.txt)
- Launches WOP22.py or WOP.py to process
- Removes tickets after processing
- Logs all activity

**Log file:** `data/output/ticket_reader_log.txt`

### 3. Keep Awake (keep_awake.py) ⏰

**Purpose:** Prevents Windows sleep

**What it does:**
- Keeps system awake (prevents hibernation)
- Keeps display on
- Uses minimal resources
- Refreshes every 30 seconds

**Output:** Status in console window

## How They Work Together

```
Email arrives
    ↓
Email_Scanner.py (watched by Supervisor)
    ↓
Drops ticket in scripts/queue/
    ↓
Ticket_Reader.py detects it
    ↓
Launches WOP22.py
    ↓
Removes processed ticket
    ↓
Ready for next ticket

(Keep Awake runs continuously to keep system awake)
```

## Required Setup

Before running, you need:

1. **Email_Scanner.py** - Copy to `scripts/core/`
2. **WOP22.py** or **WOP.py** - Copy to `scripts/WOP/`
3. **queue folder** - Auto-created at `scripts/queue/`

## Configuration

In `config/config.py`:

```python
"10": {
    "name": "🟢 Start Claim Watcher Suite",
    "enabled": True,
    "background_scripts": [
        os.path.join(BASE_DIR, "scripts", "monitoring", "supervisor1.py"),
        os.path.join(BASE_DIR, "scripts", "monitoring", "Ticket_Reader.py"),
        os.path.join(BASE_DIR, "scripts", "monitoring", "keep_awake.py"),
    ]
},
```

## Monitoring & Logs

Each service writes detailed logs:

```bash
# View supervisor activity
type data\output\supervisor_log.txt

# View ticket processing
type data\output\ticket_reader_log.txt
```

Check logs weekly for:
- Crashes or errors
- Processing failures
- System issues

## Stopping Services

Each runs in its own window. To stop:

1. Click the window
2. Press Ctrl+C
3. Window closes

To stop all three:
- Close each window individually

## Troubleshooting

**Services crash:**
- Check logs for error messages
- Verify Email_Scanner.py exists
- Verify WOP22.py or WOP.py exists
- Ensure scripts are in correct folders

**Tickets not processing:**
- Check if Ticket_Reader window is open
- Verify WOP script location
- Check ticket_reader_log.txt for errors

**Keep Awake doesn't work:**
- Windows only (not Mac/Linux)
- Requires admin privileges
- Safe to skip if not needed

## Files in This Module

- **supervisor1.py** - Email watcher supervisor (100+ lines)
- **Ticket_Reader.py** - Claim ticket processor (200+ lines)
- **keep_awake.py** - System sleep preventer (100+ lines)
- **CLAIM_WATCHER_SUITE.md** - Complete detailed guide
- **README.md** - This file

## Complete Documentation

For comprehensive details, see:
→ **CLAIM_WATCHER_SUITE.md**

This includes:
- Detailed explanation of each service
- Setup instructions
- Configuration options
- Troubleshooting guide
- Startup automation
- Performance notes
- Security considerations

## Performance

All three services together:
- **RAM Usage:** ~60-80 MB
- **CPU Usage:** <2% (idle)
- **Safe to run 24/7:** Yes
- **System impact:** Minimal

## Examples

### Start at Startup

Create Windows Task Scheduler job:
1. Task: "Start Claim Watcher"
2. Trigger: At startup
3. Action: `python scripts\core\Bismillah.py`

### Monitor Logs in Real-time

Keep a terminal watching:
```bash
type data\output\supervisor_log.txt
```

Then check periodically for issues.

### Process Tickets Manually

If you have a ticket file, move it to:
```
scripts/queue/wop_ticket_test.txt
```

Ticket_Reader will process it immediately.

---

## Ready to Start?

1. Ensure Email_Scanner.py is in scripts/core/
2. Ensure WOP22.py is in scripts/WOP/
3. Run: `python scripts\core\Bismillah.py`
4. Select: [10] 🟢 Start Claim Watcher Suite
5. Three windows open - you're running!

👉 **For complete details, see CLAIM_WATCHER_SUITE.md**
