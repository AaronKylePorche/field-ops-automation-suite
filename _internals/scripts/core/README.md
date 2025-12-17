# Core Module

This module contains the main launcher and core utilities.

## Files

- **Bismillah.py** - Main application launcher
  - Displays menu of all available scripts
  - Executes selected scripts
  - Handles background services
  - Configuration-driven (reads from config.py)

## Usage

### Run the Launcher

```bash
python Bismillah.py
```

Or double-click the file in Windows Explorer.

### From Command Line

```bash
cd scripts\core
python Bismillah.py
```

### Menu Options

The launcher displays different options based on your `config/config.py` settings.

**Typical Options:**
1. Add Claims To Tracker
2. Combine Claims & Permits
3. Daily LED Report
4. Weekly Timesheet
5. JIS Automation
6. Master JIS
7. Claim Watcher Suite (background services)
8. Packing Slips Email
q. Quit

### Background Services

Some scripts (like the Claim Watcher Suite) run in separate windows:
- They don't block the menu
- They continue running in the background
- You can launch multiple scripts simultaneously

## Configuration

The launcher reads all its settings from:
```
config/config.py
```

To add, remove, or customize scripts:
1. Edit `config/config.py`
2. Add/remove entries in the `SCRIPTS` dictionary
3. Save the file
4. Restart the launcher

## Creating a Shortcut

On Windows, you can create a desktop shortcut:

1. Right-click Bismillah.py
2. Select "Send to" → "Desktop (create shortcut)"
3. Or manually:
   - Right-click Desktop
   - New → Shortcut
   - Path: `C:\path\to\scripts\core\Bismillah.py`

## Advanced Features

### Debug Mode

Enable debugging to see more information:

1. Open `config/config.py`
2. Set `DEBUG = True`
3. Restart the launcher

This will:
- Show configuration details on startup
- Validate all script paths
- Warn about missing files

### Custom Menus

You can customize which scripts appear in the menu:

1. Edit `config/config.py`
2. Set `"enabled": False` for scripts you don't want to show
3. Set `"enabled": True` to show them
4. Restart the launcher

---

For more information, see the main README.md in the parent directory.
