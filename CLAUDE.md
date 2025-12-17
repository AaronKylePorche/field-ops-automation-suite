# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**KD Assistant (G5 Tools)** is a Windows automation suite for SCE (Southern California Edison) claims processing, document management, and reporting. The system uses a **configuration-driven architecture** where all customization happens in a single file (`_internals/config/config.py`), making it portable and easy to deploy.

## Architecture

### Core Components

1. **Launcher System** (`KD Assistant.py`)
   - Menu-driven interface that dynamically discovers scripts from config.py
   - Handles both foreground (blocking) and background (windowed) execution modes
   - Scripts are registered in `config.py` SCRIPTS dictionary with menu numbers as keys

2. **Configuration System** (`_internals/config/config.py`)
   - Master configuration file - all paths, settings, and script definitions
   - Auto-detects BASE_DIR for portability across drives/systems
   - Contains multiple setting sections: OUTLOOK_SETTINGS, WOP22_SETTINGS, JIS_SETTINGS, etc.
   - Scripts are read-only consumers of config (never modify it)

3. **Script Organization** (`_internals/scripts/`)
   - **core/**: Email scanning, claim processing, report generation, config editing
   - **monitoring/**: Background services (ClaimWatcherSuite, Ticket_Reader, keep_awake)
   - **JIS Automation/**: Job Instruction Sheet generation and consolidation
   - **Daily Report/**: LED change-out reporting
   - **Document Processing/**: PDF/document merging
   - **WOP/**: Workflow Operations Processor (AI-powered claim extraction)

4. **Data Flow**
   ```
   input/          →  Scripts Process  →  output/
   (user files)        (automation)        (results/logs)

   Outlook Inbox   →  Email_Scanner   →  Queue  →  WOP22  →  Excel
   ```

### Key Patterns

#### Config Import (Universal Pattern)

Every script MUST follow this exact import pattern:

```python
import sys
from pathlib import Path

# Navigate to config folder (adjust parent levels based on script location)
sys.path.insert(0, str(Path(__file__).parent.parent.parent / "config"))
import config

# Use config settings
WORKBOOK_PATH = config.WOP22_SETTINGS["excel_path"]
OUTPUT_FOLDER = Path(config.DATA_PATHS["output"])
```

**Location-specific paths:**
- Scripts in `core/`: `.parent.parent.parent / "config"`
- Scripts in `JIS Automation/`: `.parent.parent.parent / "config"`
- Scripts in root: `.parent / "_internals" / "config"`

#### Path Handling

- **config.py uses**: `os.path.join()` for Windows drive letter compatibility
- **Scripts use**: `Path()` objects from pathlib for modern path operations
- **Conversion at import**: `Path(config.SETTING["path"])`
- **Never hardcode paths** - all paths come from config

#### Excel Interaction

**Use COM Automation when file may be open:**
```python
import win32com.client as win32
excel = win32.Dispatch('Excel.Application')
excel.Visible = True  # Keep Excel visible to prevent file closing
workbook = excel.Workbooks.Open(WORKBOOK_PATH, ReadOnly=False)
worksheet = workbook.Sheets("SheetName")
data = worksheet.UsedRange.Value  # Read entire range
```

**Use openpyxl when file must be closed:**
```python
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
wb = load_workbook(WORKBOOK_PATH)
ws = wb["SheetName"]
ws.cell(row=1, column=1).value = "Data"
wb.save(WORKBOOK_PATH)
```

#### Outlook Integration

```python
import win32com.client as win32
outlook = win32.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# Navigate to folder using config path list
folder = namespace.Folders(config.OUTLOOK_SETTINGS["mailbox_name"])
for part in config.OUTLOOK_SETTINGS["target_folder_path"]:
    folder = folder.Folders(part)

# Create draft email (not sent automatically)
mail = outlook.CreateItem(0)  # 0 = MailItem
mail.To = "recipient@example.com"
mail.Subject = "Subject"
mail.Body = "Body text"
mail.Attachments.Add(attachment_path)
mail.Display()  # Show draft (user sends manually)
```

#### Queue-Based Async Processing

```python
# Producer (Email_Scanner) creates tickets:
ticket_path = QUEUE_FOLDER / f"wop_ticket_{timestamp}.txt"
ticket_path.write_text(json.dumps({
    "entry_id": mail.EntryID,
    "timestamp": datetime.now().isoformat()
}))

# Consumer (Ticket_Reader) processes tickets:
tickets = sorted(QUEUE_DIR.glob("wop_ticket_*.txt"))
if tickets:
    ticket = tickets[0]
    ticket.unlink()  # Remove BEFORE processing to avoid double-run
    subprocess.run([sys.executable, WOP_SCRIPT_PATH])
```

## Common Development Tasks

### Adding a New Script to the Launcher

1. **Place script** in appropriate `_internals/scripts/` subfolder
2. **Import config** using the universal pattern (adjust parent levels)
3. **Update config.py** SCRIPTS dictionary:
   ```python
   "4": {
       "name": "📊 Your Script Name",
       "path": os.path.join(BASE_DIR, "_internals", "scripts", "folder", "script.py"),
       "enabled": True,
       "description": "Brief description"
   }
   ```
4. **Test** by running `KD Assistant.py` and selecting the menu option

### Creating a New Report Script

1. Use the kd_report_generator.py as a template
2. Read source data using config paths
3. Transform data with pandas/openpyxl
4. Apply formatting (center text, auto-width columns, freeze panes, filters)
5. Save to `output/<report_type>_output/`
6. Create Outlook draft (not auto-send)

### Working with Excel Formatting

Common patterns from kd_report_generator.py:

```python
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.formatting.rule import ColorScaleRule

# Center alignment with wrap text
ws.cell(row, col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

# Header formatting
ws.cell(1, col).font = Font(bold=True)
ws.cell(1, col).fill = PatternFill(start_color='B4C7E7', end_color='B4C7E7', fill_type='solid')

# Date formatting
ws.cell(row, col).number_format = 'm/d/yyyy'

# Freeze panes (at C2 = freeze cols A-B and header row)
ws.freeze_panes = 'C2'

# Auto-filter
ws.auto_filter.ref = ws.dimensions

# Conditional formatting (color scale)
color_scale = ColorScaleRule(
    start_type='min', start_color='63BE7B',  # Green
    mid_type='percentile', mid_value=50, mid_color='FFEB84',  # Yellow
    end_type='max', end_color='F8696B'  # Red
)
ws.conditional_formatting.add('L2:L100', color_scale)
```

### Testing Scripts Independently

All scripts can run standalone:
```bash
cd "C:\Path\To\G5 Tools"
python "_internals\scripts\core\kd_report_generator.py"
```

## Important Conventions

### File Naming
- Scripts: PascalCase with underscores for readability (`Email_Scanner.py`, `Stand_Alone_Processor.py`)
- Config keys: lowercase_with_underscores (`whitelist_emails`, `target_folder_path`)
- Folders: Spaces allowed (`Daily Report/`, `JIS Automation/`)

### Error Handling
- Graceful degradation: Try COM → fallback to openpyxl
- User-friendly error messages (avoid raw stack traces in production)
- Log errors to files in `output/` folders

### Data Flow Rules
1. Input files: Users place in `input/<module>/`
2. Output files: Scripts write to `output/<module>/`
3. Templates: Bundled in `_internals/data/templates/`
4. Queue: Ephemeral ticket files (deleted after processing)
5. Logs: Append-only (never truncate automatically)

### Security
- **Never hardcode credentials** - use environment variables or config
- **Whitelist filtering** - only process emails from trusted senders (OUTLOOK_SETTINGS["whitelist_emails"])
- **No absolute user paths** in code - all paths via config for portability

## Key Technical Details

### Excel COM vs openpyxl Decision Tree

Use **COM Automation** when:
- File may be open by user
- Need to preserve complex formatting
- Working with Excel formulas/charts
- Keep `excel.Visible = True` to prevent workbook closing on script end

Use **openpyxl** when:
- File must be closed
- Simple read/write operations
- Batch processing multiple files
- Applying cell formatting programmatically

### Timezone Handling for Excel

Excel doesn't support timezone-aware datetimes. Always strip timezone info:
```python
for col in df.columns:
    df[col] = df[col].apply(lambda x:
        x.replace(tzinfo=None) if isinstance(x, pd.Timestamp) and x.tzinfo is not None
        else x
    )
```

### Background Service Management

The ClaimWatcherSuite_Unified.py manages multiple background processes:
- Email_Scanner (Outlook monitoring - event-driven)
- Ticket_Reader (queue processing - 1-second polling)
- keep_awake (system awake - 30-second refresh)

Services auto-start with launcher option 7 and run in unified window with prefixed output.

**Critical Pattern - Process Group Isolation:**

All subprocesses MUST be created in separate process groups on Windows:
```python
subprocess.Popen(
    [sys.executable, script_path],
    creationflags=subprocess.CREATE_NEW_PROCESS_GROUP if sys.platform == "win32" else 0,
    # ... other parameters
)
```

**Why:** Without `CREATE_NEW_PROCESS_GROUP`, signals like `CTRL_BREAK_EVENT` propagate to the parent process, causing unwanted shutdowns. This is essential for ClaimWatcherSuite to survive when stopping child processes.

**Outlook Monitoring Pattern - WMI + Polling Hybrid:**

The Outlook monitor uses BOTH WMI events and polling simultaneously:
```python
# Check WMI events (instant notification)
if stop_wmi:
    evt = next_event_nonblocking(stop_wmi, timeout_ms=250)
    if evt is not None:
        outlook_is_running = False
        stop_email_scanner()

# ALWAYS poll as safety net (even when WMI available)
now_up = is_outlook_running()
if now_up != outlook_is_running:
    # Handle state change
```

**Why:** WMI events can be missed due to timing issues, event queue overflow, or Windows delivery delays. Polling (1-second interval) runs continuously as a fallback to guarantee state changes are detected.

## Codebase-Specific Notes

### Config.py Settings Structure

Key sections to understand:
- **SCRIPTS**: Menu registry (dict with number keys)
- **DATA_PATHS**: Standard input/output folders
- **OUTLOOK_SETTINGS**: Email whitelist, folder paths (hierarchical list), mailbox name
- **WOP22_SETTINGS**: Excel tracker path, OpenAI model, district mapping
- **JIS_SETTINGS**: Templates, conversion charts, setup file
- **SLIPSENDER_SETTINGS**: Material tracking, packing slip paths, email CC list

### Special Configuration Files

- `Setup.xlsx`: JIS header text and structure exclusions
- `District.xlsx`: Maps sender emails to district codes
- `ConversionChart.xlsx`: LED/HPSV/MH wattage mappings
- `.env`: API keys (OPENAI_API_KEY) - never commit

### Column Letter to Index Conversion

When working with Excel columns by letter:
```python
def col_letter_to_index(letter):
    """Convert Excel column letter to 0-based index"""
    result = 0
    for char in letter:
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result - 1

# Example: 'AA' = 26, 'CJ' = 87
```

### Deployment Portability

The system is designed for zero-configuration deployment:
1. Extract zip to any location
2. Edit config.py (update paths, emails)
3. Run setup.bat (installs Python + dependencies)
4. Launch KD Assistant.py

All scripts automatically adapt to new BASE_DIR location.
