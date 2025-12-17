# Time Tracking & Timesheet Module

This module handles time clock tracking and weekly timesheet generation.

## Scripts

- **generate_weekly_timesheet.py** - Generates weekly timesheet reports

## Data Folders

- **Input/** - Place time clock data files here
- **Output/** - Generated timesheet reports will be saved here

## Setup

1. Place your time clock Excel file in the `Input/` folder
   - File should be named: `Time Clock.xlsx` (or update config.py)
2. Run the script via the Bismillah launcher: `[6] 🕒 Generate Weekly Timesheet`

## Data Format

Your time clock file should have columns like:
- Employee Name
- Date
- Hours Worked
- Task/Project

See the script documentation for exact requirements.

## Configuration

Update `config/config.py` to customize paths:

```python
TIMESHEET_SETTINGS = {
    "input_folder": os.path.join(DATA_PATHS["input"], "timesheet_input"),
    "output_folder": os.path.join(DATA_PATHS["output"], "timesheet_output"),
    "time_clock_file": os.path.join(DATA_PATHS["input"], "Time Clock.xlsx"),
}
```
