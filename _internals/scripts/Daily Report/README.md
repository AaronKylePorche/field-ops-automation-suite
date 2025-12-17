# Daily LED Change Out Report

## What It Does

This script generates daily LED change out reports from an Excel inventory file. It:

1. Reads an Excel file from the `Input` folder
2. Processes dates from the "LED Change Date" column
3. Groups data by date and week
4. Creates a formatted text report with daily totals and running totals
5. Saves the report to the `Output` folder

## Quick Start

### 1. Add Your Excel File

Place your Excel file (with LED change data) in the `Input` folder:
```
scripts/Daily Report/Input/
```

**Requirements:**
- File must contain a column named: `LED Change Date`
- Dates must be in a recognizable format (Excel dates, YYYY-MM-DD, etc.)
- File name pattern: `*.xlsx`, `*.xls`, or `*.xlsm`

### 2. Run the Script

Option A: From the Bismillah launcher
```
[5] 💡 Run Daily LED Change Out Report
```

Option B: Directly from command line
```bash
cd scripts\Daily Report
python Daily.py
```

### 3. Follow the Prompts

```
Enter the city/region for the report (or press Enter for none): Ventura
Enter number of weeks to include (default: 2): 2
```

### 4. Check the Output

Your report will be saved in:
```
scripts/Daily Report/Output/
Daily_Report_Ventura_YYYY-MM-DD_to_YYYY-MM-DD.txt
```

## Input File Format

Your Excel file should have at least these columns:

| LED Change Date | Other Column | Another Column |
|----------------|--------------|----------------|
| 2025-01-15     | ...          | ...            |
| 2025-01-16     | ...          | ...            |
| 2025-01-17     | ...          | ...            |

**Important:**
- The date column MUST be named exactly: `LED Change Date`
- Dates must be in a date format Excel recognizes
- Other columns are ignored but can exist

## Example Output

```
=== Week 1.13.25 – 1.19.25 ===

Daily Report Ventura 1.17.25
Total: 5
Total to Date: 15

Daily Report Ventura 1.16.25
Total: 3
Total to Date: 10

------------------------------------------------------------

=== Week 1.6.25 – 1.12.25 ===

Daily Report Ventura 1.12.25
Total: 7
Total to Date: 7
```

## Customization

### Change the Date Column Name

If your column is named differently (e.g., "Change Date" instead of "LED Change Date"):

1. Open `Daily.py`
2. Find line ~25: `DEFAULT_DATE_COL = "LED Change Date"`
3. Change to your column name: `DEFAULT_DATE_COL = "Change Date"`
4. Save and run

### Use Different Input/Output Folders

Edit `config/config.py` and update:

```python
DAILY_REPORT_SETTINGS = {
    "input_folder": os.path.join(DATA_PATHS["input"], "daily_report_input"),
    "output_folder": os.path.join(DATA_PATHS["output"], "daily_report_output"),
}
```

## Troubleshooting

### "No files found in Input matching pattern"

- Make sure your Excel file is in: `scripts/Daily Report/Input/`
- Check the filename ends with `.xlsx`, `.xls`, or `.xlsm`
- Make sure you're not editing the file in Excel (close it first)

### "Date column 'LED Change Date' not found"

- Your Excel file doesn't have a column named `LED Change Date`
- Check the exact spelling (it's case-sensitive)
- Or update `DEFAULT_DATE_COL` in the script

### "No rows with valid dates found"

- The dates in your Excel file aren't recognized as dates
- Excel might have them as text
- Try converting the column to Date format in Excel before running

### Script runs but no output file created

- Check the `Output` folder has write permissions
- Make sure there's enough disk space
- Try creating a file manually in the Output folder to test permissions

## Advanced Usage

### Running with Custom Parameters

You can modify the defaults by editing the script, or use the interactive prompts to override them each time.

### Filtering by Region

To only include records for a specific region:

1. Open `Daily.py`
2. In the `main()` function, find the `SimpleNamespace` around line 220
3. Uncomment and modify:
   ```python
   filter_col="Region",        # Column to filter by
   filter_val="Ventura",       # Value to match
   ```

### Handling Multiple Excel Files

The script automatically selects the most recently modified Excel file in the Input folder. To use a specific file:

1. Delete or move other files from the Input folder
2. Keep only the file you want to use
3. Run the script

## File Details

- **Daily.py** - Main script
- **Input/** - Where you put your Excel files
- **Output/** - Where reports are saved
- **README.md** - This file

## Dependencies

- pandas (for Excel reading)
- openpyxl (for Excel file handling)

These are included in the main `requirements.txt`.

---

**Questions or issues?** Check the main README.md in the parent directory.
