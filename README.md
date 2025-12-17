# Field Ops Automation Suite

**Windows Automation Suite for Claims Processing, Document Management, and Reporting**

A configuration-driven automation framework designed for streamlining workflow operations, featuring email monitoring, AI-powered data extraction, report generation, and background process management.

## Features

### Core Automation
- **Claims Processing** - Automated claim tracking and Excel workbook updates
- **Document Merging** - Combine email attachments (PDFs, images, documents) into single files
- **Email Monitoring** - Event-driven Outlook integration with whitelist-based filtering
- **AI-Powered Extraction** - OpenAI integration for intelligent claim data parsing

### Report Generation
- **Daily LED Change-Out Reports** - Automated reporting for lighting replacement projects
- **New Development Reports** - Filtered data reports with automated email drafts
- **Job Instruction Sheets (JIS)** - Automated generation from Excel templates with LED/HPSV/MH conversion

### Background Services
- **Claim Watcher Suite** - Unified monitoring system with:
  - Outlook process detection (WMI + polling hybrid)
  - Automated email scanning when Outlook is running
  - Queue-based asynchronous ticket processing
  - System keep-alive functionality
- **Packing Slip Manager** - District-based email drafting with attachment routing

## System Requirements

- **Operating System:** Windows 10 or Windows 11
- **Python:** 3.12 or higher
- **Microsoft Outlook:** Installed and configured with an active mailbox
- **Microsoft Excel:** 2016+ or Microsoft 365

## Installation

### 1. Clone the Repository

```bash
git clone https://github.com/yourusername/kd-assistant.git
cd kd-assistant
```

### 2. Run Setup Script (IMPORTANT: Read Carefully!)

The `setup.bat` script will install Python dependencies automatically:

```bash
setup.bat
```

**⚠️ CRITICAL: First-Time Installation Requirement**

If this is your **first time installing** the dependencies:

1. **Run `setup.bat`** - This installs Python packages and dependencies
2. **Restart your computer** - This is REQUIRED for PATH updates to take effect
3. **Run `setup.bat` again** - This verifies all dependencies installed correctly

**Why restart?** Windows needs to reload environment variables (especially PATH) after installing Python packages. Without a restart, some dependencies may not be accessible to the scripts.

**Note:** After the first installation cycle (run → restart → run again), you won't need to restart for subsequent updates.

### 3. Configure Environment Variables

Copy the `.env.template` file to `.env` and add your API keys:

```bash
cp _internals\.env.template _internals\.env
```

Edit `_internals\.env` and add your credentials:
```
OPENAI_API_KEY=your-openai-api-key-here
CLIENT_ID=your-azure-client-id-here
TENANT_ID=your-azure-tenant-id-here
```

Get your OpenAI API key from: https://platform.openai.com/api-keys

### 4. Configure Settings

The `_internals\config\config.py` file contains placeholder values that you must customize.

**IMPORTANT:** Before running the application, edit `config.py` and update the following:

#### Required Settings:

**User Information:**
```python
USER_NAME = "YourName"  # Replace with your name
```

**Email Settings:**
```python
OUTLOOK_SETTINGS = {
    "mailbox_name": "your.email@company.com",  # Your Outlook email
    "whitelist_emails": [
        "colleague1@company.com",
        "colleague2@company.com"
    ],
    # Outlook folder path where claims are stored
    "target_folder_path": ["Inbox", "KD Assistant", "Claims"]
}
```

**File Paths:**
```python
WOP22_SETTINGS = {
    "excel_path": r"C:\Path\To\Your\Tracker.xlsx",
    "district_xlsx": r"C:\Path\To\District.xlsx",
}

DOCUMENT_COMBINER_SETTINGS = {
    "save_dir": r"C:\Path\To\Output\Folder",
}

SLIPSENDER_SETTINGS = {
    "workbook_path": r"C:\Path\To\Material_Worksheet.xlsx",
    "packing_slips_root": r"C:\Path\To\PackingSlips\Folder",
}
```

### 5. Customize Sample Data Files

The following data files contain placeholder values that you should customize:

**District Mapping** (`_internals\scripts\WOP\District.xlsx`):
- Maps sender email addresses to district codes
- Edit this file to add your actual district mappings

**JIS Setup** (`_internals\scripts\JIS Automation\Setup.xlsx`):
- Contains job header text and structure exclusions
- Edit this file to add your job location information

These files are tracked in git with safe placeholder values. Customize them for your environment.

## Usage

### Launch the Menu Interface

Run the main launcher to access all tools:

```bash
python "launcher.py"
```

### Menu Options

```
╔═════════════════════════════════════════════════════╗
║          KD ASSISTANT - Main Menu                    ║
╚═════════════════════════════════════════════════════╝

[1] Add Claims To Tracker
[2] 🧩 Combine Claims & Permits
[3] 💡 Run Daily LED Change Out Report
[4] 📊 Generate New Development Report
[5] Run JIS Automation
[6] Create Master JIS File
[7] Launch Claim Watcher Suite (Background Services)
[8] Draft Packing Slips Emails

[q] ❌ Quit
```

### Running Scripts Individually

All scripts can also be run standalone:

```bash
python "_internals\scripts\core\Stand_Alone_Processor.py"
python "_internals\scripts\Daily Report\Daily.py"
python "_internals\scripts\JIS Automation\JIS.py"
```

## Architecture Overview

### Configuration-Driven Design

All customization happens in a single file (`_internals/config/config.py`), making the system portable across drives and installations. Scripts are read-only consumers of configuration.

### Project Structure

```
field-ops-automation-suite/
├── launcher.py               # Main launcher
├── setup.bat                 # Dependency installer
├── CLAUDE.md                 # Developer documentation
├── requirements.txt          # Python dependencies
├── input/                    # User-provided data files
├── output/                   # Generated reports and logs
└── _internals/
    ├── config/
    │   └── config.py         # Master configuration
    ├── scripts/
    │   ├── core/             # Email, claims, reports
    │   ├── monitoring/       # Background services
    │   ├── JIS Automation/   # Job instruction sheets
    │   ├── Daily Report/     # LED reporting
    │   ├── Document Processing/
    │   └── WOP/              # AI claim extraction
    └── data/
        └── templates/        # Icons, samples
```

### Key Patterns

**Universal Config Import:**
```python
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent / "config"))
import config

WORKBOOK_PATH = config.WOP22_SETTINGS["excel_path"]
```

**Excel Interaction:**
- Use COM automation (`win32com.client`) when files may be open
- Use `openpyxl` when files must be closed

**Outlook Integration:**
```python
import win32com.client as win32
outlook = win32.Dispatch("Outlook.Application")
# Navigate to folders using config paths
# Create draft emails (never auto-send)
```

## Adding Custom Scripts

1. Place your script in the appropriate `_internals/scripts/` subfolder
2. Import configuration using the universal pattern
3. Update `config.py` SCRIPTS dictionary:

```python
"9": {
    "name": "📊 Your Script Name",
    "path": os.path.join(BASE_DIR, "_internals", "scripts", "folder", "script.py"),
    "enabled": True,
    "description": "Brief description"
}
```

4. Run `launcher.py` to see your new menu option

## Development

For detailed development guidelines, architectural patterns, and code conventions, see [CLAUDE.md](CLAUDE.md).

### Key Development Principles

- **Configuration-Driven:** Never hardcode paths or credentials
- **Portable:** Auto-detects BASE_DIR for drive/system independence
- **Graceful:** Try COM → fallback to openpyxl for Excel operations
- **Secure:** Whitelist filtering, no auto-send emails
- **Modular:** Scripts organized by functionality

### Running Tests

```bash
# Test individual scripts
python "_internals\scripts\core\Email_Scanner.py"

# Test configuration
python "_internals\config\config.py"
```

## Contributing

Contributions are welcome! Please follow these guidelines:

1. Read [CLAUDE.md](CLAUDE.md) for architectural patterns
2. Follow the universal config import pattern
3. Never hardcode paths, emails, or credentials
4. Test on a clean Windows installation if possible
5. Submit pull requests with clear descriptions

## Security Best Practices

### Never Commit Sensitive Data

The `.gitignore` file excludes:
- `.env` (API keys)
- `config.py` (personal paths/emails)
- `input/` and `output/` folders (real data)
- Log files and Excel tracking sheets

### Email Safety

- Emails are created as **drafts only** - never auto-sent
- Whitelist filtering prevents processing untrusted senders
- Attachments are validated before processing

### API Key Management

- Store keys in `.env` file (never in code)
- Use environment variables: `os.getenv("OPENAI_API_KEY")`
- Rotate keys regularly
- Use separate keys for development/production

## Troubleshooting

### Common Issues

**"Module not found" errors after first installation:**

If you get this error immediately after installing:
1. **Did you restart your computer?** This is REQUIRED after the first `setup.bat` run
2. After restarting, run `setup.bat` again to verify installation
3. If still failing, manually check: `python --version` and `pip --version`

For other module errors (after initial setup):
```bash
# Re-run setup to install dependencies
setup.bat
```

**Outlook integration not working:**
- Ensure Outlook is installed and configured
- Check that your mailbox name matches `OUTLOOK_SETTINGS["mailbox_name"]`
- Verify folder paths in `target_folder_path`

**Excel COM errors:**
- Close Excel and try again
- Check file paths in `config.py`
- Ensure Excel is installed (not just Excel Viewer)

**OpenAI API errors:**
- Verify `.env` file exists and contains `OPENAI_API_KEY`
- Check API key validity at https://platform.openai.com/api-keys
- Ensure you have available API credits

### Debug Mode

Enable debug logging in `config.py`:
```python
DEBUG = True
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Built with [openpyxl](https://openpyxl.readthedocs.io/) for Excel manipulation
- Uses [pywin32](https://github.com/mhammond/pywin32) for Windows COM automation
- Powered by [OpenAI API](https://platform.openai.com/) for intelligent data extraction

---

**Made with care for workflow automation | Portable | Config-driven | Secure**
