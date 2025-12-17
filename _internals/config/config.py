"""
CONFIGURATION FILE FOR FIELD OPS AUTOMATION SUITE
===================================================

SETUP INSTRUCTIONS:
1. Replace all placeholder values with your actual paths and settings:
   - Email addresses (your.email@company.com, user1@company.com, etc.)
   - File paths (C:\Users\YourName\...)
   - User name
2. Copy .env.template to .env and add your API keys
3. Update District.xlsx and Setup.xlsx with your data

NOTE: This file is tracked in git with safe placeholder values.
Customize your local copy - git will not track your personal changes.

=============================================================================

Field Ops Automation Suite - Configuration File
================================================
Edit this file to customize the script launcher for your system.
All paths are relative to the installation folder.

INSTRUCTIONS:
1. Update USER_NAME to your actual name
2. Verify all file paths exist and are correct
3. Enable/disable scripts by changing ENABLED to True/False
4. Save and run launcher.py
"""

import os
from pathlib import Path

# ============================================================================
# SYSTEM CONFIGURATION
# ============================================================================

# Your name/identifier (displayed in menu and logs)
USER_NAME = "YourName"

# Automatically detect the base installation directory
# config.py is at _internals/config/config.py, so BASE_DIR = root (up 3 levels)
BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Enable debug mode for troubleshooting
DEBUG = False

# ============================================================================
# SCRIPT DEFINITIONS
# =============================================================================
# Each script entry has:
#   "key": (display_name, script_path, enabled)
#
# - key: Menu hotkey (number or letter)
# - display_name: What shows in the menu
# - script_path: Relative path to the script (from BASE_DIR)
# - enabled: True to show in menu, False to hide

SCRIPTS = {
    # ========================= CORE SCRIPTS =========================
    # Note: Configuration is now part of setup.bat for first-time users
    # Run setup.bat again if you need to reconfigure

    "1": {
        "name": "Add Claims To Tracker",
        "path": os.path.join(BASE_DIR, "_internals", "scripts", "core", "Stand_Alone_Processor.py"),
        "enabled": True,
        "description": "Process and add claims to the tracking system"
    },

    "2": {
        "name": "🧩 Combine Claims & Permits",
        "path": os.path.join(BASE_DIR, "_internals", "scripts", "Document Processing", "DocumentCombiner.py"),
        "enabled": True,
        "description": "Combine claims and permit documents"
    },

    # ========================= REPORTING SCRIPTS =========================

    "3": {
        "name": "💡 Run Daily LED Change Out Report",
        "path": os.path.join(BASE_DIR, "_internals", "scripts", "Daily Report", "Daily.py"),
        "enabled": True,
        "description": "Generate daily LED change out report"
    },

    "4": {
        "name": "📊 Generate New Development Report",
        "path": os.path.join(BASE_DIR, "_internals", "scripts", "core", "kd_report_generator.py"),
        "enabled": True,
        "description": "Generate filtered development report and create draft email"
    },

    # ========================= JIS SCRIPTS =========================

    "5": {
        "name": "Run JIS Automation",
        "path": os.path.join(BASE_DIR, "_internals", "scripts", "JIS Automation", "JIS.py"),
        "enabled": True,
        "description": "Run Job Instruction Sheet automation"
    },

    "6": {
        "name": "Create Master JIS File",
        "path": os.path.join(BASE_DIR, "_internals", "scripts", "JIS Automation", "MasterJIS.py"),
        "enabled": True,
        "description": "Create master JIS file"
    },

    # ========================= BACKGROUND SERVICES =========================
    # Claim Watcher Suite launches all services in a single unified window via wrapper.
    # The wrapper (ClaimWatcherSuite_Unified.py) manages:
    # - Outlook monitoring (detects when Outlook starts/stops)
    # - Email_Scanner (auto-launched when Outlook opens, stopped when Outlook closes)
    # - Ticket_Reader (always running, processes tickets from queue)
    # - keep_awake (always running, prevents system sleep)
    # Result: 1 unified window with prefixed output from all 4 services
    # Note: supervisor1.py is no longer used (replaced by unified wrapper)

    "7": {
        "name": "Launch Claim Watcher Suite",
        "enabled": True,
        "description": "Launch background claim monitoring services",
        "background_scripts": [
            os.path.join(BASE_DIR, "_internals", "scripts", "monitoring", "ClaimWatcherSuite_Unified.py"),
        ]
    },

    # ========================= EMAIL/COMMUNICATION =========================

    "8": {
        "name": "Draft Packing Slips Emails",
        "path": os.path.join(BASE_DIR, "_internals", "scripts", "core", "SlipSender.py"),
        "enabled": True,
        "description": "Draft and send packing slip emails"
    },

    # ========================= DISABLED SCRIPTS (not in original menu) =========================

    "ClaimPro1": {
        "name": "📥 Add New Claims to Excel",
        "path": os.path.join(BASE_DIR, "_internals", "scripts", "Document Processing", "ClaimPro1.py"),
        "enabled": False,
        "description": "Add new claims to Excel tracking file"
    },

    "PackingSlipScanner": {
        "name": "🔍 SCAN Slips Folder",
        "path": os.path.join(BASE_DIR, "_internals", "scripts", "Document Processing", "packing_slip_scanner.py"),
        "enabled": False,
        "description": "Scan and process packing slips folder"
    },

    # ========================= EXIT =========================

    "q": {
        "name": "❌ Quit",
        "path": None,
        "enabled": True,
        "description": "Exit the launcher"
    },
}

# ============================================================================
# DATA FOLDER CONFIGURATION
# ============================================================================
# Define where various data files are stored

DATA_PATHS = {
    "input": os.path.join(BASE_DIR, "input"),
    "output": os.path.join(BASE_DIR, "output"),
    "templates": os.path.join(BASE_DIR, "_internals", "data", "templates"),
    "packing_slips": os.path.join(BASE_DIR, "input", "packing_slips"),
    "excel_files": os.path.join(BASE_DIR, "input", "excel_files"),
}

# ============================================================================
# SCRIPT-SPECIFIC SETTINGS
# ============================================================================
# Add script-specific configuration here if needed

JIS_SETTINGS = {
    # Input folder containing Excel files with LED change data (USER PROVIDES THESE)
    "input_folder": os.path.join(DATA_PATHS["input"], "jis_input"),

    # Output folder where JIS workbooks are saved
    "output_folder": os.path.join(DATA_PATHS["output"], "jis_output"),

    # Log folder for preflight and final discrepancy logs
    "log_folder": os.path.join(DATA_PATHS["output"], "jis_logs"),

    # JIS Template file (script dependency - bundled with script)
    "template_file": os.path.join(BASE_DIR, "_internals", "scripts", "JIS Automation", "JIS Template.xlsx"),

    # Setup file with header info (A2), AF5 value (B2), AO35 date (C2), and exclusions (D column)
    # (script dependency - bundled with script)
    "setup_file": os.path.join(BASE_DIR, "_internals", "scripts", "JIS Automation", "Setup.xlsx"),

    # Conversion chart for LED/HPSV/MH mappings (script dependency - bundled with script)
    "conversion_chart_folder": os.path.join(BASE_DIR, "_internals", "scripts", "JIS Automation"),
    "conversion_chart_file": os.path.join(BASE_DIR, "_internals", "scripts", "JIS Automation", "ConversionChart.xlsx"),
}

DAILY_REPORT_SETTINGS = {
    "input_folder": os.path.join(DATA_PATHS["input"], "daily_report_input"),
    "output_folder": os.path.join(DATA_PATHS["output"], "daily_report_output"),
}

TIMESHEET_SETTINGS = {
    "input_folder": os.path.join(DATA_PATHS["input"], "timesheet_input"),
    "output_folder": os.path.join(DATA_PATHS["output"], "timesheet_output"),
    "time_clock_file": os.path.join(DATA_PATHS["input"], "Time Clock.xlsx"),
}

# ============================================================================
# OUTLOOK CONFIGURATION (for Email_Scanner & Claim Watcher Suite)
# ============================================================================
# These settings configure how the Email Scanner monitors your Outlook account
# Users MUST customize these for their specific email setup

OUTLOOK_SETTINGS = {
    # Email addresses that trigger claim processing
    # Add YOUR email address and any team members who send claims
    # Format: ["email@company.com", "another@company.com"]
    "whitelist_emails": [
        "user1@company.com",
        "user2@company.com"
    ],

    # Folder path in Outlook where claims are stored
    # Format: ["Inbox", "FolderName", "SubfolderName"]
    # This should match your actual Outlook folder structure
    # Example: ["Inbox", "Claims", "New Claims"] for Inbox > Claims > New Claims
    "target_folder_path": [
        "Inbox",
        "KD Assistant",
        "Claims"
    ],

    # Queue folder where tickets are created (for Ticket_Reader to process)
    # This is auto-detected relative to BASE_DIR, but can be customized
    "queue_folder": os.path.join(BASE_DIR, "_internals", "queue"),

    # Primary mailbox name for WOP22.py and claim processing
    # Change to YOUR email address
    "mailbox_name": "your.email@company.com",
}

# ============================================================================
# WOP22 CONFIGURATION (for Claim Watcher Suite - Email Processing)
# ============================================================================
# WOP22 (Workflow Operations Processor) reads email claims and writes to Excel

WOP22_SETTINGS = {
    # Path to the master claims Excel workbook
    # CRITICAL: Update this to YOUR OneDrive path or local Excel file location
    # This file contains the main tracking table (MasterTable)
    "excel_path": r"C:\Users\YourName\Documents\CHP\CHP 2.5.xlsx",  # ← USER WILL SET THIS

    # Sheet name and table name in the Excel workbook
    "sheet_name": "MainSheet",
    "table_name": "MasterTable",
    "table_claim_col_index": 4,  # Column D (1-based index within table row)

    # District mapping file (maps sender email to district)
    "district_xlsx": r"C:\Users\YourName\Documents\CHP\District.xlsx",  # ← USER WILL SET THIS

    # OpenAI configuration
    "model_name": "gpt-3.5-turbo",
    "timeout_seconds": 30,

    # Source mode for MEDIUM/LOW confidence claims:
    # "auto"  -> prefer tasking sheet; fallback to email values when sheet is blank
    # "email" -> always use email-parsed values
    # "sheet" -> always use tasking sheet values
    # "prompt"-> ask interactively per email
    "medium_low_source_mode": "auto",

    # Safety toggles
    "dry_run": False,
    "limit_n": None,  # Process all emails if None, or limit to N emails
    "always_leave_workbook_open": True,
    "always_leave_excel_running": True,
}

# ============================================================================
# DOCUMENT COMBINER CONFIGURATION
# ============================================================================
# DocumentCombiner.py merges email attachments (PDFs, images, documents) into single PDFs

DOCUMENT_COMBINER_SETTINGS = {
    # Directory where final merged PDFs are saved (USER-CONFIGURED)
    "save_dir": r"C:\Users\YourName\Documents\CHP\Images",  # ← USER WILL SET THIS

    # Email folder for Claims documents (uses same hierarchy as Claim Processor)
    # This references OUTLOOK_SETTINGS["target_folder_path"] - configured once, used by both
    "claims_folder_path": OUTLOOK_SETTINGS["target_folder_path"],

    # Email folder for Permits documents (same hierarchy, but with "Permits" instead of "Claims")
    # Auto-derived: replaces last folder in claims path with "Permits"
    "permits_folder_path": (
        OUTLOOK_SETTINGS["target_folder_path"][:-1] + ["Permits"]
        if len(OUTLOOK_SETTINGS["target_folder_path"]) > 0
        else ["Inbox", "KD Assistant", "Permits"]
    ),

    # Mailbox name for document processing (same as Claim Processor)
    "mailbox_name": OUTLOOK_SETTINGS["mailbox_name"],

    # Final destination folders after processing (where processed emails are moved)
    "claims_dest_folder": ["Inbox", "CLAIMS"],
    "permits_dest_folder": ["Inbox", "PERMITS"],
}

# ============================================================================
# MASTERJIIS CONFIGURATION
# ============================================================================
# MasterJIS.py combines all JIS output files into a single master workbook

MASTERJIS_SETTINGS = {
    # Location of JIS output files to consolidate
    "jis_output_folder": os.path.join(DATA_PATHS["output"], "jis_output"),

    # Where the master consolidated file is saved
    "master_output_folder": os.path.join(DATA_PATHS["output"], "jis_output"),

    # Setup file to read job location name for master filename
    "setup_file": os.path.join(DATA_PATHS["input"], "jis_input", "Setup.xlsx"),
}

# ============================================================================
# SLIPSENDER CONFIGURATION
# ============================================================================
# SlipSender.py creates Outlook drafts with packing slip attachments per district

SLIPSENDER_SETTINGS = {
    # Material Received Worksheet (tracks work orders and districts)
    # CRITICAL: Update to YOUR OneDrive or local path
    "workbook_path": r"C:\Users\YourName\Documents\CHP\Material Received Worksheet.xlsx",  # ← USER WILL SET THIS
    "worksheet_name": "Recieved Materials Template",

    # District mapping file (maps district code to email addresses)
    "district_map_file": r"C:\Users\YourName\Documents\CHP\District.xlsx",  # ← USER WILL SET THIS

    # Root folder where packing slips are stored
    # CRITICAL: Update to YOUR packing slips folder path
    "packing_slips_root": r"C:\Documents\PackingSlips\Unprocessed",  # ← USER WILL SET THIS

    # Allowed file extensions for packing slips
    "allowed_extensions": {".pdf", ".jpg", ".jpeg"},

    # Email addresses to always CC on packing slip drafts
    "always_cc": "user1@company.com; user2@company.com",

    # Log folder for tracking packing slip emails
    "log_folder": os.path.join(BASE_DIR, "_internals", "scripts", "core", "Logs"),
    "log_filename": "PackingSlipEmailLog.xlsx",

    # HTML table styling
    "yellow_hex": "#FFF2CC",         # Color for marking processed rows
    "missing_row_bg": "#FDECEA",     # Light red for missing packing slips in table
}

# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

def get_enabled_scripts():
    """Return only scripts that are enabled"""
    return {k: v for k, v in SCRIPTS.items() if v.get("enabled", False)}

def validate_paths():
    """Check if all required paths exist. Returns list of missing paths."""
    missing = []

    for key, script_config in get_enabled_scripts().items():
        if "path" in script_config and script_config["path"]:
            if not os.path.exists(script_config["path"]):
                missing.append(script_config["path"])

        if "background_scripts" in script_config:
            for script_path in script_config["background_scripts"]:
                if not os.path.exists(script_path):
                    missing.append(script_path)

    return missing

def print_config_info():
    """Print configuration summary for debugging"""
    if DEBUG:
        print(f"Base Directory: {BASE_DIR}")
        print(f"User: {USER_NAME}")
        print(f"Enabled Scripts: {len(get_enabled_scripts())}")
        missing = validate_paths()
        if missing:
            print("\nWARNING - Missing script files:")
            for m in missing:
                print(f"  - {m}")
