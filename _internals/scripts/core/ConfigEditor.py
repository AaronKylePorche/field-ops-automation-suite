"""
Configuration Editor - Interactive Setup for New Users
========================================================
This script guides users through setting up all required configuration settings.
It reads the current config.py, prompts for updates, and saves changes back.

Users simply type values or press Enter to keep defaults.
All file paths will be cleaned of quotes automatically.
"""

import sys
import os
import re
from pathlib import Path

# Add parent directories to path for imports
# ConfigEditor is at: _internals/scripts/core/ConfigEditor.py
# Need to go up 3 levels to reach ROOT
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))
CONFIG_FILE = os.path.join(BASE_DIR, "_internals", "config", "config.py")

# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

def clean_path(path_str):
    """Remove quotes from file paths (from 'Copy as path' in Windows)"""
    return path_str.strip().strip('"').strip("'")

def read_current_config():
    """Read current config.py and extract values"""
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            content = f.read()
        return content
    except Exception as e:
        print(f"ERROR: Could not read config file: {e}")
        return None

def extract_value(content, pattern):
    """Extract a value from config content using regex"""
    match = re.search(pattern, content, re.DOTALL)
    if match:
        return match.group(1).strip()
    return None

def extract_list_value(content, pattern):
    """Extract a list value from config content"""
    match = re.search(pattern, content, re.DOTALL)
    if match:
        list_str = match.group(1)
        # Parse Python list syntax
        # Simple extraction: get all quoted strings
        items = re.findall(r'"([^"]*)"', list_str)
        return items
    return []

def prompt_user(question, default_value, allow_empty=False):
    """Prompt user and return input or default"""
    if default_value:
        prompt_text = f"{question}\n  [Default: {default_value}]\n  > "
    else:
        prompt_text = f"{question}\n  > "

    response = input(prompt_text).strip()

    if not response:
        if allow_empty:
            return None
        return default_value

    return response

def update_config_value(content, key_path, new_value):
    """Update a configuration value in the config file content"""
    # This is a simplified version - we'll use regex to find and replace values
    # For nested dicts like OUTLOOK_SETTINGS["mailbox_name"], we need careful handling

    return content

# ============================================================================
# MAIN CONFIGURATION FLOW
# ============================================================================

def main():
    """Main configuration editing flow"""

    print("\n" + "="*70)
    print("  ⚙️  KD ASSISTANT - CONFIGURATION EDITOR")
    print("  Interactive Setup for New Users")
    print("="*70)

    # Read current config
    config_content = read_current_config()
    if not config_content:
        print("\nERROR: Could not load current configuration.")
        input("Press Enter to exit...")
        return

    # Extract current values
    current_username = extract_value(config_content, r'USER_NAME\s*=\s*"([^"]*)"')
    current_mailbox = extract_value(config_content, r'"mailbox_name":\s*"([^"]*)"')
    current_excel = extract_value(config_content, r'"excel_path":\s*r?"([^"]*)"')
    current_district = extract_value(config_content, r'"district_xlsx":\s*r?"([^"]*)"')
    current_packing = extract_value(config_content, r'"packing_slips_root":\s*r?"([^"]*)"')
    current_whitelist = extract_list_value(config_content, r'"whitelist_emails":\s*\[(.*?)\]')
    current_folder_path = extract_list_value(config_content, r'"target_folder_path":\s*\[(.*?)\]')

    print("\n" + "-"*70)
    print("STEP 1: USER INFORMATION")
    print("-"*70)

    # 1. USERNAME
    print("\n[1/10] What is your name? (for logs and reports)")
    username = prompt_user(
        "Enter your name",
        current_username or "AKP"
    )

    # 2. OUTLOOK MAILBOX
    print("\n" + "-"*70)
    print("[2/10] OUTLOOK MAILBOX NAME (REQUIRED)")
    print("-"*70)
    print("""
This is YOUR email address in Outlook. This is CRITICAL for Email_Scanner
to connect to your mailbox and process claims.

Example: john.doe@company.com
         jane.smith@yourorganization.com
    """)
    mailbox = prompt_user(
        "Enter your Outlook mailbox name (email address)",
        current_mailbox or ""
    )

    if not mailbox:
        print("\nERROR: Mailbox name is required!")
        input("Press Enter to exit...")
        return

    # 3. OPENAI API KEY (from .env file)
    print("\n" + "-"*70)
    print("[3/10] OPENAI API KEY (Optional - stored in .env file)")
    print("-"*70)
    print(f"""
WOP22 uses OpenAI's API to extract claim information from emails.

Your API key is stored in a .env file (not in this config).

WHERE THE SCRIPTS LOOK FOR IT:
Scripts automatically load the .env file from the PROJECT ROOT folder:
   {BASE_DIR}\\.env

HOW TO CREATE AND ADD YOUR API KEY:
1. In the project root folder, create a new file named: .env
   (Note: It starts with a dot. In Windows, you may need to:
    - Create it as "temp.txt", then rename to ".env")

2. Open .env in a text editor and add this line:
   OPENAI_API_KEY=sk-your-actual-api-key-here

3. Save the file

4. Scripts will automatically read it when they run

EXAMPLE of what .env should contain:
   OPENAI_API_KEY=sk-proj-abc123xyz789...

If you don't have an OpenAI API key yet:
   https://platform.openai.com/api-keys

Note: The .env file is HIDDEN by default in Windows.
To see it later, enable "Show hidden files" in File Explorer.
    """)
    input("Press Enter to continue...")

    # 4. EXCEL FILE PATH (WOP22)
    print("\n" + "-"*70)
    print("[4/10] EXCEL FILE PATH (REQUIRED - WOP22 Claims Workbook)")
    print("-"*70)
    print("""
This is your main claims tracking Excel file (e.g., "CHP 3.0.xlsx").

HOW TO GET THE PATH:
1. Open File Explorer
2. Find your Excel file (e.g., on OneDrive or local drive)
3. Right-click the file → "Copy as path"
4. Paste the path here (it will include quotes, which we'll remove)

Example: "C:\\Users\\John\\OneDrive\\CHP 3.0.xlsx"
         "D:\\OneDrive - COMPANY\\Claims\\CHP 3.0.xlsx"
    """)
    excel_path = prompt_user(
        "Paste the Excel file path",
        current_excel or ""
    )

    if not excel_path:
        print("\nERROR: Excel file path is required!")
        input("Press Enter to exit...")
        return

    # Clean the path (remove quotes)
    excel_path = clean_path(excel_path)

    # 4.5 DISTRICT.XLSX FILE (WOP22 & SlipSender mapping file)
    print("\n" + "-"*70)
    print("[4.5/10] DISTRICT MAPPING FILE (REQUIRED - WOP22 & SlipSender)")
    print("-"*70)
    print("""
This is the District.xlsx file that maps:
  - Email addresses to districts (for WOP22 claim processing)
  - Districts to email addresses (for SlipSender packing slip emails)

The file structure:
  Column A: District code/name
  Columns B-K: Email addresses for that district

This file is typically stored in the SAME FOLDER as your other Excel files
(like your claims workbook and Material Received Worksheet).

HOW TO GET THE PATH:
1. Open File Explorer
2. Find your District.xlsx file
3. Right-click the file → "Copy as path"
4. Paste the path here

Example: "C:\\Users\\YourName\\Documents\\District.xlsx"
         "C:\\Users\\John\\OneDrive\\District.xlsx"
    """)
    district_xlsx = prompt_user(
        "Paste the District.xlsx file path",
        current_district or ""
    )

    if not district_xlsx:
        print("\nERROR: District.xlsx path is required!")
        input("Press Enter to exit...")
        return

    # Clean the path
    district_xlsx = clean_path(district_xlsx)

    # 5. PACKING SLIPS ROOT FOLDER
    print("\n" + "-"*70)
    print("[5/10] PACKING SLIPS FOLDER (REQUIRED - SlipSender)")
    print("-"*70)
    print("""
This is where your packing slip PDFs are stored. Specifically, there should
be an "Unprocessed" subfolder where you place slips ready to process.

FOLDER STRUCTURE:
  Your Packing Slips Root/
    └── Unprocessed/           ← Put unsent slips here
        ├── District_A_slips.pdf
        ├── District_B_slips.pdf
        └── ...

When you're ready to send emails, run "Draft Packing Slips Emails" script
which will automatically create Outlook draft emails with these attachments.

You can customize which districts get which email addresses by editing:
  scripts/WOP/District.xlsx

HOW TO GET THE PATH:
1. Open File Explorer
2. Navigate to your "Packing Slips" or "Unprocessed" folder
3. Right-click the "Unprocessed" folder → "Copy as path"
4. Paste the path here

Example: "C:\\Users\\YourName\\Desktop\\Packing Slips\\Unprocessed"
         "D:\\OneDrive - Company\\Packing Slips\\Unprocessed"
    """)
    packing_path = prompt_user(
        "Paste the Packing Slips 'Unprocessed' folder path",
        current_packing or ""
    )

    if not packing_path:
        print("\nERROR: Packing slips folder path is required!")
        input("Press Enter to exit...")
        return

    # Clean the path
    packing_path = clean_path(packing_path)

    # 5.5 MATERIAL RECEIVED WORKSHEET (SlipSender workbook)
    current_workbook = extract_value(config_content, r'"workbook_path":\s*r?"([^"]*)"')

    print("\n" + "-"*70)
    print("[5.5/10] MATERIAL RECEIVED WORKSHEET (REQUIRED - SlipSender)")
    print("-"*70)
    print("""
This is your "Material Received Worksheet" Excel file that tracks work orders
and which districts they belong to. SlipSender uses this to draft packing slip
emails for the correct districts.

Example file: "Material Received Worksheet.xlsx"

HOW TO GET THE PATH:
1. Open File Explorer
2. Find your Material Received Worksheet Excel file
3. Right-click the file → "Copy as path"
4. Paste the path here

Example: "C:\\Users\\YourName\\OneDrive\\Material Received Worksheet.xlsx"
         "D:\\OneDrive - Company\\Worksheets\\Material Received Worksheet.xlsx"
    """)
    workbook_path = prompt_user(
        "Paste the Material Received Worksheet path",
        current_workbook or ""
    )

    if not workbook_path:
        print("\nERROR: Material Received Worksheet path is required!")
        input("Press Enter to exit...")
        return

    # Clean the path
    workbook_path = clean_path(workbook_path)

    # 5.7 MERGED PDF SAVE DIRECTORY (DocumentCombiner)
    current_save_dir = extract_value(config_content, r'"save_dir":\s*r?"([^"]*)"')

    print("\n" + "-"*70)
    print("[5.7/10] MERGED PDFs SAVE DIRECTORY (REQUIRED - DocumentCombiner)")
    print("-"*70)
    print("""
When you combine/merge Claims and Permits documents, the final merged PDFs
are saved to a folder you specify. This is your destination for merged files.

Example folder: "Final Worksheet_Images" or "OneDrive/Final Documents"

HOW TO GET THE PATH:
1. Open File Explorer
2. Find or create a folder where you want merged PDFs saved
3. Right-click the folder → "Copy as path"
4. Paste the path here

Example: "C:\\Users\\YourName\\OneDrive\\Final Worksheet_Images"
         "D:\\OneDrive - Company\\Final Documents"
    """)
    save_dir = prompt_user(
        "Paste the folder path where merged PDFs should be saved",
        current_save_dir or ""
    )

    if not save_dir:
        print("\nERROR: PDF save directory is required!")
        input("Press Enter to exit...")
        return

    # Clean the path
    save_dir = clean_path(save_dir)

    # 6. EMAIL WHITELIST
    print("\n" + "-"*70)
    print("[6/10] EMAIL WHITELIST (Who can trigger claim processing)")
    print("-"*70)
    print("""
Only emails from addresses in this whitelist will trigger the claim processor.
Current whitelist: {}

Enter email addresses separated by commas, or press Enter to keep defaults.

Example: john@company.com, jane@company.com, claims@company.com
    """.format(", ".join(current_whitelist) if current_whitelist else "None set"))

    whitelist_input = prompt_user(
        "Enter email addresses (comma-separated)",
        ", ".join(current_whitelist) if current_whitelist else "your.email@company.com",
        allow_empty=True
    )

    if whitelist_input:
        whitelist = [email.strip() for email in whitelist_input.split(",")]
    else:
        whitelist = current_whitelist or []

    # 7. OUTLOOK FOLDER PATHS
    print("\n" + "-"*70)
    print("[7/10] OUTLOOK FOLDER STRUCTURE (Where to monitor for claims)")
    print("-"*70)
    print("""
The Email_Scanner watches specific folders in Outlook for incoming claims.
You can use the default structure or create your own.

DEFAULT STRUCTURE (recommended for new users):
  Inbox
    └── KD Assistant
        ├── Claims              ← Claims go here
        ├── Permits             ← Permits go here
        ├── Processed           ← Processed items archive
        └── Queue               ← Temp folder for tickets

INSTRUCTIONS TO CREATE THIS STRUCTURE IN OUTLOOK:
1. Open Outlook
2. Right-click "Inbox" → "New Folder..."
3. Create folder: "KD Assistant"
4. Inside KD Assistant, create these 4 subfolders:
   - "Claims"
   - "Permits"
   - "Processed"
   - "Queue"

Your folder structure in config will be:
  ["Inbox", "KD Assistant", "Claims"]

This means: Inbox > KD Assistant > Claims

Keep defaults (press Enter) or enter custom path (e.g., "Inbox,Custom Folder,Claims")
    """)

    folder_input = prompt_user(
        "Enter Outlook folder path (comma-separated)",
        ", ".join(current_folder_path) if current_folder_path else "Inbox, KD Assistant, Claims",
        allow_empty=True
    )

    if folder_input:
        folder_path = [f.strip() for f in folder_input.split(",")]
    else:
        folder_path = current_folder_path or ["Inbox", "KD Assistant", "Claims"]

    # 8. INPUT/OUTPUT FOLDERS (informational only)
    print("\n" + "-"*70)
    print("[8/10] INPUT/OUTPUT FOLDERS (Informational)")
    print("-"*70)
    print("""
Scripts automatically use these folders for input/output:

INPUT FOLDER:
  data/input/
  └── Where scripts look for files to process
      Examples:
      - Daily Report script looks for LED change files here
      - JIS Automation looks for Excel sheets in: data/input/jis_input/

OUTPUT FOLDER:
  data/output/
  └── Where scripts save results
      Examples:
      - Daily Report saves reports here: data/output/daily_report_output/
      - JIS Automation saves sheets here: data/output/jis_output/

These folders are automatically created when you first run the scripts.
You should NOT change these unless you have a specific reason.
    """)

    input("Press Enter to continue...")

    # SUMMARY & SAVE
    print("\n" + "="*70)
    print("  CONFIGURATION SUMMARY")
    print("="*70)

    print(f"""
USER NAME:                      {username}
OUTLOOK MAILBOX:                {mailbox}
OPENAI API KEY:                 [Stored in .env file]
EXCEL FILE (WOP22):             {excel_path}
DISTRICT.XLSX (WOP22/SlipSender): {district_xlsx}
PACKING SLIPS FOLDER:           {packing_path}
MATERIAL RECEIVED WORKSHEET:    {workbook_path}
MERGED PDFs SAVE DIRECTORY:     {save_dir}
EMAIL WHITELIST:                {", ".join(whitelist) if whitelist else "None"}
OUTLOOK FOLDER PATH:            {" > ".join(folder_path)}

    """)

    confirm = input("Save these changes? (yes/no): ").strip().lower()

    if confirm not in ['yes', 'y']:
        print("\nChanges cancelled. Configuration not updated.")
        input("Press Enter to exit...")
        return

    # Save configuration
    if save_configuration(config_content, username, mailbox,
                         excel_path, district_xlsx, packing_path, workbook_path, save_dir, whitelist, folder_path):
        print("\n✅ Configuration saved successfully!")
        print(f"   File: {CONFIG_FILE}")
        print("\nYou can now run the main KD Assistant launcher.")
    else:
        print("\n❌ ERROR: Failed to save configuration.")

    input("Press Enter to exit...")

def save_configuration(content, username, mailbox, excel_path, district_xlsx,
                      packing_path, workbook_path, save_dir, whitelist, folder_path):
    """Save all configuration changes back to config.py"""

    try:
        # Escape backslashes in paths for safe regex replacement
        excel_path_escaped = excel_path.replace("\\", "\\\\")
        district_xlsx_escaped = district_xlsx.replace("\\", "\\\\")
        packing_path_escaped = packing_path.replace("\\", "\\\\")
        workbook_path_escaped = workbook_path.replace("\\", "\\\\")
        save_dir_escaped = save_dir.replace("\\", "\\\\")

        # Update USER_NAME
        content = re.sub(
            r'USER_NAME\s*=\s*"[^"]*"',
            f'USER_NAME = "{username}"',
            content
        )

        # Update mailbox_name (appears in multiple settings sections)
        content = re.sub(
            r'("mailbox_name":\s*)"[^"]*"',
            f'\\1"{mailbox}"',
            content
        )

        # Update excel_path (handle raw string)
        content = re.sub(
            r'("excel_path":\s*)r?"[^"]*"',
            f'\\1r"{excel_path_escaped}"',
            content
        )

        # Update district_xlsx in WOP22_SETTINGS
        content = re.sub(
            r'("district_xlsx":\s*)r?"[^"]*"',
            f'\\1r"{district_xlsx_escaped}"',
            content
        )

        # Update district_map_file in SLIPSENDER_SETTINGS (same file, different config key)
        content = re.sub(
            r'("district_map_file":\s*)r?"[^"]*"',
            f'\\1r"{district_xlsx_escaped}"',
            content
        )

        # Update packing_slips_root
        content = re.sub(
            r'("packing_slips_root":\s*)r?"[^"]*"',
            f'\\1r"{packing_path_escaped}"',
            content
        )

        # Update workbook_path (Material Received Worksheet)
        content = re.sub(
            r'("workbook_path":\s*)r?"[^"]*"',
            f'\\1r"{workbook_path_escaped}"',
            content
        )

        # Update save_dir (Document Combiner merged PDF save directory)
        content = re.sub(
            r'("save_dir":\s*)r?"[^"]*"',
            f'\\1r"{save_dir_escaped}"',
            content
        )

        # Update whitelist_emails
        whitelist_str = ",\n        ".join([f'"{email}"' for email in whitelist])
        content = re.sub(
            r'("whitelist_emails":\s*\[)[^\]]*(\])',
            f'\\1\n        {whitelist_str}\n    \\2',
            content
        )

        # Update target_folder_path
        folder_str = ",\n        ".join([f'"{folder}"' for folder in folder_path])
        content = re.sub(
            r'("target_folder_path":\s*\[)[^\]]*(\])',
            f'\\1\n        {folder_str}\n    \\2',
            content
        )

        # Write back to config file
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            f.write(content)

        return True

    except Exception as e:
        print(f"ERROR saving configuration: {e}")
        return False

# ============================================================================
# ENTRY POINT
# ============================================================================

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nConfiguration cancelled.")
        sys.exit(0)
    except Exception as e:
        print(f"\nERROR: {e}")
        import traceback
        traceback.print_exc()
        input("Press Enter to exit...")
        sys.exit(1)
