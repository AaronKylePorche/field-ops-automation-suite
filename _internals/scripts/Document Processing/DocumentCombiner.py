# DocumentCombiner.py - Email Attachment Merger
# Refactored for New User Package (Configuration-driven)
#
# Merges email attachments (PDFs, images, documents) into single PDF files
# Processes Claims and Permits folders separately
# All settings come from config.py - No hardcoding!
#
# Features:
# - Extracts claim numbers using OpenAI
# - Converts documents (DOCX, XLSX, images) to PDF
# - Merges all attachments into single PDFs
# - Handles nested .msg files (forwards) recursively
# - Tags emails and moves them to destination folders
#
# Run: python DocumentCombiner.py

import os
import sys
import win32com.client
from openai import OpenAI
from PyPDF2 import PdfReader, PdfWriter
from PIL import Image
from docx2pdf import convert as convert_docx
from datetime import datetime
import tempfile
import pythoncom
from dotenv import load_dotenv
from pathlib import Path
import json
import re

# Import config from parent directory
sys.path.insert(0, str(Path(__file__).parent.parent.parent / "config"))
import config

load_dotenv(os.path.join(config.BASE_DIR, "_internals", ".env"))
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

# === CONFIG (from config.py) ===
SAVE_DIR = config.DOCUMENT_COMBINER_SETTINGS["save_dir"]
CLAIMS_PATH = config.DOCUMENT_COMBINER_SETTINGS["claims_folder_path"]
PERMITS_PATH = config.DOCUMENT_COMBINER_SETTINGS["permits_folder_path"]
MAILBOX_NAME = config.DOCUMENT_COMBINER_SETTINGS["mailbox_name"]
CLAIMS_DEST_FOLDER = config.DOCUMENT_COMBINER_SETTINGS["claims_dest_folder"]
PERMITS_DEST_FOLDER = config.DOCUMENT_COMBINER_SETTINGS["permits_dest_folder"]


def extract_claim_number_from_email(full_text):
    """Extract claim number using OpenAI LLM."""
    prompt = f"""
    You are an AI assistant extracting claim numbers from Outlook email content.

    Extract only the single, most recent claim number. A valid claim number must satisfy exactly one of the following:

    1. **Starts with "904"** and is exactly 9 digits long. **Do not include any spaces or other characters—return exactly (e.g.) "904123456"**
    2. Otherwise, it begins with "TD" followed immediately by digits (any length and not starting with 904). In that case include the full "TD…" string (e.g. "TD2331627").

    ❗❗ **Example claim number edge case – If the claim appears as "TD: 904835662" or "TD 904835662," strip the "TD" and return only "904835662" (with no spaces). 904 Claim number should always be 9 digits.**

    🚫 Do NOT extract addresses, structure numbers, etc.
    🚫 Only return one claim number, no spaces.

    Format:
    {{
      "Claim Number": ""
    }}

    Email:
    {full_text.strip()}
    """
    response = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": prompt}]
    )
    try:
        raw = json.loads(response.choices[0].message.content).get("Claim Number", "").strip()
    except:
        raw = ""

    # Sanitization Step
    normalized = raw.strip()

    # If it's "TD: 904..." or "TD 904..." → strip "TD" and any punctuation/spaces
    m = re.match(r"^TD[:\-]?\s*(904\d{6}[A-E]?)$", normalized, re.IGNORECASE)
    if m:
        normalized = m.group(1)

    # If it starts with "904" + 6 digits + optional letter, strip any extra chars
    m2 = re.match(r"^(904\d{6}[A-E]?)", normalized, re.IGNORECASE)
    if m2:
        normalized = m2.group(1)

    normalized = normalized.upper()
    return normalized


def save_attachments_to_temp(msg, temp_dir=None, seen_entryids=None):
    """
    Saves all attachments of `msg` into a temporary folder.
    Handles nested .msg files recursively.
    Returns flat list of file paths.
    """
    if temp_dir is None:
        temp_dir = tempfile.mkdtemp()
    if seen_entryids is None:
        seen_entryids = set()

    saved_paths = []
    eid = msg.EntryID

    # Convert message body to PDF (unique per EntryID)
    if eid not in seen_entryids:
        seen_entryids.add(eid)

        body_pdf_name = f"{eid}.email_body.pdf"
        body_pdf = os.path.join(temp_dir, body_pdf_name)
        try:
            inspector = msg.GetInspector
            word_editor = inspector.WordEditor
            word_editor.ExportAsFixedFormat(
                OutputFileName=body_pdf,
                ExportFormat=17,  # wdExportFormatPDF
                OpenAfterExport=False,
                OptimizeFor=0,
                Range=0
            )
            saved_paths.append(body_pdf)
        except Exception as e:
            print(f"Failed to export email body {eid}: {e}")

    # Save attachments and handle nested .msg files
    for attachment in msg.Attachments:
        filename = attachment.FileName
        filepath = os.path.join(temp_dir, filename)
        try:
            attachment.SaveAsFile(filepath)
        except Exception as e:
            print(f"Failed to save attachment {filename}: {e}")
            continue

        # Recurse into nested .msg files
        if filename.lower().endswith(".msg"):
            try:
                pythoncom.CoInitialize()
                outlook = win32com.client.Dispatch("Outlook.Application")
                nested_mail = outlook.CreateItemFromTemplate(filepath)

                nested_paths = save_attachments_to_temp(nested_mail, temp_dir, seen_entryids)
                saved_paths.extend(nested_paths)
            except Exception as e:
                print(f"Failed to unpack nested .msg {filename}: {e}")
        else:
            # Regular attachment (PDF, image, docx, etc.)
            saved_paths.append(filepath)

    return saved_paths


def convert_to_pdf(filepath):
    """Convert various file types to PDF."""
    ext = os.path.splitext(filepath)[1].lower()
    fname = os.path.basename(filepath).lower()

    # Skip sensitive files
    if ("tasking" in fname or "pricing" in fname
        or fname.strip() == "zero dollar sasco tasking sheet - template.xlsx"):
        print(f"Skipping sensitive file: {fname}")
        return None

    if ext == '.pdf':
        return filepath

    output_path = filepath + ".converted.pdf"
    try:
        if ext in ['.jpg', '.jpeg', '.png']:
            image = Image.open(filepath)
            rgb = image.convert('RGB')
            rgb.save(output_path)
            return output_path

        elif ext == '.docx':
            temp_pdf_path = filepath.replace('.docx', '.converted.pdf')
            convert_docx(filepath, temp_pdf_path)
            return temp_pdf_path

        elif ext == '.xlsx':
            excel = win32com.client.Dispatch("Excel.Application")
            wb = excel.Workbooks.Open(filepath)
            wb.ExportAsFixedFormat(0, output_path)
            wb.Close(True)
            return output_path
    except Exception as e:
        print(f"Conversion failed for {filepath}: {e}")
        return None


def merge_pdfs(pdf_list, output_path, append=False):
    """Merge multiple PDFs into single output file. Skips corrupted PDFs."""
    writer = PdfWriter()
    skipped_files = []

    # Add existing file if in append mode
    if append and os.path.exists(output_path):
        try:
            reader = PdfReader(output_path)
            for page in reader.pages:
                writer.add_page(page)
        except Exception as e:
            print(f"Warning: Skipping corrupted existing file {os.path.basename(output_path)}: {e}")
            skipped_files.append(f"{os.path.basename(output_path)} (existing)")

    # Add each PDF page, skipping ones that fail
    for pdf in pdf_list:
        if pdf:
            try:
                reader = PdfReader(pdf)
                for page in reader.pages:
                    writer.add_page(page)
            except Exception as e:
                print(f"Warning: Skipping corrupted or invalid PDF {os.path.basename(pdf)}: {e}")
                skipped_files.append(os.path.basename(pdf))

    # Only write if we have at least one valid page
    if len(writer.pages) > 0:
        try:
            with open(output_path, 'wb') as f:
                writer.write(f)
            if skipped_files:
                print(f"Note: {len(skipped_files)} file(s) were skipped due to format issues")
        except Exception as e:
            print(f"Error writing merged PDF: {e}")
            raise
    else:
        print("Error: No valid PDFs to merge!")


def process_folder(folder_path, category_label, is_permit, dest_folder_path):
    """Process a specific Outlook folder (Claims or Permits)."""
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    folder = outlook.Folders.Item(MAILBOX_NAME)
    for level in folder_path:
        folder = folder.Folders[level]
    messages = folder.Items
    messages.Sort("[ReceivedTime]", True)

    print(f"Processing {folder_path[-1]} folder...")

    if len(messages) == 0:
        print(f"No emails found in folder: {' > '.join(folder_path)}")
        return

    for message in list(messages):
        try:
            print(f"Processing: {message.Subject} | {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

            # Skip already processed permits
            if is_permit and message.Categories and "Permit Combined" in message.Categories:
                print(f"Skipping: already marked as 'Permit Combined'")
                continue

            full_text = (
                f"From: {message.SenderName} <{message.SenderEmailAddress}>\n"
                f"To: {message.To}\n"
                f"CC: {message.CC}\n"
                f"Sent: {message.ReceivedTime}\n"
                f"Subject: {message.Subject}\n\n"
                f"{message.Body.strip()}"
            )
            claim_number = extract_claim_number_from_email(full_text)
            if not claim_number:
                print("No claim number found, skipping.")
                print("-" * 60)
                continue

            output_file = os.path.join(SAVE_DIR, f"SiteDocs.{claim_number}.pdf")
            file_exists = os.path.exists(output_file)

            # Skip if claims file already exists
            if not is_permit and file_exists:
                print(f"Skipped: {output_file} already exists.")
                print("-" * 60)
                continue

            # Warn if permits file doesn't exist yet
            if is_permit and not file_exists:
                print(f"No existing SiteDocs file found for permit: {output_file}")

            # Get all attachments (including nested email PDFs)
            attachments = save_attachments_to_temp(message)
            pdfs = [convert_to_pdf(f) for f in attachments]
            pdfs = [p for p in pdfs if p is not None]

            if not pdfs and not is_permit:
                print("No valid attachments to merge.")
                print("-" * 60)
                continue

            # Merge PDFs
            if pdfs:
                joined_names = "\n   - " + "\n   - ".join([os.path.basename(p) for p in pdfs])
                label = "appended" if is_permit and file_exists else "combined"
                print(f"Files {label}:{joined_names}")
                merge_pdfs(pdfs, output_file, append=is_permit)
                print(f"Merged and saved to: {output_file}")

            # Tag email
            message.Categories = category_label
            message.Save()

            # Move to final destination folder
            dest_folder = outlook.Folders.Item(MAILBOX_NAME)
            for level in dest_folder_path:
                dest_folder = dest_folder.Folders[level]
            message.Move(dest_folder)
            print(f"Moved email to: {' > '.join(dest_folder_path)}")
            print("-" * 60)

        except Exception as e:
            print(f"Error: {e}")


def combine_attachments():
    """Main function: process both Claims and Permits folders."""
    pythoncom.CoInitialize()
    print("DocumentCombiner - Email Attachment Merger (Config-Driven)")
    process_folder(CLAIMS_PATH, "Combined", is_permit=False, dest_folder_path=CLAIMS_DEST_FOLDER)
    process_folder(PERMITS_PATH, "Permit Combined", is_permit=True, dest_folder_path=PERMITS_DEST_FOLDER)


if __name__ == "__main__":
    combine_attachments()
