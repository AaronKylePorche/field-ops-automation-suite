# Document Processing Module

This module handles claims, permits, packing slips, and document combining.

## Scripts

- **DocumentCombiner.py** - Combines claims and permit documents
- **ClaimPro1.py** - Adds new claims to Excel tracking file
- **packing_slip_scanner.py** - Scans and processes packing slips folder
- **pspllm5.py** - Legacy packing slip processor

## Data Folders

- **Input/** - Source documents (PDFs, images, etc.)
- **Output/** - Processed/combined documents
- **Packing Slips/** - Packing slip files to be scanned

## Workflow

### Adding Claims
1. Run: `[3] 📥 Add New Claims to Excel`
2. Script updates your Excel tracking file

### Combining Documents
1. Run: `[2] 📦 Combine Claims & Permits`
2. Documents in Input folder will be combined

### Processing Packing Slips
1. Run: `[4] 🔍 SCAN Slips Folder`
2. Packing slips will be processed and organized

## Input File Types

- PDF files (.pdf)
- Image files (.jpg, .png, .tiff)
- Excel files (.xlsx, .xls)

## Configuration

Update `config/config.py` if your data is stored elsewhere:

```python
DATA_PATHS = {
    "input": os.path.join(BASE_DIR, "data", "input"),
    "output": os.path.join(BASE_DIR, "data", "output"),
    # ... other paths
}
```

## Notes

- Each script has its own docstring with detailed usage
- Check the script comments for customization options
- Some scripts may require additional configuration in their code

---

For more details, see the individual script files or the main README.md.
