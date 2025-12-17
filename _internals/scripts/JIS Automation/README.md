# JIS Automation Module

This module handles Job Instruction Sheet (JIS) automation.

## Scripts

- **JIS.py** - Main JIS automation script
- **MasterJIS.py** - Master JIS file generation

## Data Folders

- **Input/** - Place input files here
- **Output/** - Generated files will be saved here
- **ConversionChart/** - Reference conversion files
- **Templates/** - JIS template files

## Setup

1. Add your JIS template file to the `Templates/` folder
2. Place input data in the `Input/` folder
3. Run the script via the Bismillah launcher: `[7] 📄 Run JIS Automation`

For detailed instructions, see the script docstrings or the main README.md.

## Configuration

Update `config/config.py` to customize paths:

```python
JIS_SETTINGS = {
    "input_folder": os.path.join(DATA_PATHS["input"], "jis_input"),
    "output_folder": os.path.join(DATA_PATHS["output"], "jis_output"),
    "template_file": os.path.join(DATA_PATHS["templates"], "JIS Template.xlsx"),
}
```
