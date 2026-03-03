# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Python script that generates Thai language student certificates by overlaying student data from an Excel spreadsheet onto a PDF template. The script is designed for creating Thai-language certificates with proper Thai numeral support and font rendering.

## Development Commands

```bash
# Activate the virtual environment (Windows)
env\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run the certificate generation script
python main.py
```

## Key Dependencies

- **PyPDF2**: PDF manipulation (reading templates, merging pages)
- **reportlab**: PDF generation with Thai font support
- **openpyxl**: Reading student data from Excel files
- **pillow**: Image processing (used by reportlab)

## Architecture

### Data Flow
1. Student data is read from `name_list.xlsx` (Sheet1, starting from row 2)
2. A PDF template `examples.pdf` is used as the base certificate
3. Thai text overlay is created using reportlab with the `DSN-LaiThai.ttf` font
4. Each student generates one page in the output PDF `output_with_thai_text.pdf`

### Main Components

**`main.py`** - Single-file script containing:
- `convert_to_thai_number()`: Converts Arabic numerals to Thai numerals (๐๑๒๓...)
- `get_excel_data()`: Reads student records from Excel file
- Certificate generation loop: Creates overlay PDF for each student and merges with template

### Excel Data Format (name_list.xlsx)

Expected columns (zero-indexed):
- `[0]`: Running number (converted to Thai numerals)
- `[1]`: Student title/prefix
- `[2]`: First name
- `[3]`: Last name
- `[4]`: Birth date number (converted to Thai numerals)
- `[5]`: Birth month (Thai text)
- `[6]`: Birth year (converted to Thai numerals)

### Configuration Constants

- `FONT_NAME = "DSN-LaiThai"` - Thai font for certificate text
- `FONT_SIZE = 20` - Font size for certificate text
- `remove_background = True` - When True, only the overlay is output; when False, merges with template PDF

### Coordinate System

The PDF is rotated 90 degrees and uses a transformed coordinate system:
- `can.translate(page_width, 0)` + `can.rotate(90)` transforms coordinates
- Text positioning is relative to the rotated page dimensions
- Center-aligned text uses `can.stringWidth()` to calculate center position

### Hard-Coded Values

The following values are currently static (may need updates for different use cases):
- School name: "โรงเรียนบ้านโพนแท่น"
- Province: "ร้อยเอ็ด"
- Office: "สพป. ร้อยเอ็ด เขต ๒"
- Graduation date: ๓๑ มีนาคม ๒๕๖๘
- Head teacher: "(นางสาวอำพร วรวงษ์)"
- Position: "รักษาการในตำแหน่งผู้อำนวยการโรงเรียนบ้านโพนแท่น"

## File Encoding

The `requirements.txt` file is UTF-16 LE encoded. When editing, preserve this encoding or convert to UTF-8.
