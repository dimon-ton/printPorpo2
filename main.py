from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from io import BytesIO
from reportlab.lib.colors import red, green, black, Color
import json
import os

from openpyxl import load_workbook


def hex_to_color(hex_color):
    """Convert hex color string to reportlab Color object."""
    hex_color = hex_color.lstrip('#')
    if len(hex_color) == 6:
        r = int(hex_color[0:2], 16) / 255.0
        g = int(hex_color[2:4], 16) / 255.0
        b = int(hex_color[4:6], 16) / 255.0
        return Color(r, g, b)
    return black  # Default to black if invalid


def load_config():
    """Load configuration from certificate_config.json if it exists."""
    config_file = "certificate_config.json"
    defaults = {
        "school_name": "โรงเรียนบ้านโพนแท่น",
        "province_name": "ร้อยเอ็ด",
        "office_name": "สพป. ร้อยเอ็ด เขต 2",
        "graduated_date": "31",
        "graduated_month": "มีนาคม",
        "graduated_year": "2568",
        "head_teacher_name": "(นางสาวอำพร วรวงษ์)",
        "position_name": "รักษาการในตำแหน่งผู้อำนวยการโรงเรียนบ้านโพนแท่น",
        "excel_file": "name_list.xlsx",
        "template_pdf": "examples.pdf",
        "output_pdf": "output_with_thai_text.pdf",
        "font_file": "DSN-LaiThai.ttf",
        "font_size": 20,
        "remove_background": True,
        # Color options (hex format)
        "color_running_number": "#00AA00",  # Green
        "color_student_name": "#000000",    # Black
        "color_birth_date": "#000000",      # Black
        "color_school_name": "#000000",     # Black
        "color_province": "#000000",        # Black
        "color_office": "#000000",          # Black
        "color_graduation_date": "#000000", # Black
        "color_signer": "#000000",          # Black
        "color_position": "#000000"         # Black
    }

    if os.path.exists(config_file):
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
                # Convert string font_size to int and remove_background to bool
                config["font_size"] = int(config.get("font_size", 20))
                config["remove_background"] = config.get("remove_background", "True") == "True"
                return {**defaults, **config}
        except Exception as e:
            print(f"Warning: Could not load config file: {e}")
            print("Using default values.")

    return defaults


# Load configuration at module level
CONFIG = load_config()


def convert_to_thai_number(number_str):
    """
    Converts a standard number string to a Thai numeral string.

    Parameters:
        number_str (str): A string containing standard Arabic numerals (e.g., "123456").

    Returns:
        str: A string containing Thai numerals (e.g., "๑๒๓๔๕๖").
    """
    if number_str is None:
        return ""
    number_str = str(number_str)
    # Mapping of Arabic numerals to Thai numerals
    arabic_to_thai = {
        '0': '๐',
        '1': '๑',
        '2': '๒',
        '3': '๓',
        '4': '๔',
        '5': '๕',
        '6': '๖',
        '7': '๗',
        '8': '๘',
        '9': '๙'
    }

    # Convert each character in the input string using the mapping
    thai_number_str = ''.join(arabic_to_thai.get(char, char) for char in number_str)

    return thai_number_str



def get_excel_data(file_path):

    workbook = load_workbook(file_path, data_only=True)
    sheet = workbook["Sheet1"]

    data = []

    for row in sheet.iter_rows(min_row=2, values_only=True):
        # Skip rows where all cells are empty
        if all(cell is None for cell in row):
            continue
        data.append(row)

    return data


# Use configuration values
file_path = CONFIG["excel_file"]
list_data = get_excel_data(file_path)

print(f"Loaded {len(list_data)} records from {file_path}")

remove_background = CONFIG["remove_background"]

FONT_NAME = os.path.splitext(CONFIG["font_file"])[0]  # Use filename without extension as font name
FONT_SIZE = CONFIG["font_size"]

font_path = CONFIG["font_file"]
pdfmetrics.registerFont(TTFont(FONT_NAME, font_path))

existing_pdf_path = CONFIG["template_pdf"]
output_pdf_path = CONFIG["output_pdf"]

existing_pdf = PdfReader(existing_pdf_path)
output_pdf = PdfWriter()

# read pdf file in first page
page = existing_pdf.pages[0]

page.rotate(90)

page_width = float(page.mediabox.upper_right[0])
page_height = float(page.mediabox.upper_right[1])


print(page_width)
print(page_height)


for data in list_data:
    # create canvas for editing PDF file 
    packet = BytesIO()  # In-memory buffer to hold the overlay PDF
    can = canvas.Canvas(packet, pagesize=(page_width, page_height))

    # Set the font to the registered Thai font
    can.setFont(FONT_NAME, FONT_SIZE)  # Font size 16

    can.translate(page_width, 0)
    can.rotate(90)

    can.saveState()

    # insert running number
    can.setFillColor(hex_to_color(CONFIG["color_running_number"]))
    running_number = convert_to_thai_number(data[0])
    can.drawString(500, 361, running_number)

    # insert name of student
    can.setFillColor(hex_to_color(CONFIG["color_student_name"]))
    name = f"{data[1] or ''}{data[2] or ''} {data[3] or ''}"
    text_width = can.stringWidth(name, FONT_NAME, FONT_SIZE)
    page_center = (page_height - text_width) / 2
    can.drawString(page_center, 260, name)

    # insert birth date
    can.setFillColor(hex_to_color(CONFIG["color_birth_date"]))
    birthNum = convert_to_thai_number(data[4])
    birthMonth = data[5] or ''
    birthYear = convert_to_thai_number(data[6])

    can.drawString(196, 230, birthNum)
    can.drawString(269, 230, birthMonth)
    can.drawString(405, 230, birthYear)

    # insert school name
    can.setFillColor(hex_to_color(CONFIG["color_school_name"]))
    school_name = CONFIG["school_name"]
    can.drawString(150, 178, school_name)

    # insert province name
    can.setFillColor(hex_to_color(CONFIG["color_province"]))
    province_name = CONFIG["province_name"]
    can.drawString(155, 153, province_name)

    # insert office
    can.setFillColor(hex_to_color(CONFIG["color_office"]))
    office_name = convert_to_thai_number(CONFIG["office_name"])
    can.drawString(310, 153, office_name)

    # insert graduated date (convert Arabic numerals to Thai)
    can.setFillColor(hex_to_color(CONFIG["color_graduation_date"]))
    graduated_date = convert_to_thai_number(CONFIG["graduated_date"])
    graduated_month = CONFIG["graduated_month"]
    graduated_year = convert_to_thai_number(CONFIG["graduated_year"])

    can.drawString(195, 126, graduated_date)
    can.drawString(285, 126, graduated_month)
    can.drawString(400, 126, graduated_year)

    # insert signature end
    can.setFillColor(black)
    dotted_line = 90*"."
    dotted_width = can.stringWidth(dotted_line, FONT_NAME, FONT_SIZE)
    dotted_page_center = (page_height - dotted_width) / 2
    can.drawString(dotted_page_center, 60, dotted_line)

    # insert name head teacher
    can.setFillColor(hex_to_color(CONFIG["color_signer"]))
    head_teacher_name = CONFIG["head_teacher_name"]
    head_teacher_width = can.stringWidth(head_teacher_name, FONT_NAME, FONT_SIZE)
    head_page_center = (page_height - head_teacher_width) / 2
    can.drawString(head_page_center, 37, head_teacher_name)

    # insert position
    can.setFillColor(hex_to_color(CONFIG["color_position"]))
    position_name = CONFIG["position_name"]
    position_name_width = can.stringWidth(position_name, FONT_NAME, FONT_SIZE)
    position_page_center = (page_height - position_name_width) / 2
    can.drawString(position_page_center, 13, position_name)


    # insert position
    position_name = CONFIG["position_name"]
    position_name_width = can.stringWidth(position_name, FONT_NAME, FONT_SIZE)
    position_page_center = (page_height - position_name_width) / 2
    can.drawString(position_page_center, 13, position_name)


    can.restoreState()

    # Save the overlay PDF to the in-memory buffer
    can.save()

    # Move the buffer's pointer to the beginning so it can be read
    packet.seek(0)

    if not remove_background:

        existing_pdf = PdfReader(existing_pdf_path)
        
        # read pdf file in first page
        page = existing_pdf.pages[0]

        page.rotate(90)

        # Merge the overlay PDF (with Thai text) onto the existing PDF page
        overlay_pdf = PdfReader(packet)
        page.merge_page(overlay_pdf.pages[0])        

        # Add the modified page to the output PDF
        output_pdf.add_page(page)
    else:
       overlay_pdf = PdfReader(packet)
       page = overlay_pdf.pages[0]
       page.rotate(90)
       output_pdf.add_page(page) 


# Write the final PDF to disk
with open(output_pdf_path, "wb") as output_file:
    output_pdf.write(output_file)

print(f"Thai text has been added to the PDF. Output saved as: {output_pdf_path}")