from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from io import BytesIO

# Step 1: Register the Thai font
font_path = "DSN-LaiThai.ttf"  # Path to your Thai font file
pdfmetrics.registerFont(TTFont("DSN-LaiThai", font_path))

# Step 3: Read the existing PDF and overlay the Thai text
existing_pdf_path = "examples.pdf"  # Path to your existing PDF
output_pdf_path = "output_with_thai_text.pdf"

# Read the existing PDF
existing_pdf = PdfReader(existing_pdf_path)
output_pdf = PdfWriter()

# Add the first page of the existing PDF to the output PDF
page = existing_pdf.pages[0]

page.rotate(90)

page_width = float(page.mediabox.upper_right[0])
page_height = float(page.mediabox.upper_right[1])

# page.mediabox.upper_right = (page_height, page_width)
# output_pdf.add_page(page)


# new_page_width = float(page.mediabox.upper_right[0])
# new_page_height = float(page.mediabox.upper_right[1])

print(page_width)
print(page_height)

# Step 2: Create a temporary PDF with Thai text using ReportLab
packet = BytesIO()  # In-memory buffer to hold the overlay PDF
can = canvas.Canvas(packet, pagesize=(page_height, page_width))

# Set the font to the registered Thai font
can.setFont("DSN-LaiThai", 16)  # Font size 16


# Add Thai text to the PDF
thai_text = "นายธนวัฒน์ PDF"


text_width = can.stringWidth(thai_text, "DSN-LaiThai", 16)
page_center = (page_width - text_width) / 2

can.saveState()

can.translate(page_width, 0)
can.rotate(90)

can.drawString(0, 0, thai_text)  # Position (x=50, y=750)

can.restoreState()

# Save the overlay PDF to the in-memory buffer
can.save()

# Move the buffer's pointer to the beginning so it can be read
packet.seek(0)


# Merge the overlay PDF (with Thai text) onto the existing PDF page
overlay_pdf = PdfReader(packet)
page.merge_page(overlay_pdf.pages[0])
# Add the modified page to the output PDF
output_pdf.add_page(page)

# Write the final PDF to disk
with open(output_pdf_path, "wb") as output_file:
    output_pdf.write(output_file)

print(f"Thai text has been added to the PDF. Output saved as: {output_pdf_path}")