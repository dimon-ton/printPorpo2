from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from io import BytesIO


FONT_NAME = "DSN-LaiThai"
FONT_SIZE = 20 

font_path = "DSN-LaiThai.ttf"  # Path to your Thai font file
pdfmetrics.registerFont(TTFont(FONT_NAME, font_path))

existing_pdf_path = "examples.pdf"  # Path to your existing PDF
output_pdf_path = "output_with_thai_text.pdf"

existing_pdf = PdfReader(existing_pdf_path)
output_pdf = PdfWriter()

# read pdf file in first page
page = existing_pdf.pages[0]

page.rotate(90)

page_width = float(page.mediabox.upper_right[0])
page_height = float(page.mediabox.upper_right[1])


print(page_width)
print(page_height)

# create canvas for editing PDF file 
packet = BytesIO()  # In-memory buffer to hold the overlay PDF
can = canvas.Canvas(packet, pagesize=(page_width, page_height))

# Set the font to the registered Thai font
can.setFont(FONT_NAME, FONT_SIZE)  # Font size 16


# Add Thai text to the PDF
name = "นายณรงค์ฤทธิ์ สมรูป"
text_width = can.stringWidth(name, FONT_NAME, FONT_SIZE)
page_center = (page_height - text_width) / 2

can.saveState()

can.translate(page_width, 0)
can.rotate(90)

can.drawString(page_center, 265, name)  # Position (x=50, y=750)

birthNum = "๒๔"
birthMonth = "กุมภาพันธ์"
birthYear = "๒๕๕๒"

can.drawString(196, 232, birthNum)
can.drawString(269, 232, birthMonth)
can.drawString(405, 232, birthYear)

# insert school name
school_name = "โรงเรียนบ้านโพนแท่น"
can.drawString(150, 178, school_name)

# insert province name
province_name = "ร้อยเอ็ด"
can.drawString(155, 153, province_name)


# insert office
office_name = "สพป. ร้อยเอ็ด เขต ๒"
can.drawString(310, 153, office_name)

# insert graduated date
graduated_date = "๓๑"
graduated_month = "มีนาคม"
graduated_year = "๒๕๖๘"

can.drawString(195, 126, graduated_date)
can.drawString(285, 126, graduated_month)
can.drawString(400, 126, graduated_year)

# insert signature end
dotted_line = 90*"."
dotted_width = can.stringWidth(dotted_line, FONT_NAME, FONT_SIZE)
dotted_page_center = (page_height - dotted_width) / 2
can.drawString(dotted_page_center, 60, dotted_line)


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