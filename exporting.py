import openpyxl
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE

# Load names from Excel file
def load_names_from_excel(file_path):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    names = []
    for row in sheet.iter_rows(min_row=2, max_col=1, values_only=True):
        if row[0]:
            names.append(row[0])
    return names

# Create lower-third slide for each name
def create_lowerthird_ppt(names, output_file='lowerthirds.pptx'):
    prs = Presentation()

    for name in names:
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout

        # Add textbox for lower third
        left = Inches(1)
        top = Inches(5.5)  # Near the bottom
        width = Inches(8)
        height = Inches(1)

        # Set slide size to 16:9 (13.33 x 7.5 inches)
        prs.slide_width = Inches(13.33)
        prs.slide_height = Inches(7.5)

        # Set slide background color to red
        background = slide.background
        fill = background.fill
        fill.solid()
        # Set slide background color to hex #0070C0
        fill.fore_color.rgb = RGBColor.from_string('0070C0')
        textbox = slide.shapes.add_textbox(left, top, width, height)
        text_frame = textbox.text_frame
        text_frame.text = name

        # Format text
        p = text_frame.paragraphs[0]
        from pptx.enum.text import PP_ALIGN
        p.alignment = PP_ALIGN.CENTER  # Center horizontally

        # Vertically center the text frame content
        from pptx.enum.text import MSO_ANCHOR
        text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

        run = p.runs[0]
        run.font.size = Pt(28)
        run.font.bold = True
        run.font.color.rgb = RGBColor.from_string('FFFFFF')

        # Make textbox width same as slide width
        textbox.width = prs.slide_width
        textbox.left = 0

        # Set background color of the textbox
        fill = textbox.fill
        fill.solid()  # Start with solid fill to reset
        fill.gradient()  # Switch to gradient fill
        fill.gradient_angle = 0  # Horizontal gradient (left to right)
        # Set gradient stops
        stop_main = fill.gradient_stops[0]
        stop_main.position = 0.5  # Center focus
        stop_main.color.rgb = RGBColor.from_string('6F0F11')
        stop_side = fill.gradient_stops[1]
        stop_side.position = 1.0
        # Use a similar color for the side (slightly lighter)
        stop_side.color.rgb = RGBColor.from_string('8F1F21')

        # Add rounded border
        # Add only a bottom border (simulate by adding a thin rectangle shape)
        border_height = Pt(6)  # Thickness of the bottom line
        border = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            textbox.left,
            textbox.top + textbox.height - border_height,
            textbox.width,
            border_height
        )
        border.fill.solid()
        border.fill.fore_color.rgb = RGBColor(255, 215, 0)  # Gold color
        border.line.fill.background()  # No outline
        border.shadow.inherit = False

    prs.save(output_file)
    print(f"Presentation saved as {output_file}")

# Main
excel_file = 'names.xlsx'  # <-- Change if needed
names = load_names_from_excel(excel_file)
create_lowerthird_ppt(names)
