from docx import Document
import fitz  # PyMuPDF
from typing import Dict

def parse_word_form(docx_path: str) -> Dict[str, str]:
    """Parse the first table in a Word document and extract required fields."""
    doc = Document(docx_path)
    table = doc.tables[0]
    data = {}
    for row in table.rows:
        if len(row.cells) == 2:
            key = row.cells[0].text.strip()
            value = row.cells[1].text.strip()
            data[key] = value
    return {
        "Company Name": data.get("Company Name", ""),
        "Address": data.get("Address", ""),
        "ISO Standard": data.get("ISO Standard Required", "").splitlines()[-1],
        "Scope": data.get("Scope", "")
    }

def get_text_height(text: str, fontsize: float, fontname: str, max_width: float) -> float:
    """Estimate the height of a text block when wrapped to fit max_width."""
    font = fitz.Font(fontname=fontname)
    words = text.split()
    lines = []
    current_line = ""
    for word in words:
        test_line = current_line + (" " if current_line else "") + word
        if font.text_length(test_line, fontsize) <= max_width:
            current_line = test_line
        else:
            lines.append(current_line)
            current_line = word
    if current_line:
        lines.append(current_line)
    return len(lines) * fontsize * 1.2  # Approximate line height with spacing

def insert_centered_textbox(
    page: fitz.Page,
    rect: fitz.Rect,
    text: str,
    fontname: str,
    fontsize: float,
    color: tuple
) -> None:
    """Insert text centered both vertically and horizontally in the given rectangle."""
    print(f"  🔧 insert_centered_textbox called with: text='{text}', rect={rect}, fontsize={fontsize}")
    
    # Calculate text height and center it vertically
    text_height = get_text_height(text, fontsize, fontname, rect.width)
    start_y = rect.y0 + (rect.height - text_height) / 2
    box = fitz.Rect(rect.x0, start_y, rect.x1, start_y + text_height)
    
    print(f"  📐 Calculated: text_height={text_height:.1f}, start_y={start_y:.1f}, box={box}")
    
    page.insert_textbox(
        box,
        text,
        fontsize=fontsize,
        fontname=fontname,
        color=color,
        align=1  # Centered
    )
    print(f"  ✅ insert_textbox completed for '{text}'")

def generate_certificate(base_pdf_path: str, output_pdf_path: str, values: Dict[str, str]) -> None:
    """Generate a certificate PDF by overlaying extracted values onto a template."""
    doc = fitz.open(base_pdf_path)
    page = doc[0]

    # Test basic text rendering
    page.insert_text((100, 100), "TEST TEXT - RED", fontsize=20, fontname="Times-Bold", color=(1, 0, 0))
    print("🔍 Added test text at position (100, 100)")
    
    # Test insert_centered_textbox function directly
    test_rect = fitz.Rect(200, 200, 400, 250)
    insert_centered_textbox(page, test_rect, "CENTERED TEST", "Times-Bold", 20, (1, 0, 0))
    print("🔍 Added centered test text in rectangle (200, 200, 400, 250)")
    
    # Test with simple insert_textbox (without our custom function)
    page.insert_textbox(
        fitz.Rect(300, 300, 500, 350),
        "SIMPLE TEST",
        fontsize=20,
        fontname="Times-Bold",
        color=(1, 0, 0),
        align=1
    )
    print("🔍 Added simple test text in rectangle (300, 300, 500, 350)")

    # --- Configuration ---
    color = (1, 0, 0)  # Red - to test if text is being written
    fontname = "Times-Bold"  # Use bold font
    coords = {
        "Company Name": fitz.Rect(179.2, 227.6, 476.0, 266.4),
        "Address": fitz.Rect(169.4, 264.8, 485.9, 310), #Keep upwards
        "ISO Standard": fitz.Rect(194.9, 334, 460.3, 370),  # Moved up by 30 points
        "Scope": fitz.Rect(87.9, 386, 578, 475),
    }
    font_starts = {
        "Company Name": 40,
        "Scope": 35,
        "ISO Standard": 50,
        "Address": 30
    }
    # --- End Configuration ---

    for field, text in values.items():
        rect = coords[field]
        start_size = font_starts.get(field, 30)
        font_size = start_size
        
        print(f"🔍 Processing '{field}': text='{text}', rect={rect}, start_size={start_size}")
        
        # Draw a colored rectangle around each field for debugging
        if field == "Company Name":
            page.draw_rect(rect, color=(1, 0, 0), width=2)  # Red for Company Name
        elif field == "Address":
            page.draw_rect(rect, color=(0, 1, 0), width=2)  # Green for Address
        elif field == "ISO Standard":
            page.draw_rect(rect, color=(0, 0, 1), width=2)  # Blue for ISO Standard
        elif field == "Scope":
            page.draw_rect(rect, color=(1, 1, 0), width=2)  # Yellow for Scope
        
        # Reduce font size if it doesn't fit
        while font_size >= 10:
            text_height = get_text_height(text, font_size, fontname, rect.width)
            limit = rect.height if field != "Company Name" else rect.height * 2
            print(f"  📏 Font size {font_size}: text_height={text_height:.1f}, limit={limit:.1f}")
            if text_height <= limit:
                break
            font_size -= 1
        
        print(f"🖋 Writing '{field}' at font size {font_size}")
        # Use simple insert_textbox instead of our custom function
        page.insert_textbox(
            rect,
            text,
            fontsize=font_size,
            fontname=fontname,
            color=color,
            align=1  # Centered
        )

    doc.save(output_pdf_path)
    print(f"✅ Certificate saved at: {output_pdf_path}")

if __name__ == "__main__":
    word_path = "generate_certificate/form.docx"
    template_path = "generate_certificate/Draft.pdf"
    output_path = "Generated_Certificate.pdf"

    values = parse_word_form(word_path)
    generate_certificate(template_path, output_path, values)
