from docx import Document
import fitz  # PyMuPDF

def parse_word_form(docx_path):
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

def generate_certificate(base_pdf_path, output_pdf_path, values):
    doc = fitz.open(base_pdf_path)
    page = doc[0]

    fontname = "Times-Roman"
    font_size = 30
    color = (1, 0, 0)  # Red for text (for debugging)

    coords = {
        "Company Name": fitz.Rect(179.2, 233.6, 476.0, 266.4),
        "Address": fitz.Rect(169.4, 264.8, 485.9, 294.8),  # Increased height by 30 units
        "ISO Standard": fitz.Rect(194.9, 314.1, 460.3, 357.2),
        "Scope": fitz.Rect(87.9, 362.3, 488.3, 402.3),
    }

    print("Extracted values:", values)
    print("Field coordinates:")
    for field, rect in coords.items():
        print(f"{field}: {rect}")

    for field, text in values.items():
        rect = coords[field]
        # Draw a green rectangle for debugging
        page.draw_rect(rect, color=(0, 1, 0), width=1)
        if field == "Scope":
            if len(text.strip()) < 50:
                # Short scope — use single-line center
                text_width = fitz.get_text_length(text, fontname=fontname, fontsize=font_size)
                center_x = (rect.x0 + rect.x1) / 2 - text_width / 2
                y = rect.y0 + 5  # a little vertical offset
                print(f"Drawing short Scope: '{text}' at font size {font_size}")
                page.insert_text(
                    (center_x, y),
                    text,
                    fontsize=font_size,
                    fontname=fontname,
                    color=color,
                    render_mode=0  # normal text
                )
            else:
                # Long scope — wrap and auto-fit
                best_size = font_size
                size = font_size
                while size >= 10:
                    # Use a dummy page to check fitting
                    dummy_doc = fitz.open()
                    dummy_page = dummy_doc.new_page(width=page.rect.width, height=page.rect.height)
                    result = dummy_page.insert_textbox(
                        rect,
                        text,
                        fontsize=size,
                        fontname=fontname,
                        color=color,
                        align=0  # left aligned for better wrapping
                    )
                    dummy_doc.close()
                    print(f"Trying Scope at size {size}, result: {result}")
                    if result == 0:
                        best_size = size
                        break
                    best_size = size
                    size -= 1
                print(f"Drawing long Scope: '{text}' at font size {best_size}")
                page.insert_textbox(
                    rect,
                    text,
                    fontsize=best_size,
                    fontname=fontname,
                    color=color,
                    align=0
                )
        else:
            # Auto-fit and center-align all other fields
            best_size = font_size
            size = font_size
            while size >= 10:
                # Use a dummy page to check fitting
                dummy_doc = fitz.open()
                dummy_page = dummy_doc.new_page(width=page.rect.width, height=page.rect.height)
                result = dummy_page.insert_textbox(
                    rect,
                    text,
                    fontsize=size,
                    fontname=fontname,
                    color=color,
                    align=1  # center aligned
                )
                dummy_doc.close()
                print(f"Trying {field} at size {size}, result: {result}")
                if result == 0:
                    best_size = size
                    break
                best_size = size
                size -= 1
            print(f"Drawing {field}: '{text}' at font size {best_size}")
            page.insert_textbox(
                rect,
                text,
                fontsize=best_size,
                fontname=fontname,
                color=color,
                align=1
            )

    doc.save(output_pdf_path)
    print(f"✅ Certificate saved at: {output_pdf_path}")

# Example usage
if __name__ == "__main__":
    word_path = "generate_certificate/form.docx"
    template_path = "generate_certificate/Draft.pdf"
    output_path = "Generated_Certificate.pdf"

    values = parse_word_form(word_path)
    print(values)
    generate_certificate(template_path, output_path, values)
