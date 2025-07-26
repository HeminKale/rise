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

    # --- Configuration ---
    color = (0, 0, 0)  # Black text
    fontname = "Times-Bold"  # Use bold font
    coords = {
        "Company Name and Address": fitz.Rect(87.9, 227.6, 578, 310),  # Combined rectangle
        "ISO Standard": fitz.Rect(194.9, 334, 460.3, 370),  # Moved up by 30 points
        "Scope": fitz.Rect(87.9, 386, 578, 475),
    }
    font_starts = {
        "Company Name and Address": 45,  # Company Name starts from 45pt
        "Scope": 35,
        "ISO Standard": 50,
    }
    # --- End Configuration ---

    for field, text in values.items():
        if field == "Company Name":
            # Handle Company Name and Address together
            company_text = text
            address_text = values.get("Address", "")
            
            # Combine text with natural spacing (single line break)
            combined_text = f"{company_text}\n{address_text}"  # \n creates ~2-3pt spacing
            
            rect = coords["Company Name and Address"]
            start_size = font_starts.get("Company Name and Address", 30)
            font_size = start_size
            
            print(f"🔍 Processing 'Company Name and Address': company='{company_text}', address='{address_text}', rect={rect}, start_size={start_size}")
            
            # Draw a colored rectangle around the combined field
            page.draw_rect(rect, color=(1, 0, 0), width=2)  # Red for combined field
            
            # PowerPoint-style centering with automatic font size reduction
            # Company Name starts at 45pt, Address starts at 13.6pt
            company_font_size = 45
            address_font_size = 13.6
            
            # Variables to store the final wrapped lines and font sizes
            final_wrapped_lines = []
            final_line_font_sizes = []
            final_company_lines = []
            final_address_lines = []
            
            # Reduce font sizes until text fits within box boundaries
            while company_font_size >= 8 and address_font_size >= 6:  # Minimum font sizes
                wrapped_lines = []
                line_font_sizes = []
                
                # Process Company Name with wrapping
                company_words = company_text.split()
                company_line = ""
                company_lines = []
                
                for word in company_words:
                    test_line = company_line + (" " if company_line else "") + word
                    font_obj = fitz.Font(fontname=fontname)
                    if font_obj.text_length(test_line, company_font_size) <= rect.width - 10:  # Leave margin
                        company_line = test_line
                    else:
                        if company_line:
                            company_lines.append(company_line)
                        company_line = word
                
                if company_line:
                    company_lines.append(company_line)
                
                # Add Company Name lines with company font size
                for line in company_lines:
                    wrapped_lines.append(line)
                    line_font_sizes.append(company_font_size)
                
                # No spacing line - Company Name and Address flow directly
                # wrapped_lines.append("")
                # line_font_sizes.append(0)
                
                # Process Address with wrapping
                address_words = address_text.split()
                address_line = ""
                address_lines = []
                
                for word in address_words:
                    test_line = address_line + (" " if address_line else "") + word
                    font_obj = fitz.Font(fontname=fontname)
                    if font_obj.text_length(test_line, address_font_size) <= rect.width - 10:  # Leave margin
                        address_line = test_line
                    else:
                        if address_line:
                            address_lines.append(address_line)
                        address_line = word
                
                if address_line:
                    address_lines.append(address_line)
                
                # Add Address lines with address font size
                for line in address_lines:
                    wrapped_lines.append(line)
                    line_font_sizes.append(address_font_size)
                
                # Calculate total height of all lines
                total_height = 0
                for i, line in enumerate(wrapped_lines):
                    if line.strip():
                        line_height = line_font_sizes[i] * 1.1  # Reduced from 1.2 to 1.1 for tighter spacing
                        total_height += line_height
                
                # Check if text fits vertically within box boundaries
                if total_height <= rect.height - 10:  # Leave margin
                    # Store the final working configuration
                    final_wrapped_lines = wrapped_lines.copy()
                    final_line_font_sizes = line_font_sizes.copy()
                    final_company_lines = company_lines.copy()
                    final_address_lines = address_lines.copy()
                    break
                
                # Reduce both font sizes proportionally
                company_font_size -= 1
                address_font_size -= 0.5  # Reduce address font size more slowly
            
            print(f"  📏 Company Name and Address: Company font size {company_font_size}, Address font size {address_font_size}")
            print(f"  📝 Company Name lines: {len(final_company_lines)}, Address lines: {len(final_address_lines)}")
            print(f"  📐 Total height: {total_height:.1f}pt, Available height: {rect.height - 10:.1f}pt")
            print(f"  🔍 Final wrapped lines count: {len(final_wrapped_lines)}")
            
            # Draw the text using the final working configuration
            if final_wrapped_lines:
                print(f"  ✅ Drawing {len(final_wrapped_lines)} lines")
                # Calculate total height and center vertically
                total_height = 0
                for i, line in enumerate(final_wrapped_lines):
                    if line.strip():
                        line_height = final_line_font_sizes[i] * 1.1  # Reduced from 1.2 to 1.1 for tighter spacing
                        total_height += line_height
                
                # Calculate starting Y position for centering
                current_y = rect.y0 + (rect.height - total_height) / 2
                
                # Draw each line centered horizontally
                for i, line in enumerate(final_wrapped_lines):
                    if line.strip():  # Only draw non-empty lines
                        line_height = final_line_font_sizes[i] * 1.1  # Reduced from 1.2 to 1.1 for tighter spacing
                        y_pos = current_y + line_height/2  # Adjust for baseline
                        center_x = (rect.x0 + rect.x1) / 2
                        line_width = font_obj.text_length(line, final_line_font_sizes[i])
                        x_pos = center_x - line_width / 2
                        
                        print(f"    📝 Drawing line {i}: '{line}' at font size {final_line_font_sizes[i]}")
                        
                        page.insert_text(
                            (x_pos, y_pos),
                            line,
                            fontsize=final_line_font_sizes[i],
                            fontname=fontname,
                            color=color
                        )
                        
                        current_y += line_height
            else:
                print(f"  ❌ No final wrapped lines found! Using fallback approach")
                # Fallback: Use the last calculated font sizes and recreate the structure
                wrapped_lines = []
                line_font_sizes = []
                
                # Process Company Name with wrapping using final font size
                company_words = company_text.split()
                company_line = ""
                company_lines = []
                
                for word in company_words:
                    test_line = company_line + (" " if company_line else "") + word
                    font_obj = fitz.Font(fontname=fontname)
                    if font_obj.text_length(test_line, company_font_size) <= rect.width - 10:
                        company_line = test_line
                    else:
                        if company_line:
                            company_lines.append(company_line)
                        company_line = word
                
                if company_line:
                    company_lines.append(company_line)
                
                # Add Company Name lines with company font size
                for line in company_lines:
                    wrapped_lines.append(line)
                    line_font_sizes.append(company_font_size)
                
                # Process Address with wrapping using final font size
                address_words = address_text.split()
                address_line = ""
                address_lines = []
                
                for word in address_words:
                    test_line = address_line + (" " if address_line else "") + word
                    font_obj = fitz.Font(fontname=fontname)
                    if font_obj.text_length(test_line, address_font_size) <= rect.width - 10:
                        address_line = test_line
                    else:
                        if address_line:
                            address_lines.append(address_line)
                        address_line = word
                
                if address_line:
                    address_lines.append(address_line)
                
                # Add Address lines with address font size
                for line in address_lines:
                    wrapped_lines.append(line)
                    line_font_sizes.append(address_font_size)
                
                print(f"  🔄 Fallback: Drawing {len(wrapped_lines)} lines")
                
                # Calculate total height and center vertically
                total_height = 0
                for i, line in enumerate(wrapped_lines):
                    if line.strip():
                        line_height = line_font_sizes[i] * 1.1
                        total_height += line_height
                
                # Calculate starting Y position for centering
                current_y = rect.y0 + (rect.height - total_height) / 2
                
                # Draw each line centered horizontally
                for i, line in enumerate(wrapped_lines):
                    if line.strip():  # Only draw non-empty lines
                        line_height = line_font_sizes[i] * 1.1
                        y_pos = current_y + line_height/2  # Adjust for baseline
                        center_x = (rect.x0 + rect.x1) / 2
                        line_width = font_obj.text_length(line, line_font_sizes[i])
                        x_pos = center_x - line_width / 2
                        
                        print(f"    📝 Fallback drawing line {i}: '{line}' at font size {line_font_sizes[i]}")
                        
                        page.insert_text(
                            (x_pos, y_pos),
                            line,
                            fontsize=line_font_sizes[i],
                            fontname=fontname,
                            color=color
                        )
                        
                        current_y += line_height
            
            # Skip Address processing since it's handled above
            continue
            
        elif field == "Address":
            # Skip Address processing since it's handled with Company Name
            continue
        
        # Handle other fields normally
        rect = coords[field]
        start_size = font_starts.get(field, 30)
        font_size = start_size
        
        print(f"🔍 Processing '{field}': text='{text}', rect={rect}, start_size={start_size}")
        
        # Draw a colored rectangle around each field for debugging
        if field == "ISO Standard":
            page.draw_rect(rect, color=(0, 0, 1), width=2)  # Blue for ISO Standard
        elif field == "Scope":
            page.draw_rect(rect, color=(1, 1, 0), width=2)  # Yellow for Scope
        
        # Reduce font size if it doesn't fit, but ensure minimum size
        while font_size >= 12:  # Increased minimum from 10 to 12
            text_height = get_text_height(text, font_size, fontname, rect.width)
            limit = rect.height if field != "Company Name" else rect.height * 2
            print(f"  📏 Font size {font_size}: text_height={text_height:.1f}, limit={limit:.1f}")
            if text_height <= limit:
                break
            font_size -= 1
        
        print(f"🖋 Writing '{field}' at font size {font_size}")
        
        if field == "Scope":
            # PowerPoint-style centering with automatic font size reduction
            original_font_size = font_size
            
            # Reduce font size until text fits within box boundaries
            while font_size >= 8:  # Minimum font size
                # Split text into lines that fit within the rectangle width
                words = text.split()
                lines = []
                current_line = ""
                
                for word in words:
                    test_line = current_line + (" " if current_line else "") + word
                    font_obj = fitz.Font(fontname=fontname)
                    if font_obj.text_length(test_line, font_size) <= rect.width - 10:  # Leave margin
                        current_line = test_line
                    else:
                        if current_line:
                            lines.append(current_line)
                        current_line = word
                
                if current_line:
                    lines.append(current_line)
                
                # Calculate total height of all lines
                line_height = font_size * 1.2
                total_height = len(lines) * line_height
                
                # Check if text fits vertically within box boundaries
                if total_height <= rect.height - 10:  # Leave margin
                    break
                
                font_size -= 1
            
            print(f"  📏 Scope: Reduced font size from {original_font_size} to {font_size}")
            
            # Now draw the text with the fitting font size
            words = text.split()
            lines = []
            current_line = ""
            
            for word in words:
                test_line = current_line + (" " if current_line else "") + word
                font_obj = fitz.Font(fontname=fontname)
                if font_obj.text_length(test_line, font_size) <= rect.width - 10:
                    current_line = test_line
                else:
                    if current_line:
                        lines.append(current_line)
                    current_line = word
            
            if current_line:
                lines.append(current_line)
            
            # Calculate total height and center vertically
            line_height = font_size * 1.2
            total_height = len(lines) * line_height
            start_y = rect.y0 + (rect.height - total_height) / 2 + line_height/2  # Adjust for baseline
            
            # Draw each line centered horizontally
            for i, line in enumerate(lines):
                y_pos = start_y + i * line_height
                center_x = (rect.x0 + rect.x1) / 2
                line_width = font_obj.text_length(line, font_size)
                x_pos = center_x - line_width / 2
                
                page.insert_text(
                    (x_pos, y_pos),
                    line,
                    fontsize=font_size,
                    fontname=fontname,
                    color=color
                )
        
        elif field == "ISO Standard":
            # Perfect centering for ISO Standard - both horizontal and vertical
            center_x = (rect.x0 + rect.x1) / 2
            center_y = (rect.y0 + rect.y1) / 2 + font_size/3  # Adjust for baseline
            
            # Calculate text width for centering
            font_obj = fitz.Font(fontname=fontname)
            text_width = font_obj.text_length(text, font_size)
            start_x = center_x - text_width / 2
            
            page.insert_text(
                (start_x, center_y),
                text,
                fontsize=font_size,
                fontname=fontname,
                color=color
            )

    doc.save(output_pdf_path)
    print(f"✅ Certificate saved at: {output_pdf_path}")

if __name__ == "__main__":
    word_path = "generate_certificate/form.docx"
    template_path = "generate_certificate/Draft.pdf"
    output_path = "Generated_Certificate.pdf"

    values = parse_word_form(word_path)
    generate_certificate(template_path, output_path, values)
