"""
Document populator: fills a Word template with structured content and Unicode checkboxes.
"""
import os
import sys
import shutil
import tempfile
import traceback
from datetime import datetime
from typing import Dict, List, Any, Optional
from pathlib import Path

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from PIL import Image as PILImage
import pillow_avif  # noqa: F401 - registers AVIF format with Pillow


def _set_cell_shading(cell, fill_color: str):
    """Set cell background shading (e.g., 'D9E2F3' for light blue)."""
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), fill_color)
    cell._tc.get_or_add_tcPr().append(shading)


def _fix_doc_styles(doc: Document):
    """Override Word's built-in style defaults so spacing/font sizes stick."""
    from docx.shared import Pt as _Pt
    from docx.enum.text import WD_LINE_SPACING

    title_style = doc.styles['Title']
    title_style.paragraph_format.space_before = _Pt(0)
    title_style.paragraph_format.space_after = _Pt(0)
    title_style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    title_style.font.size = _Pt(20)
    title_style.font.bold = True
    title_style.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    heading1_style = doc.styles['Heading 1']
    heading1_style.paragraph_format.space_before = _Pt(6)
    heading1_style.paragraph_format.space_after = _Pt(0)
    heading1_style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    heading1_style.font.size = _Pt(13)
    heading1_style.font.bold = True
    heading1_style.font.color.rgb = RGBColor(0x00, 0x00, 0x00)


def _set_run_font(run, font_name: str = "Calibri", font_size: int = 11, bold: bool = False):
    """Set font properties on a run."""
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.bold = bold


def _replace_placeholder_text(doc: Document, placeholder: str, replacement: str):
    """Replace placeholder text anywhere in the document (paragraphs and tables)."""
    for para in doc.paragraphs:
        if placeholder in para.text:
            para.text = para.text.replace(placeholder, replacement)
            for run in para.runs:
                _set_run_font(run)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if placeholder in para.text:
                        para.text = para.text.replace(placeholder, replacement)
                        for run in para.runs:
                            _set_run_font(run)


def _find_paragraph_containing(doc: Document, text: str):
    """Find first paragraph containing the given text."""
    for para in doc.paragraphs:
        if text in para.text:
            return para
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if text in para.text:
                        return para
    return None


def _insert_paragraph_after(paragraph, text: str, style: Optional[str] = None):
    """Insert a new paragraph after the given paragraph object."""
    new_p = OxmlElement("w:p")
    paragraph._element.addnext(new_p)
    new_para = paragraph._parent._element[paragraph._parent._element.index(paragraph._element) + 1]
    # We need to wrap the new element in a Paragraph object
    from docx.text.paragraph import Paragraph
    return Paragraph(new_para, paragraph._parent)


def _render_checkbox(checked: bool) -> str:
    return "☑" if checked else "☐"


def _prepare_image(image_path: str) -> str:
    """Prepare image for insertion. Convert AVIF/WebP to PNG if needed."""
    path = Path(image_path)
    # Sanitize filename: remove spaces for safer handling
    if " " in str(path):
        safe_path = Path(tempfile.gettempdir()) / path.name.replace(" ", "_")
        shutil.copy2(str(path), str(safe_path))
        path = safe_path

    # Try to open with PIL to determine actual format
    try:
        img = PILImage.open(str(path))
        actual_format = img.format
    except Exception:
        actual_format = None

    # Convert non-native formats to PNG
    non_native = {"AVIF", "WEBP", "BMP", "GIF", "TIFF", "TGA"}
    if actual_format and actual_format.upper() in non_native:
        png_path = Path(tempfile.gettempdir()) / (path.stem + ".png")
        img.save(str(png_path), "PNG")
        return str(png_path)

    return str(path)


def _insert_images_after_heading(doc: Document, heading_text: str, images: List[str], max_width_inches: float = 5.5):
    """Insert images after the first paragraph containing heading_text."""
    from docx.text.paragraph import Paragraph

    anchor = _find_paragraph_containing(doc, heading_text)
    if not anchor:
        print(f"Warning: could not find heading '{heading_text}' for image insertion")
        return

    # Find the content paragraph after the heading (the placeholder)
    # We'll insert images after that paragraph
    content_para = None
    found_heading = False
    for para in doc.paragraphs:
        if found_heading:
            content_para = para
            break
        if heading_text in para.text:
            found_heading = True

    # If no content paragraph found, use the heading itself
    insert_after = content_para._element if content_para else anchor._element

    last_element = insert_after
    for img_path in images:
        prepared = _prepare_image(img_path)
        try:
            # Verify PIL can open it
            PILImage.open(prepared)

            # Add picture at end of document (handles relationships correctly)
            doc.add_picture(prepared, width=Inches(max_width_inches))
            pic_para = doc.paragraphs[-1]
            pic_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # Move picture paragraph to after last_element
            last_element.addnext(pic_para._element)
            last_element = pic_para._element

            # Add spacing paragraph after image
            doc.add_paragraph()
            spacing_para = doc.paragraphs[-1]
            spacing_para.paragraph_format.space_after = Pt(6)
            last_element.addnext(spacing_para._element)
            last_element = spacing_para._element

        except Exception as e:
            print(f"Warning: could not insert image {img_path}: {e}")
            traceback.print_exc()


def _build_distribution_table(doc: Document, data: Dict[str, Any]):
    """Build the Distribution Information table with exact Google structure, Unicode checkboxes, and Assessment Questions."""
    distribution = data.get("distribution", {})
    source_channels = distribution.get("source_channels", [])
    distribution_channels = distribution.get("distribution_channels", [])
    distribution_methods = distribution.get("distribution_methods", [])
    status_date = distribution.get("status_date", datetime.now().strftime("%Y-%m-%d"))
    resource_links = distribution.get("resource_links", "N/A")
    approval = data.get("approval_status", [])
    assessment_questions = data.get("assessment_questions", [])

    # Exact Google names from template
    all_source = ["Social", "Voice Call", "Chat", "Community", "Repair Centers", "Global / HQ"]
    all_dist = [
        "Social/Community", "Reactive (DM)", "Proactive", "Voice Call",
        "Chat", "Repair Centers", "Carrier", "Retail", "E-Com CS"
    ]
    all_methods = [
        "SSAM (AI recommendation engine)", "Training (Classroom)",
        "Support S.com (Customer Facing)", "Vendor Huddle (Conf Call)",
        "Alert Communication (Email)", "System Modification (Pop-Up)",
        "IVR Message (1800 GOOGLE)", "Retail Communication (FSM)"
    ]
    all_approval = ["CSD Internal", "Legal", "PR", "Product Mkt", "Product Mgmt", "HQ PR"]

    # Find the Distribution Information heading
    heading = _find_paragraph_containing(doc, "Distribution Information")
    if not heading:
        print("Warning: Distribution Information heading not found")
        return

    # Create table after the heading
    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"

    # Header row
    hdr_cells = table.rows[0].cells
    headers = ["Source Channel", "Distribution Channel", "Distribution Method"]
    for i, header_text in enumerate(headers):
        hdr_cells[i].text = header_text
        for para in hdr_cells[i].paragraphs:
            for run in para.runs:
                run.bold = True
                _set_run_font(run, bold=True)
        _set_cell_shading(hdr_cells[i], "D9E2F3")

    # Option rows with checkboxes
    max_rows = max(len(all_source), len(all_dist), len(all_methods))
    for row_idx in range(max_rows):
        row = table.add_row()
        cells = row.cells
        if row_idx < len(all_source):
            opt = all_source[row_idx]
            cells[0].text = f"{_render_checkbox(opt in source_channels)} {opt}"
        if row_idx < len(all_dist):
            opt = all_dist[row_idx]
            cells[1].text = f"{_render_checkbox(opt in distribution_channels)} {opt}"
        if row_idx < len(all_methods):
            opt = all_methods[row_idx]
            cells[2].text = f"{_render_checkbox(opt in distribution_methods)} {opt}"

    # Bottom section: Status Date, Approval Status, Resource Links
    # Blank row for spacing
    table.add_row()

    # Header row for bottom section
    bottom_hdr = table.add_row()
    bottom_hdr.cells[0].text = "Status Date:"
    bottom_hdr.cells[1].text = "Approval Status"
    bottom_hdr.cells[2].text = "Resource Links:"
    for cell in bottom_hdr.cells:
        for para in cell.paragraphs:
            for run in para.runs:
                run.bold = True
                run.underline = True
                _set_run_font(run, bold=True)

    # Content row: date | approval checkboxes | resource links
    bottom_row = table.add_row()
    bottom_row.cells[0].text = status_date

    # Stack all approval items vertically in middle column
    approval_lines = [f"{_render_checkbox(opt in approval)} {opt}" for opt in all_approval]
    bottom_row.cells[1].text = "\n".join(approval_lines)

    bottom_row.cells[2].text = resource_links if resource_links != "N/A" else ""

    # Assessment Questions section (within the same table)
    if assessment_questions:
        aq_blank = table.add_row()

        aq_hdr = table.add_row()
        aq_hdr.cells[0].text = "Assessment Questions:"
        for cell in aq_hdr.cells:
            for para in cell.paragraphs:
                for run in para.runs:
                    run.bold = True
                    run.underline = True
                    _set_run_font(run, bold=True)

        for aq in assessment_questions:
            q = aq.get("question", "")
            a = aq.get("answer", "")
            aq_row = table.add_row()
            aq_row.cells[0].text = f"Q: {q}"
            aq_row.cells[1].text = f"A: {a}"
            aq_row.cells[2].text = ""

    # Move table to after the heading
    heading._element.addnext(table._tbl)

    # Remove placeholder paragraphs instead of leaving them empty
    for placeholder in [
        "[Source Channel Checkboxes]",
        "[Distribution Channel Checkboxes]",
        "[Distribution Method Checkboxes]",
        "[Status Date to be Inserted]",
        "[Resource Links to be Inserted]",
        "[Approval Status Checkboxes]",
    ]:
        para = _find_paragraph_containing(doc, placeholder)
        if para is not None:
            p_element = para._element
            p_element.getparent().remove(p_element)


def populate_template(
    template_path: str,
    output_path: str,
    data: Dict[str, Any],
    images: Optional[List[str]] = None,
) -> str:
    """
    Populate a Word template with structured data.

    Args:
        template_path: Path to the master .docx template.
        output_path: Path to write the populated .docx.
        data: Structured JSON data with all template fields.
        images: Optional list of image file paths to insert.

    Returns:
        Path to the generated document.
    """
    doc = Document(template_path)

    # 1. Topic (part of the title heading)
    topic = data.get("topic", "N/A")
    _replace_placeholder_text(doc, "[Topic to be Inserted]", topic)

    # 2. Support Channels (audience section)
    support_channels = data.get("support_channels", [])
    if support_channels:
        channels_str = " / ".join(support_channels)
        _replace_placeholder_text(doc, "[Support Channels to be Inserted]", channels_str)
    else:
        _replace_placeholder_text(doc, "[Support Channels to be Inserted]", "Social / Community / Voice Channel / Chat / Email")

    # 3. Situation Overview
    situation = data.get("situation_overview", "N/A")
    _replace_placeholder_text(doc, "[Situation Overview to be Inserted]", situation)

    # 4. Approach
    approach = data.get("approach", "N/A")
    _replace_placeholder_text(doc, "[Approach to be Inserted]", approach)

    # 5. Materials
    materials = data.get("materials", "N/A")
    _replace_placeholder_text(doc, "[Materials to be Inserted]", materials)

    # 6. Messaging Statement
    messaging = data.get("messaging_statement", "N/A")
    _replace_placeholder_text(doc, "[Messaging Statement to be Inserted]", messaging)

    # 7. Social Messaging
    social = data.get("social_messaging", "N/A")
    _replace_placeholder_text(doc, "[Social Messaging to be Inserted]", social)

    # 8. Talking Points
    talking_points = data.get("talking_points", [])
    if talking_points:
        tp_text = "\n".join([f"• {tp}" for tp in talking_points])
        _replace_placeholder_text(doc, "[Talking Points to be Inserted]", tp_text)
    else:
        _replace_placeholder_text(doc, "[Talking Points to be Inserted]", "N/A")

    # 9. FAQs (prepend audience as first FAQ if provided)
    faqs = data.get("faqs", [])
    audience = data.get("audience", "")

    # Build FAQ text
    faq_lines = []

    # Prepend audience as first FAQ if available
    if audience and audience != "N/A":
        faq_lines.append(f"Q: Who is the target audience for this communication?")
        faq_lines.append(f"A: {audience}")
        faq_lines.append("")

    for faq in faqs:
        q = faq.get("question", "")
        a = faq.get("answer", "")
        faq_lines.append(f"Q: {q}")
        faq_lines.append(f"A: {a}")
        faq_lines.append("")  # blank line between FAQs

    if faq_lines:
        faq_text = "\n".join(faq_lines)
        _replace_placeholder_text(doc, "[FAQs to be Inserted]", faq_text)
    else:
        _replace_placeholder_text(doc, "[FAQs to be Inserted]", "N/A")

    # 10. Distribution Table (includes Status Date, Approval Status, Resource Links)
    _build_distribution_table(doc, data)

    # 11. Insert images under Situation Overview
    if images:
        _insert_images_after_heading(doc, "Situation Overview", images)

    # 12. Final pass: enforce consistent spacing and override style defaults
    _fix_doc_styles(doc)

    for para in doc.paragraphs:
        if para.style and para.style.name == "Title":
            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after = Pt(0)
            para.paragraph_format.line_spacing = 1.0
            for run in para.runs:
                run.font.size = Pt(20)
                run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
                run.bold = True
        elif para.style and "Heading" in para.style.name:
            para.paragraph_format.space_before = Pt(6)
            para.paragraph_format.space_after = Pt(0)
            para.paragraph_format.line_spacing = 1.0
            for run in para.runs:
                run.font.size = Pt(13)
                run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
                run.bold = True

    doc.save(output_path)
    return output_path


def create_placeholder_template(output_path: str):
    """
    Create a minimal placeholder template for Phase 1 testing.
    This mimics the Google template structure.
    """
    doc = Document()

    # Set page header: PRIVILEGED AND CONFIDENTIAL (right-aligned)
    section = doc.sections[0]
    header = section.header
    header.is_linked_to_previous = False
    header_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    header_para.text = "PRIVILEGED AND CONFIDENTIAL"
    header_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for run in header_para.runs:
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x40, 0x40, 0x40)
        run.italic = True

    # Override built-in styles so spacing/font sizes actually stick in Word
    _fix_doc_styles(doc)

    def _add_heading(doc, text, level=1):
        h = doc.add_heading(text, level=level)
        h.paragraph_format.space_before = Pt(6)
        h.paragraph_format.space_after = Pt(0)
        h.paragraph_format.line_spacing = 1.0
        for run in h.runs:
            run.font.size = Pt(13) if level == 1 else Pt(20)
            run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
            run.bold = True
        return h

    # Title: Communications Recommendations – [Topic]
    title = doc.add_heading("Communications Recommendations – [Topic to be Inserted]", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    title.paragraph_format.space_before = Pt(0)
    title.paragraph_format.space_after = Pt(0)
    title.paragraph_format.line_spacing = 1.0
    title.clear()
    run = title.add_run("Communications Recommendations – [Topic to be Inserted]")
    run.font.size = Pt(20)
    run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    run.bold = True

    # Audience heading
    _add_heading(doc, "Audience", level=1)

    # Support Channels (indented, tight spacing)
    sc_para = doc.add_paragraph()
    sc_para.paragraph_format.left_indent = Inches(0.5)
    sc_para.paragraph_format.space_after = Pt(0)
    sc_para.add_run("Support Channels: ").bold = True
    sc_para.add_run("[Support Channels to be Inserted]")

    # Situation Overview
    _add_heading(doc, "Situation Overview", level=1)
    doc.add_paragraph("[Situation Overview to be Inserted]")

    # Approach
    _add_heading(doc, "Approach", level=1)
    doc.add_paragraph("[Approach to be Inserted]")

    # Materials
    _add_heading(doc, "Materials", level=1)
    doc.add_paragraph("[Materials to be Inserted]")

    # Messaging Statement
    _add_heading(doc, "Messaging Statement", level=1)
    doc.add_paragraph("[Messaging Statement to be Inserted]")

    # Social Messaging
    _add_heading(doc, "Social Messaging", level=1)
    doc.add_paragraph("[Social Messaging to be Inserted]")

    # Talking Points
    _add_heading(doc, "Talking Points", level=1)
    doc.add_paragraph("[Talking Points to be Inserted]")

    # FAQs
    _add_heading(doc, "FAQs", level=1)
    doc.add_paragraph("[FAQs to be Inserted]")

    # Distribution Information (table includes Status Date, Approval Status, Resource Links, Assessment Questions)
    _add_heading(doc, "Distribution Information", level=1)
    doc.add_paragraph("[Source Channel Checkboxes]")
    doc.add_paragraph("[Distribution Channel Checkboxes]")
    doc.add_paragraph("[Distribution Method Checkboxes]")

    doc.save(output_path)
    return output_path
