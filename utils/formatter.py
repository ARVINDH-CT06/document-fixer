from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL

def format_document(doc):
    changes_report = []

    # 1. Format paragraphs
    para_count = 0
    for para in doc.paragraphs:
        if para.text.strip():
            for run in para.runs:
                run.font.name = "Times New Roman"
                run.font.size = Pt(12)
            para.paragraph_format.line_spacing = 1.5
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            para_count += 1
    changes_report.append(f"Formatted {para_count} paragraphs (Times New Roman, 12pt, justified)")

    # 2. Format headings
    heading_count = 0
    for para in doc.paragraphs:
        text = para.text.strip()
        if text.isupper() and len(text.split()) <= 6:
            for run in para.runs:
                run.font.size = Pt(16)
                run.font.bold = True
                run.font.color.rgb = RGBColor(0, 0, 255)
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            heading_count += 1
        elif text.istitle():
            for run in para.runs:
                run.font.size = Pt(14)
                run.font.bold = True
                run.font.color.rgb = RGBColor(0, 0, 255)
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            heading_count += 1
    changes_report.append(f"Formatted {heading_count} headings (center aligned, bold, color: blue)")

    # 3. Format tables
    table_count = 0
    for table in doc.tables:
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for row_idx, row in enumerate(table.rows):
            for cell in row.cells:
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                for para in cell.paragraphs:
                    if row_idx == 0:
                        for run in para.runs:
                            run.font.bold = True
                            run.font.name = "Times New Roman"
                            run.font.size = Pt(12)
                        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    else:
                        if para.text.strip().isdigit():
                            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        elif len(para.text) > 30:
                            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        else:
                            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        table_count += 1
    changes_report.append(f"Formatted {table_count} tables (center aligned, bold headers)")

    return doc, changes_report
