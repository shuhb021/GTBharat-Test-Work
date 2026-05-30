"""
FAR Automation Tool — Word Document Export
Generates a professional Word document with cover page, analysis tables,
and AI remarks summary using python-docx.
"""

import os
import logging
from datetime import datetime

try:
    from docx import Document
    from docx.shared import Inches, Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

logger = logging.getLogger(__name__)


def _set_cell_shading(cell, color_hex):
    """Set cell background color."""
    shading_elm = cell._element.get_or_add_tcPr()
    shading = shading_elm.makeelement(qn('w:shd'), {
        qn('w:fill'): color_hex,
        qn('w:val'): 'clear'
    })
    shading_elm.append(shading)


def _add_table_row(table, values, is_header=False, highlight=False):
    """Add a row to a Word table with styling."""
    row = table.add_row()
    for i, val in enumerate(values):
        cell = row.cells[i]
        cell.text = str(val) if val is not None else ''
        
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT if i == 0 else WD_ALIGN_PARAGRAPH.RIGHT
            for run in paragraph.runs:
                run.font.size = Pt(9)
                run.font.name = 'Segoe UI'
                if is_header:
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        
        if is_header:
            _set_cell_shading(cell, '252526')
        elif highlight:
            _set_cell_shading(cell, 'EAB308')
    
    return row


def export_far_docx(output_path, bs_result, pl_result, ratios,
                    remarks_bs, remarks_pl,
                    client_name, firm_name, financial_year,
                    rounding_unit, cy_year, py_year):
    """
    Generate the FAR Word document.
    """
    if not HAS_DOCX:
        logger.error("python-docx not installed — cannot generate Word document")
        return False
    
    doc = Document()
    
    # Set default font
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Segoe UI'
    font.size = Pt(10)
    
    # ── Cover Page ──
    doc.add_paragraph('')
    doc.add_paragraph('')
    
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run(firm_name)
    run.font.size = Pt(24)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x00, 0x7A, 0xCC)
    
    doc.add_paragraph('')
    
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run('Final Analytical Review Workpaper')
    run.font.size = Pt(18)
    run.font.bold = True
    
    doc.add_paragraph('')
    
    info_items = [
        f'Client: {client_name}',
        f'Financial Year: {financial_year}',
        f'Period: 31 March {py_year} to 31 March {cy_year}',
        f'Amounts stated in: {rounding_unit}',
        f'Generated: {datetime.now().strftime("%d %B %Y")}',
    ]
    for item in info_items:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(item)
        run.font.size = Pt(12)
    
    doc.add_page_break()
    
    # ── Balance Sheet Analysis ──
    doc.add_heading('Balance Sheet Analysis', level=1)
    doc.add_paragraph(
        f'As at 31 March {cy_year} vs 31 March {py_year}  |  (Rs. in {rounding_unit})')
    
    bs_headers = ['Particulars', 'Note', f'CY ({cy_year})', f'PY ({py_year})',
                  'Variance', 'Var %', 'Remarks']
    bs_table = doc.add_table(rows=1, cols=len(bs_headers))
    bs_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    bs_table.style = 'Table Grid'
    
    # Header row
    for i, h in enumerate(bs_headers):
        cell = bs_table.rows[0].cells[i]
        cell.text = h
        _set_cell_shading(cell, '252526')
        for p in cell.paragraphs:
            for run in p.runs:
                run.font.bold = True
                run.font.size = Pt(9)
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    
    for idx, row_data in enumerate(bs_result):
        remark = remarks_bs.get(idx, '')
        values = [
            row_data.get('particulars', ''),
            row_data.get('note', ''),
            f"{row_data.get('cy', 0):,.0f}",
            f"{row_data.get('py', 0):,.0f}",
            f"{row_data.get('variance_abs', 0):,.0f}",
            row_data.get('display_pct', ''),
            remark
        ]
        _add_table_row(bs_table, values,
                      highlight=row_data.get('flag', False))
    
    doc.add_page_break()
    
    # ── Profit & Loss Analysis ──
    doc.add_heading('Profit & Loss Analysis', level=1)
    doc.add_paragraph(
        f'For the year ended 31 March {cy_year} vs 31 March {py_year}  |  (Rs. in {rounding_unit})')
    
    pl_table = doc.add_table(rows=1, cols=len(bs_headers))
    pl_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    pl_table.style = 'Table Grid'
    
    for i, h in enumerate(bs_headers):
        cell = pl_table.rows[0].cells[i]
        cell.text = h
        _set_cell_shading(cell, '252526')
        for p in cell.paragraphs:
            for run in p.runs:
                run.font.bold = True
                run.font.size = Pt(9)
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    
    for idx, row_data in enumerate(pl_result):
        remark = remarks_pl.get(idx, '')
        values = [
            row_data.get('particulars', ''),
            row_data.get('note', ''),
            f"{row_data.get('cy', 0):,.0f}",
            f"{row_data.get('py', 0):,.0f}",
            f"{row_data.get('variance_abs', 0):,.0f}",
            row_data.get('display_pct', ''),
            remark
        ]
        _add_table_row(pl_table, values,
                      highlight=row_data.get('flag', False))
    
    doc.add_page_break()
    
    # ── Ratio Analysis ──
    doc.add_heading('Ratio Analysis', level=1)
    doc.add_paragraph(
        f'As at 31 March {cy_year} vs 31 March {py_year}')
    
    ratio_headers = ['Ratio', 'Formula', 'CY', 'PY', 'Change %']
    ratio_table = doc.add_table(rows=1, cols=len(ratio_headers))
    ratio_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    ratio_table.style = 'Table Grid'
    
    for i, h in enumerate(ratio_headers):
        cell = ratio_table.rows[0].cells[i]
        cell.text = h
        _set_cell_shading(cell, '252526')
        for p in cell.paragraphs:
            for run in p.runs:
                run.font.bold = True
                run.font.size = Pt(9)
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    
    for ratio in ratios:
        values = [
            ratio['key'],
            ratio['formula'],
            f"{ratio.get('cy', 0):.2f}",
            f"{ratio.get('py', 0):.2f}",
            ratio.get('display_change', '')
        ]
        _add_table_row(ratio_table, values,
                      highlight=ratio.get('flag', False))
    
    # ── AI Remarks Summary ──
    if remarks_bs or remarks_pl:
        doc.add_page_break()
        doc.add_heading('AI-Generated Audit Remarks Summary', level=1)
        
        if remarks_bs:
            doc.add_heading('Balance Sheet Remarks', level=2)
            for idx, remark in sorted(remarks_bs.items()):
                if idx < len(bs_result):
                    item_name = bs_result[idx].get('particulars', f'Item {idx}')
                    p = doc.add_paragraph()
                    run = p.add_run(f'{item_name}: ')
                    run.font.bold = True
                    run.font.size = Pt(10)
                    run = p.add_run(remark)
                    run.font.size = Pt(10)
        
        if remarks_pl:
            doc.add_heading('Profit & Loss Remarks', level=2)
            for idx, remark in sorted(remarks_pl.items()):
                if idx < len(pl_result):
                    item_name = pl_result[idx].get('particulars', f'Item {idx}')
                    p = doc.add_paragraph()
                    run = p.add_run(f'{item_name}: ')
                    run.font.bold = True
                    run.font.size = Pt(10)
                    run = p.add_run(remark)
                    run.font.size = Pt(10)
    
    # Save
    doc.save(output_path)
    logger.info("FAR Word document saved to %s", output_path)
    return True
