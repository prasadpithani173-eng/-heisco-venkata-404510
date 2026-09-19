import os
from typing import List, Dict, Any
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.section import WD_ORIENTATION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_cell_background(cell, fill_hex: str):
    """Set background color of a docx table cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_padding(cell, top=100, bottom=100, left=100, right=100):
    """Set inner margins (padding) of a cell in dxa (1 pt = 20 dxa)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)

def set_table_borders(table, color="000000", sz="4", val="single"):
    """Apply borders to a docx table."""
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def set_row_height(row, height_pt: float):
    """Set minimum row height in points."""
    trPr = row._tr.get_or_add_trPr()
    h_dxa = int(height_pt * 20)
    trHeight = parse_xml(f'<w:trHeight {nsdecls("w")} w:val="{h_dxa}" w:hRule="atLeast"/>')
    trPr.append(trHeight)

def generate_observation_register_doc(pages_data: List[Dict[str, Any]], output_path: str) -> str:
    """
    Generate the HSE OBSERVATION REGISTER Word (.docx) document matching
    the exact HEISCO + Aramco reference layout:
    - Full A4 width (Landscape orientation, 297mm x 210mm)
    - 3 observations per page with page breaks
    - Header: HEISCO logo text
    - Title banner: HSE OBSERVATION REGISTER in shaded box (10.75 in wide)
    - Metadata: Ref No: 433001/HSE-OR/XX/2026, Date: DD/MM/YYYY, CC Name: DEBOTTLENECK...
    - Table: 8 columns, Aramco + HEISCO styling, exact widths totaling 10.75 in
    - Signatures:
      * Observed By: Safety officer (Name: SAFETY OFFICER from Names+Locations)
      * Status By: Supervisor (Name: SUPERVISOR from Names+Locations)
      * Reviewed By: Engineer (Name: ENGINEER from Names+Locations)
    - Document footer: EWI 107 Att.2 Rev.1 04 Sep. 2019
    """
    doc = Document()
    
    # Configure A4 Landscape with exact margins matching sample PDF
    section = doc.sections[0]
    section.orientation = WD_ORIENTATION.LANDSCAPE
    section.page_width = Inches(11.693)   # 297 mm
    section.page_height = Inches(8.268)   # 210 mm
    section.top_margin = Inches(0.35)
    section.bottom_margin = Inches(0.35)
    section.left_margin = Inches(0.45)
    section.right_margin = Inches(0.45)

    # Set normal style font to Arial 9 pt black
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Arial'
    normal_style.font.size = Pt(9)
    normal_style.font.color.rgb = RGBColor(0, 0, 0)

    # Target full table width across margins
    FULL_TABLE_WIDTH = Inches(10.75)
    
    # Column widths for the 8 columns (Sum = 10.75 inches)
    col_widths = [
        Inches(0.55),  # NO:
        Inches(1.20),  # LOCATION
        Inches(2.85),  # OBSERVATION
        Inches(2.85),  # RECOMMENDATION
        Inches(1.15),  # RESPONSIBLE ENTITY
        Inches(0.85),  # DUE DATE
        Inches(0.65),  # TYPE
        Inches(0.65)   # STATUS
    ]
    
    headers = [
        "NO:", "LOCATION", "OBSERVATION", "RECOMMENDATION",
        "RESPONSIBLE\nENTITY", "DUE DATE", "TYPE", "STATUS"
    ]

    total_pages = len(pages_data)

    for page_idx, page in enumerate(pages_data):
        # 1. Company Name Header: HEISCO
        p_logo = doc.add_paragraph()
        p_logo.paragraph_format.space_before = Pt(0)
        p_logo.paragraph_format.space_after = Pt(2)
        run_logo = p_logo.add_run("HEISCO")
        run_logo.font.name = 'Arial'
        run_logo.font.size = Pt(13)
        run_logo.font.bold = True

        # 2. Centered Title Banner (Single cell table spanning full A4 width 10.75 in)
        title_table = doc.add_table(rows=1, cols=1)
        title_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        title_table.autofit = False
        set_table_borders(title_table, color="000000", sz="6", val="single")
        
        cell_title = title_table.cell(0, 0)
        cell_title.width = FULL_TABLE_WIDTH
        set_cell_background(cell_title, "E2EFDA")  # Pale sage green Aramco/HEISCO shading
        set_cell_padding(cell_title, top=60, bottom=60, left=100, right=100)
        
        p_title = cell_title.paragraphs[0]
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_title.paragraph_format.space_before = Pt(2)
        p_title.paragraph_format.space_after = Pt(2)
        run_title = p_title.add_run("HSE OBSERVATION REGISTER")
        run_title.font.name = 'Arial'
        run_title.font.size = Pt(11)
        run_title.font.bold = True

        # 3. Ref No, Date, and CC Name block
        meta_table = doc.add_table(rows=2, cols=2)
        meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        meta_table.autofit = False
        
        # Row 0: Ref No (Left) and Date (Right)
        cell_ref = meta_table.cell(0, 0)
        cell_ref.width = Inches(7.50)
        p_ref = cell_ref.paragraphs[0]
        p_ref.paragraph_format.space_before = Pt(3)
        p_ref.paragraph_format.space_after = Pt(1)
        run_ref_lbl = p_ref.add_run("Ref No: ")
        run_ref_lbl.font.bold = True
        run_ref_lbl.font.size = Pt(9.5)
        run_ref_val = p_ref.add_run(page.get('ref_no', ''))
        run_ref_val.font.bold = True
        run_ref_val.font.size = Pt(9.5)

        cell_date = meta_table.cell(0, 1)
        cell_date.width = Inches(3.25)
        p_date = cell_date.paragraphs[0]
        p_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_date.paragraph_format.space_before = Pt(3)
        p_date.paragraph_format.space_after = Pt(1)
        run_date_lbl = p_date.add_run("Date: ")
        run_date_lbl.font.bold = True
        run_date_lbl.font.size = Pt(9.5)
        run_date_val = p_date.add_run(page.get('page_date', ''))
        run_date_val.font.bold = True
        run_date_val.font.size = Pt(9.5)

        # Row 1: CC Name spanning full 10.75 in
        cell_cc = meta_table.cell(1, 0)
        cell_cc.merge(meta_table.cell(1, 1))
        p_cc = cell_cc.paragraphs[0]
        p_cc.paragraph_format.space_before = Pt(1)
        p_cc.paragraph_format.space_after = Pt(4)
        run_cc_lbl = p_cc.add_run("CC Name: ")
        run_cc_lbl.font.bold = True
        run_cc_lbl.font.size = Pt(9.5)
        run_cc_val = p_cc.add_run(page.get('cc_name', ''))
        run_cc_val.font.bold = True
        run_cc_val.font.size = Pt(9.5)

        # 4. Main Observations Table (HEISCO + Aramco style layout)
        obs_items = page.get('observations', [])
        num_rows = 1 + len(obs_items)
        obs_table = doc.add_table(rows=num_rows, cols=8)
        obs_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        obs_table.autofit = False
        set_table_borders(obs_table, color="000000", sz="4", val="single")

        # Style Header Row
        hdr_row = obs_table.rows[0]
        hdr_row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        set_row_height(hdr_row, 24.0)
        
        for col_idx, (hdr_text, width) in enumerate(zip(headers, col_widths)):
            c = hdr_row.cells[col_idx]
            c.width = width
            set_cell_background(c, "E2EFDA")  # Pale sage green matching sample
            set_cell_padding(c, top=80, bottom=80, left=60, right=60)
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(hdr_text)
            run.font.name = 'Arial'
            run.font.bold = True
            run.font.size = Pt(8.5)

        # Style Data Rows (Exactly 3 observations per page)
        for r_idx, item in enumerate(obs_items):
            row = obs_table.rows[r_idx + 1]
            row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
            # Allocate sufficient row height so 3 observations fill the page gracefully
            set_row_height(row, 68.0)
            
            row_data = [
                str(item.get('sn', '')),
                str(item.get('LOCATION') or item.get('location') or item.get('area', '')),
                str(item.get('OBSERVATION') or item.get('observation') or item.get('finding', '')),
                str(item.get('RECOMMENDATION') or item.get('recommendation') or item.get('corrective_action', '')),
                str(item.get('RESPONSIBLE_ENTITY') or item.get('RESPONSIBLE ENTITY') or item.get('responsible_entity') or item.get('responsible_person', '') or 'Site Supervisor'),
                str(item.get('DUE_DATE') or item.get('DUE DATE') or item.get('due_date') or item.get('observation_date', '')),
                str(item.get('TYPE') or item.get('type') or item.get('doc_type', 'Major')),
                str(item.get('STATUS') or item.get('status', 'Closed'))
            ]
            
            for col_idx, (val, width) in enumerate(zip(row_data, col_widths)):
                c = row.cells[col_idx]
                c.width = width
                set_cell_padding(c, top=100, bottom=100, left=90, right=90)
                c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p = c.paragraphs[0]
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                
                # Center align NO, LOCATION, RESPONSIBLE ENTITY, DUE DATE, TYPE, STATUS
                # Left align OBSERVATION and RECOMMENDATION
                if col_idx in [0, 1, 4, 5, 6, 7]:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    
                run = p.add_run(val)
                run.font.name = 'Arial'
                run.font.size = Pt(8.5)

        # 5. Spacing before Signature Blocks
        p_space = doc.add_paragraph()
        p_space.paragraph_format.space_before = Pt(8)
        p_space.paragraph_format.space_after = Pt(2)

        # 6. Signature Blocks (3 Columns spanning full width 10.75 in)
        # Col 0: Observed By = SAFETY OFFICER
        # Col 1: Status By = SUPERVISOR
        # Col 2: Reviewed By = ENGINEER
        sig_table = doc.add_table(rows=1, cols=3)
        sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        sig_table.autofit = False

        sig_widths = [Inches(3.65), Inches(3.55), Inches(3.55)]
        
        # Helper to clean out any default title names
        banned_names = {'site supervisor', 'site engineer', 'project manager', 'safety officer', 'nan', 'none'}
        
        so_name = (page.get('Dynamic_Raised_By') or page.get('observed_by', '')).strip()
        if so_name.lower() in banned_names:
            so_name = ''
            
        sup_name = (page.get('Dynamic_Assignee') or page.get('status_by', '')).strip()
        if sup_name.lower() in banned_names:
            sup_name = ''
            
        eng_name = 'AHMED GHALWASH'

        # Col 0: Observed By: Safety officer
        c0 = sig_table.cell(0, 0)
        c0.width = sig_widths[0]
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_before = Pt(0)
        p0.paragraph_format.space_after = Pt(2)
        r0_line = p0.add_run("_____________________________\n")
        r0_line.font.bold = True
        r0_line.font.color.rgb = RGBColor(0, 0, 0)
        
        r0_lbl = p0.add_run("Observed By: Safety officer\n")
        r0_lbl.font.bold = True
        r0_lbl.font.size = Pt(9)
        
        r0_name = p0.add_run(f"Name: {so_name}\n")
        r0_name.font.bold = True
        r0_name.font.size = Pt(9)
        
        r0_date = p0.add_run(f"Date: {page.get('page_date', '')}\n\n")
        r0_date.font.bold = True
        r0_date.font.size = Pt(9)
        
        r0_note = p0.add_run("*Note: Type  – Minor & Major\n         Status – Open & Closed")
        r0_note.font.size = Pt(8)
        r0_note.font.italic = True

        # Col 1: Status By (Engineer if available, else Supervisor)
        c1 = sig_table.cell(0, 1)
        c1.width = sig_widths[1]
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(2)
        r1_line = p1.add_run("_____________________________\n")
        r1_line.font.bold = True
        r1_line.font.color.rgb = RGBColor(0, 0, 0)
        
        status_role = page.get('status_by_role', 'Supervisor')
        r1_lbl = p1.add_run(f"Status By: {status_role}\n")
        r1_lbl.font.bold = True
        r1_lbl.font.size = Pt(9)
        
        r1_name = p1.add_run(f"Name: {sup_name}\n")
        r1_name.font.bold = True
        r1_name.font.size = Pt(9)
        
        r1_date = p1.add_run(f"Date: {page.get('page_date', '')}")
        r1_date.font.bold = True
        r1_date.font.size = Pt(9)

        # Col 2: Reviewed By: Project Manager Name: AHMED GHALWASH (on every page)
        c2 = sig_table.cell(0, 2)
        c2.width = sig_widths[2]
        p2 = c2.paragraphs[0]
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(2)
        r2_line = p2.add_run("_____________________________\n")
        r2_line.font.bold = True
        r2_line.font.color.rgb = RGBColor(0, 0, 0)
        
        r2_lbl = p2.add_run("Reviewed By: Project Manager\n")
        r2_lbl.font.bold = True
        r2_lbl.font.size = Pt(9)
        
        r2_name = p2.add_run(f"Name: {page.get('reviewed_by') or 'AHMED GHALWASH'}\n")
        r2_name.font.bold = True
        r2_name.font.size = Pt(9)
        
        r2_date = p2.add_run(f"Date: {page.get('reviewed_by_date', '')}")
        r2_date.font.bold = True
        r2_date.font.size = Pt(9)

        # 7. Document Standard Reference Footer: EWI 107 Att.2 Rev.1 04 Sep. 2019
        p_ftr = doc.add_paragraph()
        p_ftr.paragraph_format.space_before = Pt(12)
        p_ftr.paragraph_format.space_after = Pt(0)
        r_ftr = p_ftr.add_run("EWI 107 Att.2 Rev.1 04 Sep. 2019")
        r_ftr.font.size = Pt(7.5)
        r_ftr.font.color.rgb = RGBColor(80, 80, 80)

        # Page break after every 3 observations (except last page)
        if page_idx < total_pages - 1:
            doc.add_page_break()

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)
    return output_path
