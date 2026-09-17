import os
from typing import List, Dict, Any
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.drawing.image import Image as OpenpyxlImage
from services.excel_reader import format_date_for_jsl

def generate_jsl_log_excel(
    processed_observations: List[Dict[str, Any]], 
    output_path: str,
    project_name: str = "ABQ-DBN Project",
    project_number: str = "BI-10-10303",
    last_update_date: str = ""
) -> str:
    """
    Generate the HSE Action Tracking Register / JSL Log Sheet Excel file.
    Matches the exact requirements:
    A. Header Logos:
       - Left: Enppi logo
       - Right: Aramco Overseas Company logo
    B. Title Block:
       - Yellow cell with bold text: HSE Action Tracking Register
    C. Project Details:
       - Project Name: ABQ-DBN Project
       - Project Number: BI-10-10303
    D. Summary Table:
       - Last Update
       - Total No. of Action
       - Total No. of Closed Action
       - Total No. of Open Action
    E. Pie chart and color coding.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Action Tracking Register"
    
    # Enable gridlines
    ws.views.sheetView[0].showGridLines = True

    # Styling definitions
    font_family = "Calibri"
    
    # Borders
    thin_border_side = Side(border_style="thin", color="D3D3D3")
    dark_border_side = Side(border_style="thin", color="000000")
    
    thin_cell_border = Border(
        left=thin_border_side, right=thin_border_side,
        top=thin_border_side, bottom=thin_border_side
    )
    card_border = Border(
        left=dark_border_side, right=dark_border_side,
        top=dark_border_side, bottom=dark_border_side
    )
    
    # Fills
    fill_yellow_title = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid") # User specified: Yellow cell
    fill_header = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid") # Salmon / Light orange
    fill_card_label = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    fill_status_closed = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid") # Soft green
    font_status_closed = Font(name=font_family, size=9.5, bold=True, color="006100")
    
    fill_status_open = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid") # Soft red/coral
    font_status_open = Font(name=font_family, size=9.5, bold=True, color="9C0006")

    # Metrics calculation
    total_actions = len(processed_observations)
    closed_actions = sum(1 for item in processed_observations if str(item.get('status', '')).lower() == 'closed')
    open_actions = total_actions - closed_actions
    
    if not last_update_date and processed_observations:
        last_update_date = format_date_for_jsl(processed_observations[-1].get('observation_date', ''))
    elif last_update_date:
        last_update_date = format_date_for_jsl(last_update_date)

    # 1. Row 1: Header Logos (Left: Enppi, Right: Aramco Overseas Company)
    ws.row_dimensions[1].height = 42
    
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    enppi_path = os.path.join(base_dir, 'static', 'logos', 'enppi.png')
    heisco_path = os.path.join(base_dir, 'static', 'logos', 'heisco.png')
    aramco_path = os.path.join(base_dir, 'static', 'logos', 'aramco.png')
    
    if os.path.exists(enppi_path):
        img_enppi = OpenpyxlImage(enppi_path)
        img_enppi.width = 160
        img_enppi.height = 42
        ws.add_image(img_enppi, 'B1')

    if os.path.exists(heisco_path):
        img_heisco = OpenpyxlImage(heisco_path)
        img_heisco.width = 120
        img_heisco.height = 42
        ws.add_image(img_heisco, 'G1')
        
    if os.path.exists(aramco_path):
        img_aramco = OpenpyxlImage(aramco_path)
        img_aramco.width = 160
        img_aramco.height = 42
        ws.add_image(img_aramco, 'L1')

    # 2. Row 3: Title Block (Yellow cell with bold text: HSE Action Tracking Register)
    ws.merge_cells("B3:N3")
    title_cell = ws["B3"]
    title_cell.value = "HSE Action Tracking Register"
    title_cell.font = Font(name=font_family, size=14, bold=True, color="000000")
    title_cell.fill = fill_yellow_title
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.border = card_border
    ws.row_dimensions[3].height = 28

    # 3. Summary Table (Rows 5-8, Cols B & C)
    summary_data = [
        ("Last Update", last_update_date),
        ("Total No. of Action", total_actions),
        ("Total No. of Closed Action", closed_actions),
        ("Total No. of Open Action", open_actions)
    ]
    
    for i, (label, val) in enumerate(summary_data, start=5):
        ws.row_dimensions[i].height = 20
        c_lbl = ws.cell(row=i, column=2, value=label)
        c_lbl.font = Font(name=font_family, size=9.5, bold=True)
        c_lbl.fill = fill_card_label
        c_lbl.alignment = Alignment(horizontal="left", vertical="center")
        c_lbl.border = card_border
        
        c_val = ws.cell(row=i, column=3, value=val)
        c_val.font = Font(name=font_family, size=9.5, bold=True)
        c_val.alignment = Alignment(horizontal="center", vertical="center")
        c_val.border = card_border
        
        # Color closed count with soft green
        if "Closed" in label:
            c_val.fill = fill_status_closed
            c_val.font = font_status_closed
        elif "Open" in label:
            if open_actions > 0:
                c_val.fill = fill_status_open
                c_val.font = font_status_open
            else:
                c_val.fill = fill_status_closed
                c_val.font = font_status_closed

    # 4. Project Details (Rows 5-6, Cols L & M)
    ws.cell(row=5, column=12, value="Project Name").font = Font(name=font_family, size=9.5, bold=True)
    ws.cell(row=5, column=12).fill = fill_card_label
    ws.cell(row=5, column=12).border = card_border
    ws.cell(row=5, column=13, value=project_name).font = Font(name=font_family, size=9.5, bold=True)
    ws.cell(row=5, column=13).border = card_border

    ws.cell(row=6, column=12, value="Project Number").font = Font(name=font_family, size=9.5, bold=True)
    ws.cell(row=6, column=12).fill = fill_card_label
    ws.cell(row=6, column=12).border = card_border
    ws.cell(row=6, column=13, value=project_number).font = Font(name=font_family, size=9.5, bold=True)
    ws.cell(row=6, column=13).border = card_border

    # 5. Data reference for Pie Chart (Placed in non-printing column Z & AA)
    ws['Z1'] = "Status"
    ws['AA1'] = "Count"
    ws['Z2'] = "Closed"
    ws['AA2'] = closed_actions
    ws['Z3'] = "Open"
    ws['AA3'] = open_actions

    # Pie Chart
    pie = PieChart()
    labels = Reference(ws, min_col=26, min_row=2, max_row=3)
    data = Reference(ws, min_col=27, min_row=1, max_row=3)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    pie.title = "Action Status Summary"
    pie.width = 11.5
    pie.height = 5.2
    
    # Custom slice colors: Closed=Green, Open=Coral
    if len(pie.series) > 0:
        s = pie.series[0]
        dp_closed = DataPoint(idx=0)
        dp_closed.graphicalProperties.solidFill = "70AD47" # Green
        dp_open = DataPoint(idx=1)
        dp_open.graphicalProperties.solidFill = "ED7D31"   # Orange/Coral
        s.data_points = [dp_closed, dp_open]
        
    ws.add_chart(pie, "F4")

    # Table Header Row (Row 11)
    headers = [
        "S.N",
        "Inspection ID",
        "Report From",
        "Date",
        "Location",
        "Observation Details",
        "Action Required",
        "Observation By",
        "Observation Type",
        "Corrective Action By",
        "Close Out Date",
        "Observation Category",
        "Status",
        "Remarks"
    ]
    
    header_row_idx = 11
    ws.row_dimensions[header_row_idx].height = 26

    for col_idx, header_text in enumerate(headers, start=1):
        cell = ws.cell(row=header_row_idx, column=col_idx, value=header_text)
        cell.font = Font(name=font_family, size=9.5, bold=True, color="000000")
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = card_border

    # Data Rows (Row 12+)
    for row_idx, item in enumerate(processed_observations, start=12):
        ws.row_dimensions[row_idx].height = 32
        
        obs_date_fmt = format_date_for_jsl(item.get('observation_date', ''))
        close_date_fmt = format_date_for_jsl(item.get('date_closed', ''))
        
        status_str = str(item.get('status', 'Closed')).strip()
        is_closed = status_str.lower() == 'closed'
        
        row_values = [
            item.get('sn', row_idx - 11),
            item.get('inspection_id', f"HSCO-OB-{row_idx - 11:02d}"),
            "Site Observation",
            obs_date_fmt,
            item.get('area', ''),
            item.get('finding', ''),
            item.get('corrective_action', ''),
            item.get('observed_by', ''),
            item.get('jsl_observation_type', 'Unsafe Condition'),
            item.get('assignee_company', 'HEISCO'),
            close_date_fmt,
            item.get('category', 'General'),
            status_str,
            item.get('remarks', '')
        ]

        for col_idx, val in enumerate(row_values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = Font(name=font_family, size=9)
            cell.border = thin_cell_border
            
            # Alignments
            if col_idx in [1, 2, 3, 4, 11, 13]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in [5, 8, 9, 10, 12]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                
            # Status Badge Styling
            if col_idx == 13:
                if is_closed:
                    cell.fill = fill_status_closed
                    cell.font = font_status_closed
                else:
                    cell.fill = fill_status_open
                    cell.font = font_status_open

    # Column Widths
    col_widths = {
        1: 6,    # S.N
        2: 15,   # Inspection ID
        3: 15,   # Report From
        4: 13,   # Date
        5: 18,   # Location
        6: 45,   # Observation Details
        7: 45,   # Action Required
        8: 22,   # Observation By
        9: 18,   # Observation Type
        10: 18,  # Corrective Action By
        11: 14,  # Close Out Date
        12: 24,  # Observation Category
        13: 12,  # Status
        14: 15   # Remarks
    }
    
    for col_idx, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
    return output_path

