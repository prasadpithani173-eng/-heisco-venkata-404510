import re
import datetime
import pandas as pd
from typing import Dict, List, Any, Tuple

def normalize_col(name: Any) -> str:
    """Normalize column name for fuzzy matching."""
    if not isinstance(name, str):
        name = str(name) if name is not None else ""
    return re.sub(r'[^a-z0-9]', '', name.lower())

def parse_date_value(val: Any) -> str:
    """Parse various date representations to DD/MM/YYYY string format."""
    if val is None or pd.isna(val):
        return ""
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.strftime("%d/%m/%Y")
    
    val_str = str(val).strip()
    if not val_str:
        return ""
    
    # Try common formats
    formats = [
        "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y",
        "%d-%b-%y", "%d-%b-%Y", "%d/%b/%Y", "%d-%B-%Y",
        "%m/%d/%Y", "%m/%d/%y"
    ]
    for fmt in formats:
        try:
            dt = datetime.datetime.strptime(val_str, fmt)
            return dt.strftime("%d/%m/%Y")
        except ValueError:
            continue
            
    # Try dateutil parser if string has month names
    try:
        dt = pd.to_datetime(val_str, dayfirst=True)
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return val_str

def format_date_for_jsl(val: Any) -> str:
    """Format date as DD-Mon-YY (e.g. 22-Aug-26) matching JSL Log Sheet format."""
    if val is None or pd.isna(val):
        return ""
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.strftime("%d-%b-%y")
    
    val_str = str(val).strip()
    if not val_str:
        return ""
    
    try:
        dt = pd.to_datetime(val_str, dayfirst=True)
        return dt.strftime("%d-%b-%y")
    except Exception:
        return val_str

def read_weekly_observation_excel(file_path: str) -> List[Dict[str, Any]]:
    """Read weekly observation excel file and return standardized list of records."""
    df = pd.read_excel(file_path)
    
    # Map columns based on normalized names
    col_map = {}
    for col in df.columns:
        norm = normalize_col(col)
        if norm in ['sn', 'sno', 'no', 'number', 'id']:
            col_map['sn'] = col
        elif norm in ['area', 'location', 'site', 'worklocation', 'facility']:
            col_map['area'] = col
        elif norm in ['finding', 'observation', 'observationdetails', 'description', 'findings']:
            col_map['finding'] = col
        elif norm in ['correctiveaction', 'recommendation', 'actionrequired', 'actiontaken', 'correctiveactiontaken']:
            col_map['corrective_action'] = col
        elif norm in ['observationdate', 'date', 'obsdate', 'dateofobservation', 'inspectiondate']:
            col_map['observation_date'] = col
        elif norm in ['responsibleperson', 'responsibleentity', 'responsible', 'assignedto', 'actionby']:
            col_map['responsible_person'] = col
        elif norm in ['category', 'observationcategory', 'hazardcategory', 'natureofhazard']:
            col_map['category'] = col
        elif norm in ['status', 'actionstatus']:
            col_map['status'] = col
        elif norm in ['dateclosed', 'closeoutdate', 'closedate', 'closeddate']:
            col_map['date_closed'] = col
        elif norm in ['responsiblecompany', 'company']:
            col_map['responsible_company'] = col
        elif norm in ['assigneecompany', 'contractor', 'actionparty']:
            col_map['assignee_company'] = col
        elif norm in ['type', 'observationtype', 'hazardtype', 'classification']:
            col_map['observation_type'] = col
        elif norm in ['remarks', 'remark', 'comments', 'notes']:
            col_map['remarks'] = col

    records = []
    for idx, row in df.iterrows():
        # Check if row is empty
        finding_val = str(row.get(col_map.get('finding', ''), '')).strip()
        area_val = str(row.get(col_map.get('area', ''), '')).strip()
        if not finding_val and not area_val:
            continue
        
        raw_date = row.get(col_map.get('observation_date', ''), '')
        std_date = parse_date_value(raw_date)
        
        raw_close_date = row.get(col_map.get('date_closed', ''), '')
        std_close_date = parse_date_value(raw_close_date) if raw_close_date else std_date
        
        # Raw row values
        sn_val = row.get(col_map.get('sn', ''), idx + 1)
        if pd.isna(sn_val) or sn_val == '':
            sn_val = idx + 1
        else:
            try:
                sn_val = int(float(sn_val))
            except Exception:
                sn_val = str(sn_val).strip()

        resp_person = str(row.get(col_map.get('responsible_person', ''), '')).strip()
        if resp_person == 'nan':
            resp_person = ''
            
        category = str(row.get(col_map.get('category', ''), 'General')).strip()
        if category == 'nan':
            category = 'General'
            
        status = str(row.get(col_map.get('status', ''), 'Closed')).strip()
        if status == 'nan' or not status:
            status = 'Closed'
            
        resp_company = str(row.get(col_map.get('responsible_company', ''), 'HEISCO')).strip()
        if resp_company == 'nan' or not resp_company:
            resp_company = 'HEISCO'
            
        assignee_company = str(row.get(col_map.get('assignee_company', ''), resp_company)).strip()
        if assignee_company == 'nan' or not assignee_company:
            assignee_company = resp_company
            
        obs_type = str(row.get(col_map.get('observation_type', ''), '')).strip()
        if obs_type == 'nan':
            obs_type = ''
            
        remarks = str(row.get(col_map.get('remarks', ''), '')).strip()
        if remarks == 'nan':
            remarks = ''

        record = {
            'original_sn': sn_val,
            'area': area_val,
            'finding': finding_val,
            'corrective_action': str(row.get(col_map.get('corrective_action', ''), '')).strip(),
            'raw_observation_date': raw_date,
            'observation_date': std_date,
            'responsible_person': resp_person,
            'category': category,
            'status': status,
            'date_closed': std_close_date,
            'responsible_company': resp_company,
            'assignee_company': assignee_company,
            'observation_type': obs_type,
            'remarks': remarks
        }
        records.append(record)
        
    return records


def read_names_locations_excel(file_path: str) -> Dict[str, Dict[str, str]]:
    """
    Read Names + Locations excel file.
    Expected columns:
    • LOCATION
    • SAFETY OFFICER
    • SUPERVISOR
    • ENGINEER
    Returns mapping keyed by normalized location.
    """
    df = pd.read_excel(file_path)
    
    loc_col = None
    so_col = None
    sup_col = None
    eng_col = None
    pm_col = None
    
    for col in df.columns:
        norm = normalize_col(col)
        if norm in ['location', 'area', 'site', 'worklocation']:
            loc_col = col
        elif 'safety' in norm or 'officer' in norm or norm in ['observedby', 'so']:
            so_col = col
        elif 'supervisor' in norm or norm in ['statusby', 'sup', 'sitesupervisor']:
            sup_col = col
        elif 'engineer' in norm or norm in ['reviewedby', 'eng', 'siteengineer']:
            eng_col = col
        elif 'projectmanager' in norm or norm in ['pm', 'manager']:
            pm_col = col

    mapping = {}
    
    for _, row in df.iterrows():
        loc_raw = str(row.get(loc_col, '')).strip() if loc_col else ''
        if not loc_raw or loc_raw.lower() == 'nan':
            continue
            
        so = str(row.get(so_col, '')).strip() if so_col else ''
        sup = str(row.get(sup_col, '')).strip() if sup_col else ''
        eng = str(row.get(eng_col, '')).strip() if eng_col else ''
        pm = str(row.get(pm_col, '')).strip() if pm_col else ''
        
        mapping_entry = {
            'location_display': loc_raw,
            'safety_officer': so if so != 'nan' else '',
            'supervisor': sup if sup != 'nan' else '',
            'engineer': eng if eng != 'nan' else '',
            'project_manager': pm if pm != 'nan' else ''
        }
        
        # Store under normalized key
        from services.mapping import normalize_location_key
        key = normalize_location_key(loc_raw)
        mapping[key] = mapping_entry
        
    return mapping
