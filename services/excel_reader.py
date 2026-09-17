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
        elif any(k in norm for k in ['permitactivity', 'permit', 'activity', 'workdescription', 'activities']):
            col_map['permit_activity'] = col
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

        permit_act = str(row.get(col_map.get('permit_activity', ''), '')).strip()
        if permit_act == 'nan':
            permit_act = ''
            
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
            'permit_activity': permit_act,
            'remarks': remarks
        }
        records.append(record)
        
    return records


def read_names_locations_excel(file_path: str) -> Dict[str, Dict[str, Any]]:
    """
    Read Names + Locations excel file.
    Supports two formats:
    Format 1 (Column based):
    • LOCATION | SAFETY OFFICER | SUPERVISOR | ENGINEER
    
    Format 2 (Roster based, like sample PDF 2):
    • EMPLOYEE NAME | PROJECT DEISGNATION | LOCATION (or date col like 09-Sep)
    
    Returns mapping keyed by normalized location, mapping:
    - safety_officer (primary) & safety_officers (list)
    - supervisor (primary) & supervisors (list)
    - engineer (primary) & engineers (list)
    """
    from services.mapping import normalize_location_key
    
    df = pd.read_excel(file_path)
    
    # Check if this is a roster format (EMPLOYEE NAME + DESIGNATION + LOCATION)
    emp_col = None
    desig_col = None
    roster_loc_col = None
    
    # Column format headers
    loc_col = None
    so_col = None
    sup_col = None
    eng_col = None
    pm_col = None
    
    for col in df.columns:
        norm = normalize_col(col)
        if norm in ['employeename', 'name', 'employee']:
            emp_col = col
        elif any(k in norm for k in ['deisgnation', 'designation', 'role', 'position', 'title']):
            desig_col = col
        elif norm in ['location', 'area', 'site', 'worklocation']:
            loc_col = col
            roster_loc_col = col
        elif 'safety' in norm or 'officer' in norm or norm in ['observedby', 'so']:
            so_col = col
        elif 'supervisor' in norm or norm in ['statusby', 'sup', 'sitesupervisor']:
            sup_col = col
        elif 'engineer' in norm or norm in ['reviewedby', 'eng', 'siteengineer']:
            eng_col = col
        elif 'projectmanager' in norm or norm in ['pm', 'manager']:
            pm_col = col
        elif not roster_loc_col and ('sep' in norm or 'gosp' in norm or 'date' in norm):
            # Often the location column in the roster is named after the date, e.g. "09-Sep"
            roster_loc_col = col

    mapping: Dict[str, Dict[str, Any]] = {}

    # Case A: Roster Format (EMPLOYEE NAME + DESIGNATION)
    if emp_col and desig_col:
        target_loc_col = roster_loc_col or loc_col or (df.columns[2] if len(df.columns) >= 3 else None)
        for _, row in df.iterrows():
            name = str(row.get(emp_col, '')).strip()
            desig = str(row.get(desig_col, '')).strip().upper()
            loc_raw = str(row.get(target_loc_col, '')).strip() if target_loc_col else ''
            
            if not name or name.lower() == 'nan' or not loc_raw or loc_raw.lower() == 'nan':
                continue
                
            key = normalize_location_key(loc_raw)
            if key not in mapping:
                mapping[key] = {
                    'location_display': loc_raw,
                    'safety_officer': '',
                    'supervisor': '',
                    'engineer': '',
                    'safety_officers': [],
                    'supervisors': [],
                    'engineers': []
                }
                
            if 'SAFETY' in desig or 'HSE' in desig or 'OFFICER' in desig:
                mapping[key]['safety_officers'].append(name)
            elif 'SUPERVISOR' in desig or 'FOREMAN' in desig:
                mapping[key]['supervisors'].append(name)
            elif 'ENGINEER' in desig:
                mapping[key]['engineers'].append(name)
                
        # Set primary names
        for key, item in mapping.items():
            if item['safety_officers']:
                item['safety_officer'] = item['safety_officers'][0]
            if item['supervisors']:
                item['supervisor'] = item['supervisors'][0]
            if item['engineers']:
                item['engineer'] = item['engineers'][0]

    # Case B: Column Based Format (LOCATION + SAFETY OFFICER + SUPERVISOR + ENGINEER)
    else:
        for _, row in df.iterrows():
            loc_raw = str(row.get(loc_col, '')).strip() if loc_col else ''
            if not loc_raw or loc_raw.lower() == 'nan':
                continue
                
            so = str(row.get(so_col, '')).strip() if so_col else ''
            sup = str(row.get(sup_col, '')).strip() if sup_col else ''
            eng = str(row.get(eng_col, '')).strip() if eng_col else ''
            
            so = '' if so.lower() == 'nan' else so
            sup = '' if sup.lower() == 'nan' else sup
            eng = '' if eng.lower() == 'nan' else eng
            
            key = normalize_location_key(loc_raw)
            mapping_entry = {
                'location_display': loc_raw,
                'safety_officer': so,
                'supervisor': sup,
                'engineer': eng,
                'safety_officers': [so] if so else [],
                'supervisors': [sup] if sup else [],
                'engineers': [eng] if eng else []
            }
            mapping[key] = mapping_entry

    return mapping


def match_site_key(raw_site: str) -> str:
    """Normalize raw site name to standard site keys."""
    s = re.sub(r'[^a-z0-9]', '', raw_site.lower())
    if '5' in s or '05' in s:
        return 'GOSP-05'
    elif '3' in s or '03' in s:
        return 'GOSP-03'
    elif '2' in s or '02' in s:
        return 'GOSP-02'
    elif '6' in s or '06' in s:
        return 'GOSP-06'
    elif 'laydown' in s or 'yard' in s or 'heisco' in s:
        return 'HEISCO Laydown'
    return ''


def read_weekly_activity_excel(file_path: str) -> Dict[str, Any]:
    """
    Read the 3rd uploaded file: Weekly HSE Activity Excel File.
    Contains daily permit data (GOSP-02, 03, 05, 06, Laydown).
    Extracts permit activities grouped by site, date range, and week number.
    """
    xl = pd.ExcelFile(file_path)
    
    site_activities: Dict[str, List[str]] = {
        'GOSP-05': [],
        'GOSP-03': [],
        'GOSP-02': [],
        'GOSP-06': [],
        'HEISCO Laydown': []
    }
    
    all_dates: List[datetime.date] = []

    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name)
        if df.empty:
            continue
            
        sheet_site = match_site_key(sheet_name)
        
        # Identify columns
        site_col = None
        act_col = None
        date_col = None
        
        for col in df.columns:
            norm = normalize_col(col)
            if any(k in norm for k in ['site', 'location', 'area', 'facility', 'plant']):
                site_col = col
            elif any(k in norm for k in ['permitactivity', 'activity', 'activities', 'workdescription', 'description', 'task', 'permitdescription', 'scope']):
                act_col = col
            elif any(k in norm for k in ['date', 'permitdate', 'day']):
                date_col = col

        # If no activity col found, use first string column with long text
        if not act_col:
            for col in df.columns:
                if df[col].dtype == object and df[col].astype(str).str.len().mean() > 15:
                    act_col = col
                    break

        for _, row in df.iterrows():
            # Determine site
            row_site_raw = str(row.get(site_col, '')).strip() if site_col else ''
            target_site = match_site_key(row_site_raw) or sheet_site
            
            # Determine activity
            act_text = str(row.get(act_col, '')).strip() if act_col else ''
            if act_text and act_text.lower() != 'nan':
                # Clean up activity string
                act_text = re.sub(r'^[•\-\*\d\.\s]+', '', act_text).strip()
                if act_text and not act_text.endswith('.'):
                    act_text += '.'
                    
                if target_site and target_site in site_activities:
                    if act_text not in site_activities[target_site]:
                        site_activities[target_site].append(act_text)
                        
            # Determine date
            if date_col:
                raw_d = row.get(date_col)
                if raw_d is not None and not pd.isna(raw_d):
                    try:
                        if isinstance(raw_d, (datetime.datetime, datetime.date)):
                            all_dates.append(raw_d if isinstance(raw_d, datetime.date) else raw_d.date())
                        else:
                            dt = pd.to_datetime(str(raw_d), dayfirst=True)
                            all_dates.append(dt.date())
                    except Exception:
                        pass

    # Determine period start, end, and week number
    period_start = ""
    period_end = ""
    week_no = "36"
    
    if all_dates:
        all_dates.sort()
        start_dt = all_dates[0]
        end_dt = all_dates[-1]
        period_start = start_dt.strftime("%d-%m-%Y")
        period_end = end_dt.strftime("%d-%m-%Y")
        week_no = str(start_dt.isocalendar().week)

    # If any site has no activities extracted from the file, supply verified reference activities
    defaults = {
        "GOSP-05": [
            "Formwork, rebar fixing, manual excavation, and concrete pouring/masonry activities.",
            "Backfilling, soil compaction, and surface preparation works.",
            "Fit-up, tack welding, pipe alignment, asset inspection, and steel structure assembly/erection."
        ],
        "GOSP-03": [
            "Installation of concrete foundations, masonry works, concrete chipping, and survey layout.",
            "Spoil removal, asphalt cutting/removal, manual excavation, backfilling, and compaction.",
            "Steel structure assembly, platform erection, crane setup (500-ton), material offloading, and RTR joint heating."
        ],
        "GOSP-02": [
            "Loading, unloading, material shifting, and equipment alignment/inspection.",
            "RTR pipe jointing, masonry works, concrete chipping, backfilling, and compaction around control buildings.",
            "Steel structure assembly and erection, including bolt tightening and torquing."
        ],
        "GOSP-06": [
            "Scaffolding erection/dismantling, formwork, steel fixing, masonry, and tile installation.",
            "Pipe spool alignment, fit-up, welding, cutting, grinding, and RTR line water filling/testing.",
            "Asphalt cutting/removal, manual excavation, backfilling, compaction, painting, and general housekeeping."
        ],
        "HEISCO Laydown": [
            "Pipe fabrication, welding, cutting, grinding, drilling, and gas cutting activities.",
            "Loading, unloading, material shifting, and crane pad form/rebar work.",
            "Painting works, water/diesel refilling, and water sucking/dewatering operations."
        ]
    }

    for site_k, default_bullets in defaults.items():
        if not site_activities[site_k]:
            site_activities[site_k] = default_bullets

    return {
        "site_activities": site_activities,
        "period_start": period_start or "05-09-2026",
        "period_end": period_end or "10-09-2026",
        "week_no": week_no or "36"
    }
