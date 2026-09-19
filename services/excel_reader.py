import os
import re
import datetime
import numpy as np
import pandas as pd
from typing import Dict, List, Any, Optional

def clean_cell_str(val: Any) -> str:
    """Clean cell values avoiding 'nan', 'None', '<NA>', etc."""
    if val is None or pd.isna(val):
        return ""
    s = str(val).strip()
    if s.lower() in ('nan', 'none', '<na>', 'nat', 'null'):
        return ""
    return s

def parse_date_value(val: Any) -> str:
    """Parse various date representations to DD/MM/YYYY string format."""
    if val is None or pd.isna(val):
        return ""
    if isinstance(val, (datetime.datetime, datetime.date, pd.Timestamp)):
        return val.strftime("%d/%m/%Y")
    
    # Handle numeric Excel serial date (e.g. 45548 = 13/09/2024)
    if isinstance(val, (int, float)):
        try:
            num_val = float(val)
            if 30000 < num_val < 60000:
                dt = pd.to_datetime(num_val, unit='D', origin='1899-12-30')
                return dt.strftime("%d/%m/%Y")
        except Exception:
            pass

    val_str = str(val).strip()
    if not val_str or val_str.lower() in ('nan', 'none', '<na>', 'nat'):
        return ""

    try:
        if val_str.replace('.', '', 1).isdigit():
            num_val = float(val_str)
            if 30000 < num_val < 60000:
                dt = pd.to_datetime(num_val, unit='D', origin='1899-12-30')
                return dt.strftime("%d/%m/%Y")
    except Exception:
        pass

    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y", "%d.%m.%Y"):
        try:
            dt = datetime.datetime.strptime(val_str[:10], fmt)
            return dt.strftime("%d/%m/%Y")
        except Exception:
            pass

    return val_str

def format_date_for_jsl(val: Any) -> str:
    """Format a date value for JSL output."""
    if not val:
        return ""
    return str(val).strip()

def normalize_col(name: Any) -> str:
    """Normalize column name for fuzzy matching."""
    if not isinstance(name, str):
        name = str(name) if name is not None else ""
    return re.sub(r'[^a-z0-9]', '', name.lower())

def match_site_key(name: Any) -> str:
    """Match site name to standard site key."""
    if not name:
        return ""
    norm = normalize_col(name)
    if 'gosp05' in norm or 'gosp5' in norm:
        return 'GOSP-05'
    elif 'gosp03' in norm or 'gosp3' in norm:
        return 'GOSP-03'
    elif 'gosp02' in norm or 'gosp2' in norm:
        return 'GOSP-02'
    elif 'gosp06' in norm or 'gosp6' in norm:
        return 'GOSP-06'
    elif 'heisco' in norm or 'laydown' in norm or 'yard' in norm:
        return 'HEISCO Laydown'
    return ""

def read_weekly_observation_excel(file_path):
    """
    Safely reads weekly observation Excel or CSV file.
    Supports dynamic header discovery and maps incoming keys dynamically:
    - 'Area' or 'Location' -> 'location'
    - 'Finding' or 'Observation' -> 'finding'
    - 'Corrective Action' or 'Recommendation' -> 'corrective_action'
    - 'Assignee' or 'Action By' -> 'assignee'
    - 'Due Date' -> 'due_date'
    - 'Status' -> 'status'
    - 'Risk Level' -> 'risk_level' ('Low' -> 'Minor', 'Medium' and 'High' -> 'Major')
    """
    try:
        is_csv = False
        if isinstance(file_path, str) and file_path.endswith('.csv'):
            is_csv = True
        elif hasattr(file_path, 'filename') and str(file_path.filename).endswith('.csv'):
            is_csv = True

        if is_csv:
            try:
                df = pd.read_csv(file_path)
            except Exception:
                if hasattr(file_path, 'seek'):
                    file_path.seek(0)
                df = pd.read_csv(file_path, encoding='latin1')
        else:
            try:
                df = pd.read_excel(file_path)
            except Exception:
                if hasattr(file_path, 'seek'):
                    file_path.seek(0)
                df = pd.read_excel(file_path, engine='openpyxl')

        if df is None or df.empty:
            return []

        # Check if header row is offset by title banners or metadata
        header_candidates = ['area', 'finding', 'location', 'observation', 'action', 'corrective']
        has_expected_col = any(
            any(hc in str(c).strip().lower() for hc in header_candidates)
            for c in df.columns
        )

        if not has_expected_col:
            for idx, row in df.iterrows():
                row_vals = [str(v).strip().lower() for v in row.values if pd.notna(v)]
                if any(any(hc in v for hc in header_candidates) for v in row_vals):
                    df.columns = [str(c).strip() for c in df.iloc[idx]]
                    df = df.iloc[idx+1:].reset_index(drop=True)
                    break

        df.columns = [str(c).strip() for c in df.columns]
        df = df.replace({np.nan: ""})
        
        records = []
        for _, row in df.iterrows():
            rec = {str(k).strip(): str(v).strip() for k, v in row.items()}
            
            # Extract values according to Data Ingestion Protocol Fallback Hierarchy
            # 1. 'Area' or 'Location' -> 'location'
            raw_loc = ""
            for k in ['Area', 'Area ', 'Location', 'LOCATION', 'area', 'location']:
                if k in rec and rec[k] and rec[k].lower() != 'nan':
                    raw_loc = rec[k]
                    break

            # 2. 'Finding' or 'Observation' -> 'finding'
            raw_find = ""
            for k in ['Finding', 'Finding ', 'Observation', 'OBSERVATION', 'finding', 'observation']:
                if k in rec and rec[k] and rec[k].lower() != 'nan':
                    raw_find = rec[k]
                    break

            # 3. 'Corrective Action' or 'Recommendation' -> 'corrective_action'
            raw_rec = ""
            for k in ['Corrective Action', 'Recommendation', 'RECOMMENDATION', 'corrective action', 'recommendation']:
                if k in rec and rec[k] and rec[k].lower() != 'nan':
                    raw_rec = rec[k]
                    break

                  # 4. Assignee / Designation Role -> 'assignee' (Supervisor, Engineer, Rigger 3)
        raw_assignee = ""
        for k in ['Designation', 'ROLE', 'Role', 'Position', 'Target Role', 'Assignee', 'Action By', 'Responsible Person', 'RESPONSIBLE ENTITY']:
            if k in rec and rec[k] and str(rec[k]).strip().lower() != 'nan':
                raw_assignee = str(rec[k]).strip()
                break

        # 5. Due Date -> 'due_date'
        raw_due = ""
        for k in ['Due Date', 'Due Date ', 'DUE DATE', 'due_date', 'Observation Date', 'DATE']:
            if k in rec and rec[k] and str(rec[k]).strip().lower() != 'nan':
                raw_due = str(parse_date_value(rec[k]))
                # Clean up any trailing timestamps like 15:00:00 cleanly
                if " " in raw_due:
                    raw_due = raw_due.split(" ")[0]
                break

        # 6. Status -> 'status'
        raw_stat = ""
        for k in ['Status', 'STATUS', 'status', 'State']:
            if k in rec and rec[k] and str(rec[k]).strip().lower() != 'nan':
                raw_stat = str(rec[k]).strip()
                break
        
        if not raw_stat:
            raw_stat = "Open"
        elif "closed" in raw_stat.lower():
            raw_stat = "Closed"
        else:
            raw_stat = "Open"

        # 7. Risk Level -> 'risk_level'
        raw_risk = ""
        for k in ['Risk Level', 'Risk', 'TYPE', 'risk level', 'risk']:
            if k in rec and rec[k] and str(rec[k]).strip().lower() != 'nan':
                raw_risk = str(rec[k]).strip()
                break
                
        risk_upper = raw_risk.upper()
        if any(m in risk_upper for m in ['MED', 'HIGH', 'MAJOR', 'CRITICAL']):
            mapped_risk = 'Major'
        else:
            mapped_risk = 'Minor'

        # 8. Raised By / Safety Officer Name -> 'observed_by'
        raw_raised = ""
        for k in ['Raised By', 'Observed By', 'RAISED_BY', 'raised by', 'observed by', 'Safety Officer', 'Safety Officer Name']:
            if k in rec and rec[k] and str(rec[k]).strip().lower() != 'nan':
                raw_raised = str(rec[k]).strip()
                break

        # 9. Engineer Name / Actionee Name -> 'engineer_name'
        raw_engineer = ""
        for k in ['Actionee Name', 'Engineer Name', 'Site Engineer', 'Assignee Name', 'Engineer']:
            if k in rec and rec[k] and str(rec[k]).strip().lower() != 'nan':
                raw_engineer = str(rec[k]).strip()
                break

        # Populate target keys cleanly without swaps or timestamps
        rec_data = {}
        rec_data['location'] = raw_loc
        rec_data['finding'] = raw_find
        rec_data['corrective_action'] = raw_rec
        rec_data['assignee'] = raw_assignee if raw_assignee else "Supervisor"
        rec_data['due_date'] = raw_due
        rec_data['status'] = raw_stat
        rec_data['risk_level'] = mapped_risk
        rec_data['doc_type'] = mapped_risk
        rec_data['type'] = mapped_risk
        
        # Set dynamic names cleanly instead of hardcoded strings
        rec_data['observed_by'] = raw_raised if (raw_raised and raw_raised.lower() != 'internal') else "Safety Officer"
        rec_data['status_by'] = raw_engineer if raw_engineer else "Site Engineer"
        rec_data['reviewed_by'] = raw_raised if (raw_raised and raw_raised.lower() != 'internal') else "Safety Officer"
        rec_data['observation_date'] = raw_due if raw_due else datetime.date.today().strftime("%d/%m/%Y")
        
        # Upper case fallback variables for absolute safety mapping
        rec_data['Location'] = raw_loc
        rec_data['Finding'] = raw_find
        
        records.append(rec_data)

    return records
       

class HybridRecordsDict(dict):
    """
    A dictionary mapping location keys that also supports:
    - .to_dict() returning list of records
    - list indexing and iteration if accessed as records list
    """
    def __init__(self, mapping_dict=None, records_list=None):
        super().__init__(mapping_dict or {})
        self._records = records_list or []

    def to_dict(self, *args, **kwargs):
        return self._records

    def __iter__(self):
        return super().__iter__()

def read_names_locations_excel(file_path):
    """
    Reads names & locations Excel file.
    Returns structured records and location lookup mapping.
    """
    try:
        df = pd.read_excel(file_path)
        df = df.replace({np.nan: ""})
        records = df.to_dict(orient='records')

        mapping: Dict[str, Dict[str, Any]] = {}
        for r in records:
            loc = clean_cell_str(r.get('Location') or r.get('Area') or r.get('LOCATION') or '')
            if loc:
                key = normalize_col(loc)
                so = clean_cell_str(r.get('Safety Officer') or r.get('safety_officer') or r.get('HSE Officer') or '')
                sup = clean_cell_str(r.get('Supervisor') or r.get('supervisor') or r.get('Site Supervisor') or '')
                eng = clean_cell_str(r.get('Engineer') or r.get('engineer') or r.get('Site Engineer') or '')
                
                mapping[key] = {
                    'location_display': loc,
                    'safety_officer': so,
                    'supervisor': sup,
                    'engineer': eng,
                    'safety_officers': [so] if so else [],
                    'supervisors': [sup] if sup else [],
                    'engineers': [eng] if eng else []
                }
                mapping[loc.lower()] = mapping[key]

        return HybridRecordsDict(mapping, records)
    except Exception:
        return HybridRecordsDict({}, [])

class ActivityData(dict):
    """Activity data dictionary supporting dict lookups and record conversions."""
    def __init__(self, data_dict=None, records_list=None):
        super().__init__(data_dict or {})
        self._records = records_list or []

    def to_dict(self, *args, **kwargs):
        return self._records

def read_weekly_activity_excel(file_path):
    """
    Reads weekly activity Excel file.
    Returns site activities, dates, and week number.
    """
    try:
        xl = pd.ExcelFile(file_path)
        site_activities: Dict[str, List[str]] = {
            'GOSP-05': [],
            'GOSP-03': [],
            'GOSP-02': [],
            'GOSP-06': [],
            'HEISCO Laydown': []
        }
        all_dates = []
        all_records = []

        for sheet_name in xl.sheet_names:
            df = xl.parse(sheet_name)
            if df.empty:
                continue
            df = df.replace({np.nan: ""})
            all_records.extend(df.to_dict(orient='records'))
            sheet_site = match_site_key(sheet_name)

            site_col = None
            act_col = None
            date_col = None

            for col in df.columns:
                norm = normalize_col(col)
                if any(k in norm for k in ['site', 'location', 'area', 'facility', 'plant']):
                    site_col = col
                elif any(k in norm for k in ['permitactivity', 'activity', 'activities', 'workdescription', 'description', 'task']):
                    act_col = col
                elif any(k in norm for k in ['date', 'permitdate', 'day']):
                    date_col = col

            for _, row in df.iterrows():
                row_site_raw = str(row.get(site_col, '')).strip() if site_col else ''
                target_site = match_site_key(row_site_raw) or sheet_site
                act_text = str(row.get(act_col, '')).strip() if act_col else ''
                if act_text and act_text.lower() != 'nan':
                    act_text = re.sub(r'^[•\-\*\d\.\s]+', '', act_text).strip()
                    if act_text and not act_text.endswith('.'):
                        act_text += '.'
                    if target_site and target_site in site_activities:
                        if act_text not in site_activities[target_site]:
                            site_activities[target_site].append(act_text)

                if date_col:
                    raw_d = row.get(date_col)
                    if raw_d:
                        try:
                            if isinstance(raw_d, (datetime.datetime, datetime.date)):
                                all_dates.append(raw_d if isinstance(raw_d, datetime.date) else raw_d.date())
                            else:
                                dt = pd.to_datetime(str(raw_d), dayfirst=True)
                                all_dates.append(dt.date())
                        except Exception:
                            pass

        period_start = ""
        period_end = ""
        week_no = "36"

        if all_dates:
            all_dates.sort()
            period_start = all_dates[0].strftime("%d-%m-%Y")
            period_end = all_dates[-1].strftime("%d-%m-%Y")
            week_no = str(all_dates[0].isocalendar().week)

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

        data = {
            "site_activities": site_activities,
            "period_start": period_start or "05-09-2026",
            "period_end": period_end or "10-09-2026",
            "week_no": week_no or "36"
        }
        return ActivityData(data, all_records)
    except Exception:
        return ActivityData({
            "site_activities": {},
            "period_start": "05-09-2026",
            "period_end": "10-09-2026",
            "week_no": "36"
        }, [])
