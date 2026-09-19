import re
import datetime
from typing import List, Dict, Any, Optional

def normalize_location_key(text: Any) -> str:
    """Normalize location strings for consistent matching."""
    if not text:
        return ""
    s = str(text).upper().strip()
    s = re.sub(r'[\s\-_\./\(\)]+', '', s)
    return s

def parse_sortable_date(val: Any) -> datetime.date:
    """Parse date into sortable date object."""
    if not val:
        return datetime.date.max
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val if isinstance(val, datetime.date) else val.date()
    val_str = str(val).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y", "%d.%m.%Y"):
        try:
            return datetime.datetime.strptime(val_str[:10], fmt).date()
        except Exception:
            pass
    return datetime.date.max

def find_matched_names(location: str, mapping: Optional[Dict[str, Dict[str, Any]]]) -> Dict[str, str]:
    """Find matched officer and engineer names for a given location."""
    if not mapping or not location:
        return {'safety_officer': '', 'supervisor': '', 'engineer': ''}
    key = normalize_location_key(location)
    if key in mapping:
        return mapping[key]
    for map_key, data in mapping.items():
        if map_key and (map_key in key or key in map_key):
            return data
    return {'safety_officer': '', 'supervisor': '', 'engineer': ''}

def process_and_map_observations(raw_records: List[Dict[str, Any]], names_map: Optional[Any] = None) -> List[Dict[str, Any]]:
    """
    DATA INGESTION PROTOCOL:
    Maps incoming spreadsheet records dynamically using the fallback hierarchy:
    - 'Area' or 'Location' -> 'location'
    - 'Finding' or 'Observation' -> 'finding'
    - 'Corrective Action' or 'Recommendation' -> 'corrective_action'
    - 'Assignee' or 'Action By' -> 'assignee'
    - 'Due Date' -> 'due_date'
    - 'Status' -> 'status'
    - 'Risk Level' -> 'risk_level' ('Low' to 'Minor', 'Medium' and 'High' to 'Major')
    """
    processed = []
    
    # Optional date sort if sortable dates exist
    try:
        sorted_records = sorted(
            raw_records,
            key=lambda x: parse_sortable_date(
                x.get('Due Date') or x.get('DUE DATE') or x.get('due_date') or x.get('Observation Date') or x.get('DATE', '')
            )
        )
    except Exception:
        sorted_records = raw_records

    for row in sorted_records:
        # 1. Location fallback hierarchy
        loc = ""
        for k in ['Area', 'Area ', 'Location', 'LOCATION', 'area', 'location']:
            if k in row and row[k] and str(row[k]).strip().lower() != 'nan':
                loc = str(row[k]).strip()
                break

        # 2. Finding / Observation fallback hierarchy
        find = ""
        for k in ['Finding', 'Finding ', 'Observation', 'OBSERVATION', 'finding', 'observation']:
            if k in row and row[k] and str(row[k]).strip().lower() != 'nan':
                find = str(row[k]).strip()
                break

        # 3. Corrective Action / Recommendation fallback hierarchy
        rec = ""
        for k in ['Corrective Action', 'Corrective Action ', 'Recommendation', 'RECOMMENDATION', 'corrective_action', 'recommendation']:
            if k in row and row[k] and str(row[k]).strip().lower() != 'nan':
                rec = str(row[k]).strip()
                break

        # 4. Assignee / Action By fallback hierarchy
        assignee = ""
        for k in ['Assignee', 'Assignee ', 'Action By', 'Responsible Person', 'RESPONSIBLE ENTITY', 'RESPONSIBLE_ENTITY', 'assignee', 'action_by']:
            if k in row and row[k] and str(row[k]).strip().lower() != 'nan':
                assignee = str(row[k]).strip()
                break

        # 5. Due Date fallback hierarchy
        due = ""
        for k in ['Due Date', 'Due Date ', 'DUE DATE', 'due_date', 'due date', 'Observation Date', 'DATE', 'Date']:
            if k in row and row[k] and str(row[k]).strip().lower() != 'nan':
                due = str(row[k]).strip()
                break

        # 6. Status fallback hierarchy
        stat = ""
        for k in ['Status', 'STATUS', 'status', 'State', 'state']:
            if k in row and row[k] and str(row[k]).strip().lower() != 'nan':
                stat = str(row[k]).strip()
                break
        if not stat:
            stat = "Open"
        elif "closed" in stat.lower():
            stat = "Closed"
        else:
            stat = "Open"

        # 7. Risk Level translation: 'Low' to 'Minor', 'Medium' and 'High' to 'Major'
        risk = ""
        for k in ['Risk Level', 'Risk Level ', 'Risk', 'TYPE', 'risk_level', 'risk']:
            if k in row and row[k] and str(row[k]).strip().lower() != 'nan':
                risk = str(row[k]).strip()
                break
        if not risk:
            risk = "Low"

        risk_upper = risk.upper()
        if any(m in risk_upper for m in ["MEDIUM", "HIGH", "MAJOR", "CRITICAL", "SEVERE"]):
            doc_type = "Major"
        else:
            doc_type = "Minor"

        # Raised By / Observed By
        raised = ""
        for k in ['Raised By', 'Raised By ', 'Observed By', 'RAISED_BY', 'observed_by', 'raised_by']:
            if k in row and row[k] and str(row[k]).strip().lower() != 'nan':
                raised = str(row[k]).strip()
                break

        # Observation Date
        obs_date = ""
        for k in ['Observation Date', 'DATE', 'Date', 'observation_date', 'Due Date', 'DUE DATE']:
            if k in row and row[k] and str(row[k]).strip().lower() != 'nan':
                obs_date = str(row[k]).strip()
                break
        if not obs_date or obs_date.lower() == 'nan':
            obs_date = datetime.date.today().strftime("%d/%m/%Y")

        # Skip rows that have neither finding nor location
        if (not find and not loc) or (find == "nan" and loc == "nan"):
            continue

        item = {
            'location': loc if loc else "N/A",
            'finding': find if find else "No entry.",
            'corrective_action': rec if rec else "Monitor.",
            'assignee': assignee if assignee else "Site Team",
            'status_by': assignee if assignee else "Site Team",
            'due_date': due if due else "TBD",
            'status': stat if stat else "Open",
            'risk_level': risk if risk else "Low",
            'doc_type': doc_type,
            'type': doc_type,
            'observed_by': raised if raised else "Safety Officer",
            'reviewed_by': 'AHMED GHALWASH',
            'observation_date': obs_date,

            # Aliases for cross-module compatibility
            'LOCATION': loc if loc else "N/A",
            'OBSERVATION': find if find else "No entry.",
            'RECOMMENDATION': rec if rec else "Monitor.",
            'RESPONSIBLE_ENTITY': assignee if assignee else "Site Team",
            'RESPONSIBLE ENTITY': assignee if assignee else "Site Team",
            'DUE_DATE': due if due else "TBD",
            'DUE DATE': due if due else "TBD",
            'TYPE': doc_type,
            'STATUS': stat if stat else "Open",
            'RAISED_BY': raised if raised else "Safety Officer",
            'ASSIGNEE': assignee if assignee else "Site Team"
        }
        processed.append(item)
        
    return processed

def split_into_pages(
    processed_items: List[Dict[str, Any]], 
    reviewed_by_date: str = "", 
    reviewed_by_name: str = "AHMED GHALWASH"
) -> List[Dict[str, Any]]:
    """
    COMPONENT BATCHING LOGIC:
    - Group records into arrays containing a MAXIMUM of 3 items per card block unit.
    - For each 3-item block container, generate an auto-incrementing sequential tracker string:
      "Ref No: 433001/HSE-OR/[Zero-Padded-Two-Digit-Index]/2026" starting at 01.
    - Dynamically extract 'observed_by' from row data 'Raised By', 'status_by' from row data 'Assignee',
      and hardcode 'reviewed_by' to: 'AHMED GHALWASH'.
    """
    pages = []
    chunk_size = 3
    for i in range(0, len(processed_items), chunk_size):
        chunk = processed_items[i:i + chunk_size]
        page_idx = (i // chunk_size) + 1
        ref_no = f"433001/HSE-OR/{page_idx:02d}/2026"
        
        first_date = chunk[0].get('observation_date', '') if (chunk and len(chunk) > 0) else ''
        if not first_date or first_date == 'nan':
            first_date = chunk[0].get('due_date', '') if (chunk and len(chunk) > 0) else ''
        if not first_date or first_date == 'nan' or first_date == 'TBD':
            first_date = reviewed_by_date if reviewed_by_date else datetime.date.today().strftime("%d/%m/%Y")

        # Dynamically extract observed_by from chunk row data ('Raised By' / 'observed_by')
        observed_by_name = ""
        status_by_name = ""
        for item in chunk:
            if not observed_by_name:
                v = item.get('observed_by') or item.get('Raised By') or item.get('RAISED_BY')
                if v and str(v).strip().lower() not in ['nan', 'none', '', 'safety officer']:
                    observed_by_name = str(v).strip()
            if not status_by_name:
                v = item.get('status_by') or item.get('assignee') or item.get('Assignee') or item.get('ASSIGNEE')
                if v and str(v).strip().lower() not in ['nan', 'none', '', 'site team', 'site engineer', 'site supervisor']:
                    status_by_name = str(v).strip()

        if not observed_by_name:
            # Fallback to any present non-empty observed_by
            for item in chunk:
                v = item.get('observed_by') or item.get('Raised By') or item.get('RAISED_BY')
                if v and str(v).strip().lower() not in ['nan', 'none', '']:
                    observed_by_name = str(v).strip()
                    break
        if not observed_by_name:
            observed_by_name = "Safety Officer"

        if not status_by_name:
            # Fallback to any present non-empty assignee/status_by
            for item in chunk:
                v = item.get('status_by') or item.get('assignee') or item.get('Assignee') or item.get('ASSIGNEE')
                if v and str(v).strip().lower() not in ['nan', 'none', '']:
                    status_by_name = str(v).strip()
                    break
        if not status_by_name:
            status_by_name = "Site Engineer"

        page_block = {
            'ref_no': ref_no,
            'Generated_Ref_No': ref_no,
            'date': first_date,
            'page_date': first_date,
            'Dynamic_First_Record_Date': first_date,
            'cc_name': "DEBOTTLENECK PRODUCTION FACILITIES ABQAIQ BI NO. 10-10303",
            'CC_Name': "DEBOTTLENECK PRODUCTION FACILITIES ABQAIQ BI NO. 10-10303",
            'observations': chunk,
            'records': chunk,
            'observed_by': observed_by_name,
            'Dynamic_Raised_By': observed_by_name,
            'status_by': status_by_name,
            'Dynamic_Assignee': status_by_name,
            'status_by_role': 'Site Engineer',
            'reviewed_by': 'AHMED GHALWASH',
            'Reviewed_By': 'AHMED GHALWASH',
            'reviewed_by_name': 'AHMED GHALWASH',
            'reviewed_by_date': first_date,
            'page_number': page_idx,
            'group_index': page_idx,
            'page_name': f"OR/{page_idx:02d}"
        }
        pages.append(page_block)

    return pages
