import re
import datetime
from typing import Dict, List, Any, Tuple

DEFAULT_CC_NAME = "DEBOTTLENECK PRODUCTION FACILITIES ABQAIQ BI NO. 10-10303"
DEFAULT_REF_PREFIX = "433001/HSE-OR"
DEFAULT_YEAR = "2026"

def normalize_location_key(text: Any) -> str:
    """Normalize location strings like 'ABQ GOSP - 06' and 'GOSP- 06' to match cleanly."""
    if not text:
        return ""
    s = str(text).upper().strip()
    # Remove common prefixes like 'ABQ ', 'ABQ-', 'SAUDI ARAMCO '
    s = re.sub(r'^(ABQ|ARAMCO|SITE|FACILITY)[\s\-_]*', '', s)
    # Remove all non-alphanumeric chars for canonical comparison
    canonical = re.sub(r'[^A-Z0-9]', '', s)
    return canonical

def find_matched_names(location: str, mapping: Dict[str, Dict[str, str]]) -> Dict[str, str]:
    """Find matched Safety Officer, Supervisor, and Engineer for a location."""
    key = normalize_location_key(location)
    
    # 1. Direct match
    if key in mapping:
        return mapping[key]
        
    # 2. Substring match
    for map_key, data in mapping.items():
        if map_key and (map_key in key or key in map_key):
            return data
            
    # 3. Numeric match (e.g. 06, 05, 02, 03)
    loc_numbers = re.findall(r'\d+', key)
    if loc_numbers:
        for map_key, data in mapping.items():
            map_numbers = re.findall(r'\d+', map_key)
            if loc_numbers == map_numbers:
                return data
                
    # Fallback default empty
    return {
        'location_display': location,
        'safety_officer': '',
        'supervisor': '',
        'engineer': '',
        'project_manager': ''
    }

def parse_sortable_date(date_str: str) -> datetime.date:
    """Parse date string into datetime.date for sorting."""
    if not date_str:
        return datetime.date.max
    try:
        # DD/MM/YYYY
        parts = date_str.split('/')
        if len(parts) == 3:
            return datetime.date(int(parts[2]), int(parts[1]), int(parts[0]))
    except Exception:
        pass
    try:
        dt = datetime.datetime.strptime(date_str, "%d/%m/%Y")
        return dt.date()
    except Exception:
        pass
    return datetime.date.max

def process_and_map_observations(
    raw_observations: List[Dict[str, Any]], 
    mapping: Dict[str, Dict[str, str]]
) -> List[Dict[str, Any]]:
    """
    1. Sort observations by Observation Date
    2. Auto Match Names (Safety Officer -> Observed By, Supervisor -> Status By, Engineer -> Reviewed By)
    3. Assign Inspection ID (HSCO-OB-01, HSCO-OB-02, ...)
    """
    # 1. Auto Sort by Observation Date
    sorted_obs = sorted(raw_observations, key=lambda x: parse_sortable_date(x.get('observation_date', '')))
    
    processed = []
    for idx, item in enumerate(sorted_obs, start=1):
        loc = item.get('area', '')
        matched = find_matched_names(loc, mapping)
        
        # Inspection ID formatted as HSCO-OB-01, HSCO-OB-02, ...
        insp_id = f"HSCO-OB-{idx:02d}"
        
        # Type for Word doc: Minor / Major
        raw_type = item.get('observation_type', '').strip()
        type_upper = raw_type.upper()
        if 'MAJOR' in type_upper:
            doc_type = 'Major'
        elif 'MINOR' in type_upper:
            doc_type = 'Minor'
        elif 'ACT' in type_upper:
            doc_type = 'Major'
        else:
            doc_type = 'Minor' if (idx % 2 == 0) else 'Major' # sensible default if not explicitly specified
            
        # JSL observation type: Unsafe Act or Unsafe Condition
        if 'ACT' in type_upper or 'BEHAVIOR' in type_upper or 'PPE' in type_upper:
            jsl_type = 'Unsafe Act'
        elif 'CONDITION' in type_upper:
            jsl_type = 'Unsafe Condition'
        elif doc_type == 'Major':
            jsl_type = 'Unsafe Condition'
        else:
            jsl_type = 'Unsafe Condition'
            
        status = item.get('status', 'Closed')
        if not status:
            status = 'Closed'

        processed_item = {
            'sn': idx,
            'inspection_id': insp_id,
            'area': loc,
            'finding': item.get('finding', ''),
            'corrective_action': item.get('corrective_action', ''),
            'observation_date': item.get('observation_date', ''),
            'responsible_person': item.get('responsible_person', '') or 'Site Supervisor',
            'category': item.get('category', '') or 'General',
            'status': status,
            'date_closed': item.get('date_closed', item.get('observation_date', '')),
            'responsible_company': item.get('responsible_company', 'HEISCO'),
            'assignee_company': item.get('assignee_company', 'HEISCO'),
            'observation_type': raw_type or jsl_type,
            'jsl_observation_type': jsl_type,
            'doc_type': doc_type,
            'remarks': item.get('remarks', ''),
            # Auto matched names
            'observed_by': matched.get('safety_officer') or 'Safety Officer',
            'status_by': matched.get('supervisor') or matched.get('engineer') or 'Site Engineer',
            'reviewed_by': matched.get('engineer') or matched.get('project_manager') or 'Project Manager'
        }
        processed.append(processed_item)
        
    return processed

def split_into_pages(
    processed_observations: List[Dict[str, Any]], 
    reviewed_by_date: str = "",
    reviewed_by_name: str = "AHMED GHALWASH"
) -> List[Dict[str, Any]]:
    """
    Split processed observations into chunks of 3 observations per page.
    Generates:
    - ref_no: 433001/HSE-OR/XX/2026
    - page_name: OR/XX
    - cc_name: DEBOTTLENECK PRODUCTION FACILITIES ABQAIQ BI NO. 10-10303
    - 3 observations per page
    """
    pages = []
    chunk_size = 3
    total = len(processed_observations)
    
    for i in range(0, total, chunk_size):
        chunk = processed_observations[i:i + chunk_size]
        page_num = (i // chunk_size) + 1
        page_ref_num = f"{page_num:02d}"
        
        # Representative date for page header
        page_date = chunk[0].get('observation_date', '') if chunk else ''
        
        # Safety officer and status supervisor for this page's location
        page_safety_officer = chunk[0].get('observed_by', '') if chunk else ''
        page_status_by = chunk[0].get('status_by', '') if chunk else ''
        
        page_data = {
            'page_number': page_num,
            'page_name': f"OR/{page_ref_num}",
            'ref_no': f"{DEFAULT_REF_PREFIX}/{page_ref_num}/{DEFAULT_YEAR}",
            'cc_name': DEFAULT_CC_NAME,
            'page_date': page_date,
            'observations': chunk,
            'observed_by': page_safety_officer,
            'status_by': page_status_by,
            'reviewed_by': reviewed_by_name or "AHMED GHALWASH",
            'reviewed_by_date': reviewed_by_date or page_date
        }
        pages.append(page_data)
        
    return pages
