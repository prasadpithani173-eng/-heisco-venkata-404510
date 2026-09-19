import re
import datetime
from typing import Dict, List, Any, Tuple

DEFAULT_CC_NAME = "DEBOTTLENECK PRODUCTION FACILITIES ABQAIQ BI NO. 10-10303"
DEFAULT_REF_PREFIX = "433001/HSE-OR"
DEFAULT_YEAR = "2026"

DEFAULT_ABQAIQ_MAPPING: Dict[str, Dict[str, str]] = {
    'GOSP03': {'location_display': 'GOSP - 03', 'safety_officer': 'BILAL SAEED MUHAMM', 'supervisor': 'EMAD KAMAL', 'engineer': 'AAFAQ AHMAD', 'project_manager': 'AHMED GHALWASH'},
    'GOSP3': {'location_display': 'GOSP-3', 'safety_officer': 'BILAL SAEED MUHAMM', 'supervisor': 'EMAD KAMAL', 'engineer': 'AAFAQ AHMAD', 'project_manager': 'AHMED GHALWASH'},
    'GOSP02': {'location_display': 'GOSP - 02', 'safety_officer': 'RAKIB AJAMIN', 'supervisor': 'AHMED GADELKARIM', 'engineer': 'OSAMA ABDELRAHMAN', 'project_manager': 'AHMED GHALWASH'},
    'GOSP2': {'location_display': 'GOSP-2', 'safety_officer': 'RAKIB AJAMIN', 'supervisor': 'AHMED GADELKARIM', 'engineer': 'OSAMA ABDELRAHMAN', 'project_manager': 'AHMED GHALWASH'},
    'GOSP05': {'location_display': 'GOSP - 05', 'safety_officer': 'MOHAMED AARIF', 'supervisor': 'GAMAL SHAWKY', 'engineer': 'AL MOATASEMBELLAH', 'project_manager': 'AHMED GHALWASH'},
    'GOSP5': {'location_display': 'GOSP-5', 'safety_officer': 'MOHAMED AARIF', 'supervisor': 'GAMAL SHAWKY', 'engineer': 'AL MOATASEMBELLAH', 'project_manager': 'AHMED GHALWASH'},
    'GOSP06': {'location_display': 'GOSP - 06', 'safety_officer': 'SHEMEER SHAMSUDE', 'supervisor': 'AYMAN SALAHELDIN', 'engineer': 'ALAA ABDELGHANY', 'project_manager': 'AHMED GHALWASH'},
    'GOSP6': {'location_display': 'GOSP-6', 'safety_officer': 'SHEMEER SHAMSUDE', 'supervisor': 'AYMAN SALAHELDIN', 'engineer': 'ALAA ABDELGHANY', 'project_manager': 'AHMED GHALWASH'},
    'LAYDOWN': {'location_display': 'LAYDOWN', 'safety_officer': 'AAMIR SADDIQUE', 'supervisor': 'RASHEED MAKHMOOR', 'engineer': 'IBRAHIM MOHAMED', 'project_manager': 'AHMED GHALWASH'},
}

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
    
    # 1. Direct match in provided mapping
    if mapping and key in mapping:
        return mapping[key]
        
    # 2. Substring match in provided mapping
    if mapping:
        for map_key, data in mapping.items():
            if map_key and (map_key in key or key in map_key):
                return data
                
    # 3. Numeric match in provided mapping (e.g. 06, 05, 02, 03)
    loc_numbers = re.findall(r'\d+', key)
    if mapping and loc_numbers:
        for map_key, data in mapping.items():
            map_numbers = re.findall(r'\d+', map_key)
            if loc_numbers == map_numbers:
                return data

    # 4. Fallback to default project roster for Abqaiq
    if key in DEFAULT_ABQAIQ_MAPPING:
        return DEFAULT_ABQAIQ_MAPPING[key]
    for d_key, d_data in DEFAULT_ABQAIQ_MAPPING.items():
        if d_key in key or key in d_key:
            return d_data
    if loc_numbers:
        for d_key, d_data in DEFAULT_ABQAIQ_MAPPING.items():
            d_numbers = re.findall(r'\d+', d_key)
            if loc_numbers == d_numbers:
                return d_data
                
    # Final default
    return {
        'location_display': location,
        'safety_officer': '',
        'supervisor': '',
        'engineer': '',
        'project_manager': 'AHMED GHALWASH'
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
    1. Sort observations by Due Date / Observation Date
    2. Map Protocol Fields (Location, Finding, Corrective Action, Action By, Due Date, Risk Level/Type, Status)
    3. Auto Match or Extract Signatures:
       - Observed By: Safety Officer (from 'Raised By' or Location Mapping)
       - Status By: Site Engineer / Supervisor (from 'Assignee' or Location Mapping)
       - Reviewed By: Project Manager (globally 'AHMED GHALWASH')
    """
    # 1. Auto Sort by Due Date / Observation Date
    sorted_obs = sorted(
        raw_observations, 
        key=lambda x: parse_sortable_date(x.get('DUE_DATE') or x.get('due_date') or x.get('observation_date', ''))
    )
    
    processed = []
    for idx, item in enumerate(sorted_obs, start=1):
        # 1. Location -> Map to: LOCATION
        loc = item.get('LOCATION') or item.get('location') or item.get('area', '')
        matched = find_matched_names(loc, mapping)
        
        # Inspection ID formatted as HSCO-OR-01, HSCO-OR-02, ...
        insp_id = f"HSCO-OR-{idx:02d}"
        
        # 2. Finding -> Map to: OBSERVATION
        observation_val = item.get('OBSERVATION') or item.get('observation') or item.get('finding', '')
        
        # 3. Corrective Action -> Map to: RECOMMENDATION
        recommendation_val = item.get('RECOMMENDATION') or item.get('recommendation') or item.get('corrective_action', '')
        
        # 4. Action By -> Map to: RESPONSIBLE ENTITY
        resp_entity_val = item.get('RESPONSIBLE_ENTITY') or item.get('responsible_entity') or item.get('responsible_person') or item.get('action_by', '') or 'Site Supervisor'
        if resp_entity_val.lower() in ['nan', 'none']:
            resp_entity_val = 'Site Supervisor'

        # 5. Due Date -> Map to: DUE DATE
        due_date_val = item.get('DUE_DATE') or item.get('due_date') or item.get('observation_date', '')
        if not due_date_val:
            due_date_val = datetime.date.today().strftime("%d/%m/%Y")
            
        # 6. Risk Level -> Map to: TYPE (Process values: 'Low' to 'Minor', 'Medium' and 'High' to 'Major')
        raw_type = item.get('TYPE') or item.get('type') or item.get('observation_type', '') or item.get('risk_level', '')
        type_upper = str(raw_type).strip().upper()
        if 'LOW' in type_upper or 'MINOR' in type_upper:
            doc_type = 'Minor'
        elif 'MED' in type_upper or 'HIGH' in type_upper or 'MAJOR' in type_upper or 'CRITICAL' in type_upper or 'SEVERE' in type_upper:
            doc_type = 'Major'
        else:
            doc_type = 'Minor' if (idx % 2 == 0) else 'Major' # sensible balanced distribution
            
        # JSL observation type: Unsafe Act or Unsafe Condition
        if 'ACT' in type_upper or 'BEHAVIOR' in type_upper or 'PPE' in type_upper:
            jsl_type = 'Unsafe Act'
        else:
            jsl_type = 'Unsafe Condition'
            
        # 7. Status -> Map to: STATUS
        raw_status = str(item.get('STATUS') or item.get('status', 'Closed')).strip()
        status_val = 'Open' if 'open' in raw_status.lower() else 'Closed'

        # 8. Source 'Raised By' -> Map to: Observed By (Safety Officer Name signature)
        so_name = item.get('RAISED_BY') or item.get('raised_by') or item.get('observed_by', '')
        if not so_name or str(so_name).strip().lower() in ['nan', 'none', 'safety officer']:
            so_name = matched.get('safety_officer', '')
        
        # 9. Source 'Assignee' -> Map to: Status By (Site Engineer / Supervisor Name signature)
        status_by_name = item.get('ASSIGNEE') or item.get('assignee') or item.get('status_by', '')
        if not status_by_name or str(status_by_name).strip().lower() in ['nan', 'none', 'site engineer', 'site supervisor']:
            eng_name = matched.get('engineer', '').strip()
            sup_name = matched.get('supervisor', '').strip()
            if eng_name and eng_name.lower() not in ['nan', 'none', 'site engineer']:
                status_by_name = eng_name
            elif sup_name and sup_name.lower() not in ['nan', 'none', 'site supervisor']:
                status_by_name = sup_name
            else:
                status_by_name = ""

        # Clean placeholder titles from names
        for banned in ['Site Supervisor', 'Site Engineer', 'Project Manager', 'Safety Officer', 'nan', 'None']:
            if so_name.strip().lower() == banned.lower():
                so_name = ''
            if status_by_name.strip().lower() == banned.lower():
                status_by_name = ''

        processed_item = {
            # Protocol Uppercase Mappings
            'LOCATION': loc,
            'OBSERVATION': observation_val,
            'RECOMMENDATION': recommendation_val,
            'RESPONSIBLE_ENTITY': resp_entity_val,
            'RESPONSIBLE ENTITY': resp_entity_val,
            'DUE_DATE': due_date_val,
            'DUE DATE': due_date_val,
            'TYPE': doc_type,
            'STATUS': status_val,
            'RAISED_BY': so_name,
            'ASSIGNEE': status_by_name,

            # Standard and System Mappings
            'sn': idx,
            'inspection_id': insp_id,
            'area': loc,
            'location': loc,
            'finding': observation_val,
            'observation': observation_val,
            'corrective_action': recommendation_val,
            'recommendation': recommendation_val,
            'observation_date': due_date_val,
            'due_date': due_date_val,
            'responsible_person': resp_entity_val,
            'responsible_entity': resp_entity_val,
            'category': item.get('category', '') or 'General',
            'status': status_val,
            'date_closed': item.get('date_closed') or due_date_val,
            'responsible_company': item.get('responsible_company', 'HEISCO'),
            'assignee_company': item.get('assignee_company', 'HEISCO'),
            'observation_type': raw_type or jsl_type,
            'jsl_observation_type': jsl_type,
            'doc_type': doc_type,
            'type': doc_type,
            'remarks': item.get('remarks', ''),

            # Signatures
            'observed_by': so_name,
            'raised_by': so_name,
            'Dynamic_Raised_By': so_name,
            'status_by': status_by_name,
            'assignee': status_by_name,
            'Dynamic_Assignee': status_by_name,
            'status_by_role': "Site Engineer",
            'reviewed_by': 'AHMED GHALWASH'
        }
        processed.append(processed_item)
        
    return processed

def split_into_pages(
    processed_observations: List[Dict[str, Any]], 
    reviewed_by_date: str = "",
    reviewed_by_name: str = "AHMED GHALWASH"
) -> List[Dict[str, Any]]:
    """
    COMPONENT GROUPING & PAGINATION LOGIC:
    - Batches records into blocks of up to a MAXIMUM of 3 records per container unit.
    - Generates auto-incrementing tracking index matching:
      433001/HSE-OR/[Zero-Padded-Two-Digit-Index]/2026
    - Sets dynamic footer signature values at the bottom of each 3-record group:
      * Observed By: Safety officer Name: Dynamic_Raised_By
      * Reviewed By: Project Manager Name: AHMED GHALWASH (globally hardcoded)
      * Status By: Site Engineer Name: Dynamic_Assignee
    """
    pages = []
    chunk_size = 3
    total = len(processed_observations)
    
    for i in range(0, total, chunk_size):
        chunk = processed_observations[i:i + chunk_size]
        page_num = (i // chunk_size) + 1
        page_ref_num = f"{page_num:02d}"
        ref_no = f"{DEFAULT_REF_PREFIX}/{page_ref_num}/{DEFAULT_YEAR}"
        
        # Representative date for page header
        first_record_date = (
            chunk[0].get('DUE_DATE') or 
            chunk[0].get('due_date') or 
            chunk[0].get('observation_date', '')
        ) if chunk else ''
        if not first_record_date:
            first_record_date = datetime.date.today().strftime("%d/%m/%Y")
        
        # Dynamic footer values from chunk records
        dynamic_raised_by = (chunk[0].get('RAISED_BY') or chunk[0].get('raised_by') or chunk[0].get('observed_by', '')) if chunk else ''
        dynamic_assignee = (chunk[0].get('ASSIGNEE') or chunk[0].get('assignee') or chunk[0].get('status_by', '')) if chunk else ''
        
        group_data = {
            'group_index': page_num,
            'page_number': page_num,
            'page_name': f"OR/{page_ref_num}",
            'Generated_Ref_No': ref_no,
            'ref_no': ref_no,
            'cc_name': DEFAULT_CC_NAME,
            'CC_Name': DEFAULT_CC_NAME,
            'Dynamic_First_Record_Date': first_record_date,
            'page_date': first_record_date,
            'records': chunk,
            'observations': chunk,
            'Dynamic_Raised_By': dynamic_raised_by,
            'observed_by': dynamic_raised_by,
            'Dynamic_Assignee': dynamic_assignee,
            'status_by': dynamic_assignee,
            'status_by_role': 'Site Engineer',
            'Reviewed_By': "AHMED GHALWASH",
            'reviewed_by': "AHMED GHALWASH",
            'reviewed_by_name': "AHMED GHALWASH",
            'reviewed_by_date': reviewed_by_date or first_record_date
        }
        pages.append(group_data)
        
    return pages

