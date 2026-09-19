import pandas as pd
import datetime
from python_calamine import CalamineWorkbook

def parse_date_value(val):
    if pd.isna(val):
        return ""
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.strftime("%d/%m/%Y")
    val_str = str(val).strip()
    if not val_str or val_str.lower() == 'nan':
        return ""
    # Strip any trailing timestamp strings cleanly
    if " " in val_str:
        val_str = val_str.split(" ")[0]
    return val_str

def read_weekly_observation_excel(file_path):
    try:
        # Use python-calamine for fast, low-memory spreadsheet reading
        workbook = CalamineWorkbook.from_path(file_path)
        sheet_names = workbook.sheet_names
        if not sheet_names:
            return []
            
        sheet_name = sheet_names[0]
        rows = workbook.get_sheet_by_name(sheet_name).to_python()
        if not rows:
            return []
            
        headers = [str(h).strip() if h is not None else "" for h in rows[0]]
        raw_records = []
        for r in rows[1:]:
            if not any(r):
                continue
            rec = {}
            for idx, h in enumerate(headers):
                if h:
                    rec[h] = r[idx] if idx < len(r) else None
            raw_records.append(rec)
            
        records = []
        for rec in raw_records:
            # 1. Location / Area
            raw_loc = ""
            for k in ['Area', 'Location', 'AREA', 'LOCATION', 'area', 'location', 'Sub-Area']:
                if k in rec and rec[k] and str(rec[k]).strip().lower() != 'nan':
                    raw_loc = str(rec[k]).strip()
                    break

            # 2. Finding / Observation
            raw_find = ""
            for k in ['Finding', 'Finding ', 'Observation', 'OBSERVATION', 'finding', 'observation']:
                if k in rec and rec[k] and str(rec[k]).strip().lower() != 'nan':
                    raw_find = str(rec[k]).strip()
                    break

            # 3. Corrective Action
            raw_rec = ""
            for k in ['Corrective Action', 'Recommendation', 'RECOMMENDATION', 'corrective action', 'recommendation']:
                if k in rec and rec[k] and str(rec[k]).strip().lower() != 'nan':
                    raw_rec = str(rec[k]).strip()
                    break

            # 4. Assignee Role -> 'assignee' (Supervisor, Engineer, Rigger 3)
            raw_assignee = ""
            for k in ['Designation', 'ROLE', 'Role', 'Position', 'Target Role', 'Assignee', 'Action By', 'Responsible Person', 'RESPONSIBLE ENTITY']:
                if k in rec and rec[k] and str(rec[k]).strip().lower() != 'nan':
                    raw_assignee = str(rec[k]).strip()
                    break

            # 5. Due Date -> 'due_date'
            raw_due = ""
            for k in ['Due Date', 'Due Date ', 'DUE DATE', 'due_date', 'Observation Date', 'DATE']:
                if k in rec and rec[k] and str(rec[k]).strip().lower() != 'nan':
                    raw_due = parse_date_value(rec[k])
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

            # 9. Engineer Name -> 'status_by'
            raw_engineer = ""
            for k in ['Actionee Name', 'Engineer Name', 'Site Engineer', 'Assignee Name', 'Engineer']:
                if k in rec and rec[k] and str(rec[k]).strip().lower() != 'nan':
                    raw_engineer = str(rec[k]).strip()
                    break

            # Package target fields smoothly
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
            
            # Map clean roles and drop name swaps
            rec_data['observed_by'] = raw_raised if (raw_raised and raw_raised.lower() != 'internal') else "Safety Officer"
            rec_data['status_by'] = raw_engineer if raw_engineer else "Site Engineer"
            rec_data['reviewed_by'] = raw_raised if (raw_raised and raw_raised.lower() != 'internal') else "Safety Officer"
            rec_data['observation_date'] = raw_due if raw_due else datetime.date.today().strftime("%d/%m/%Y")
            
            rec_data['Location'] = raw_loc
            rec_data['Finding'] = raw_find
            
            records.append(rec_data)

        return records
    except Exception as e:
        print(f"Error parsing weekly file: {e}")
        return []

class HybridRecordsDict(dict):
    def __init__(self, mapping_dict=None, records_list=None):
        super().__init__(mapping_dict or {})
        self._records = records_list or []

    def to_dict(self, *args, **kwargs):
        return self._records

    def __iter__(self):
        return super().__iter__()

def read_names_locations_excel(file_path):
    try:
        workbook = CalamineWorkbook.from_path(file_path)
        sheet_names = workbook.sheet_names
        if not sheet_names:
            return []
        sheet_name = sheet_names[0]
        rows = workbook.get_sheet_by_name(sheet_name).to_python()
        if not rows:
            return []
        headers = [str(h).strip() if h is not None else "" for h in rows[0]]
        records = []
        for r in rows[1:]:
            if not any(r):
                continue
            rec = {}
            for idx, h in enumerate(headers):
                if h:
                    rec[h] = r[idx] if idx < len(r) else None
            records.append(rec)
        return records
    except Exception as e:
        print(f"Error reading backup file: {e}")
        return []
