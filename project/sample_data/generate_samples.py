import os
import pandas as pd

def generate_sample_files():
    sample_dir = os.path.join(os.path.dirname(__file__))
    os.makedirs(sample_dir, exist_ok=True)
    
    # 1. Names + Locations Excel
    names_data = [
        {
            "LOCATION": "GOSP - 06",
            "SAFETY OFFICER": "ASHOK POTHANI",
            "SUPERVISOR": "WAJID HUSSAIN",
            "ENGINEER": "ALAA ABDELGHANY"
        },
        {
            "LOCATION": "GOSP - 05",
            "SAFETY OFFICER": "MOHAMED AARIF",
            "SUPERVISOR": "GAMAL SHAWKY",
            "ENGINEER": "VIJAY GANESAN"
        },
        {
            "LOCATION": "GOSP - 02",
            "SAFETY OFFICER": "RAKIB AJAMIN",
            "SUPERVISOR": "OSAMA ABDELRAHMAN",
            "ENGINEER": "ABDALAZEEM ALF"
        },
        {
            "LOCATION": "GOSP - 03",
            "SAFETY OFFICER": "BILAL SAEED MUH",
            "SUPERVISOR": "MOHAMED ELAMIR",
            "ENGINEER": "AAFAQ AHMAD"
        },
        {
            "LOCATION": "LAYDOWN",
            "SAFETY OFFICER": "AAMIR SADDIQUE",
            "SUPERVISOR": "MOHAMMED ABDULM",
            "ENGINEER": "IBRAHIM MOHAMED"
        }
    ]
    df_names = pd.DataFrame(names_data)
    names_file_path = os.path.join(sample_dir, "Names_Locations_Sample.xlsx")
    df_names.to_excel(names_file_path, index=False)
    print(f"Generated {names_file_path}")

    # 2. Weekly Observations (first 25 real records from HEISCO register)
    observations_data = [
        {
            "S/N": 1,
            "Area": "GOSP- 06",
            "Finding": "Supervisor not using mandatory PPE (Helmet & safety glasses) while supervising grinding work.",
            "Corrective Action": "The supervisor was instructed to always wear the mandatory helmet and safety glasses in the work area.",
            "Observation Date": "29/08/2026",
            "Responsible Person": "WPR/Foreman",
            "Category": "PPE",
            "Status": "Closed",
            "Date Closed": "29/08/2026",
            "Responsible Company": "HEISCO",
            "Assignee Company": "HEISCO",
            "Observation Type": "Major"
        },
        {
            "S/N": 2,
            "Area": "GOSP- 06",
            "Finding": "Hot work (cutting/grinding) area was not barricaded or protected with a fire blanket.",
            "Corrective Action": "Instructed the Fire Watcher to properly cover the activity with a fire blanket and barricade the area.",
            "Observation Date": "29/08/2026",
            "Responsible Person": "Site Supervisor",
            "Category": "Welding & Cutting Operation",
            "Status": "Closed",
            "Date Closed": "29/08/2026",
            "Responsible Company": "HEISCO",
            "Assignee Company": "HEISCO",
            "Observation Type": "Major"
        },
        {
            "S/N": 3,
            "Area": "GOSP- 06",
            "Finding": "Helper was standing on an unsecured small metal piece while a fitter was grinding it, placing his foot in the line of fire.",
            "Corrective Action": "Instructed the Fitter and Helper to secure the workpiece with a proper tool (vise/clamp) and to avoid being in the line of fire.",
            "Observation Date": "29/08/2026",
            "Responsible Person": "Electrician",
            "Category": "Positions of People (Behavioral)",
            "Status": "Closed",
            "Date Closed": "29/08/2026",
            "Responsible Company": "HEISCO",
            "Assignee Company": "HEISCO",
            "Observation Type": "Major"
        },
        {
            "S/N": 4,
            "Area": "GOSP - 05",
            "Finding": "Protruding nails were not removed or bent from dismantled wooden boxes before disposal.",
            "Corrective Action": "Workers were instructed to remove or bend all protruding nails from wood waste before disposal.",
            "Observation Date": "29/08/2026",
            "Responsible Person": "Site Engineer",
            "Category": "Sharp Edge Protection",
            "Status": "Closed",
            "Date Closed": "29/08/2026",
            "Responsible Company": "HEISCO",
            "Assignee Company": "HEISCO",
            "Observation Type": "Minor"
        },
        {
            "S/N": 5,
            "Area": "GOSP - 05",
            "Finding": "A portable water tank was found unattended at the workplace.",
            "Corrective Action": "The unattended water tank was removed from the workplace.",
            "Observation Date": "29/08/2026",
            "Responsible Person": "Mechanical Engineer",
            "Category": "Housekeeping & Material Arrangement",
            "Status": "Closed",
            "Date Closed": "29/08/2026",
            "Responsible Company": "HEISCO",
            "Assignee Company": "HEISCO",
            "Observation Type": "Minor"
        },
        {
            "S/N": 6,
            "Area": "GOSP - 05",
            "Finding": "During chipping activity, the designated Fire Watcher was absent from his post.",
            "Corrective Action": "The Fire Watcher was instructed to remain at the hot work location and continuously monitor the activity.",
            "Observation Date": "29/08/2026",
            "Responsible Person": "Civil Engineer",
            "Category": "Welding & Cutting Operation",
            "Status": "Closed",
            "Date Closed": "29/08/2026",
            "Responsible Company": "HEISCO",
            "Assignee Company": "HEISCO",
            "Observation Type": "Major"
        },
        {
            "S/N": 7,
            "Area": "GOSP - 05",
            "Finding": "Scaffold tag was found with an incorrect inspection date.",
            "Corrective Action": "The scaffold tag was corrected with the proper date as per the two-week standard.",
            "Observation Date": "29/08/2026",
            "Responsible Person": "Electrician",
            "Category": "Scaffolding",
            "Status": "Closed",
            "Date Closed": "29/08/2026",
            "Responsible Company": "HEISCO",
            "Assignee Company": "HEISCO",
            "Observation Type": "Minor"
        },
        {
            "S/N": 8,
            "Area": "GOSP - 05",
            "Finding": "An air blower was being used for dust cleaning without any dust control measures (e.g., water suppression).",
            "Corrective Action": "Effective dust control measures were implemented before work was allowed to resume.",
            "Observation Date": "29/08/2026",
            "Responsible Person": "Civil Supervisor",
            "Category": "Housekeeping & Material Arrangement",
            "Status": "Closed",
            "Date Closed": "29/08/2026",
            "Responsible Company": "HEISCO",
            "Assignee Company": "HEISCO",
            "Observation Type": "Minor"
        },
        {
            "S/N": 9,
            "Area": "GOSP - 05",
            "Finding": "Personnel were off-loading steel materials at night without wearing clear safety glasses.",
            "Corrective Action": "All personnel were instructed to wear clear safety glasses for night work.",
            "Observation Date": "29/08/2026",
            "Responsible Person": "Site Engineer",
            "Category": "PPE",
            "Status": "Closed",
            "Date Closed": "29/08/2026",
            "Responsible Company": "HEISCO",
            "Assignee Company": "HEISCO",
            "Observation Type": "Minor"
        },
        {
            "S/N": 10,
            "Area": "GOSP - 02",
            "Finding": "Compactor operator was working at a high speed despite hot weather conditions, increasing heat stress risk.",
            "Corrective Action": "The operator was coached on the dangers of heat stress and instructed to work at a safe pace according to conditions.",
            "Observation Date": "29/08/2026",
            "Responsible Person": "Mechanical Engineer",
            "Category": "Worksite Medical Facilities and Related Requirements",
            "Status": "Closed",
            "Date Closed": "29/08/2026",
            "Responsible Company": "HEISCO",
            "Assignee Company": "HEISCO",
            "Observation Type": "Major"
        },
        {
            "S/N": 11,
            "Area": "GOSP - 02",
            "Finding": "A compressor was in use without being properly grounded.",
            "Corrective Action": "The compressor was immediately taken out of service until it was properly grounded.",
            "Observation Date": "29/08/2026",
            "Responsible Person": "Civil Engineer",
            "Category": "Electrical",
            "Status": "Closed",
            "Date Closed": "29/08/2026",
            "Responsible Company": "HEISCO",
            "Assignee Company": "HEISCO",
            "Observation Type": "Major"
        },
        {
            "S/N": 12,
            "Area": "GOSP - 02",
            "Finding": "A compressor was found with a missing pre-use inspection checklist.",
            "Corrective Action": "The supervisor was instructed not to use the equipment until a proper pre-use inspection is completed and documented.",
            "Observation Date": "29/08/2026",
            "Responsible Person": "Rigger 3",
            "Category": "Inspection",
            "Status": "Closed",
            "Date Closed": "29/08/2026",
            "Responsible Company": "HEISCO",
            "Assignee Company": "HEISCO",
            "Observation Type": "Minor"
        },
        {
            "S/N": 13,
            "Area": "GOSP - 02",
            "Finding": "SIF Modification activity was ongoing without the required supervision.",
            "Corrective Action": "The job was stopped until the responsible supervisor was present at the worksite.",
            "Observation Date": "29/08/2026",
            "Responsible Person": "Site Supervisor",
            "Category": "Permit To Work",
            "Status": "Closed",
            "Date Closed": "29/08/2026",
            "Responsible Company": "HEISCO",
            "Assignee Company": "HEISCO",
            "Observation Type": "Major"
        },
        {
            "S/N": 14,
            "Area": "GOSP - 02",
            "Finding": "Personnel were working without an umbrella during an 'Orange Flag' (high heat) condition.",
            "Corrective Action": "Work was stopped and umbrellas were provided to the workers as per heat stress plan requirements.",
            "Observation Date": "29/08/2026",
            "Responsible Person": "Electrician",
            "Category": "Worksite Medical Facilities and Related Requirements",
            "Status": "Closed",
            "Date Closed": "29/08/2026",
            "Responsible Company": "HEISCO",
            "Assignee Company": "HEISCO",
            "Observation Type": "Major"
        },
        {
            "S/N": 15,
            "Area": "GOSP - 02",
            "Finding": "Unauthorized personnel were observed standing under a suspended load during a lifting operation.",
            "Corrective Action": "The lifting area was immediately cleared and re-barricaded to prevent unauthorized entry.",
            "Observation Date": "29/08/2026",
            "Responsible Person": "Civil Supervisor",
            "Category": "Positions of People (Behavioral)",
            "Status": "Closed",
            "Date Closed": "29/08/2026",
            "Responsible Company": "HEISCO",
            "Assignee Company": "HEISCO",
            "Observation Type": "Major"
        },
        {
            "S/N": 16,
            "Area": "GOSP - 03",
            "Finding": "Rigging softeners used for sharp edge protection on steel structures were not long enough to protect the web slings from damage.",
            "Corrective Action": "The rigger was instructed to use softeners of the proper length that cover all sharp edges before lifting.",
            "Observation Date": "30/08/2026",
            "Responsible Person": "Site Engineer",
            "Category": "Cranes & Heavy Lifting",
            "Status": "Closed",
            "Date Closed": "30/08/2026",
            "Responsible Company": "HEISCO",
            "Assignee Company": "HEISCO",
            "Observation Type": "Major"
        },
        {
            "S/N": 17,
            "Area": "GOSP - 03",
            "Finding": "An electrical cable was routed on the ground where a steel assembly was in progress, creating a damage risk.",
            "Corrective Action": "The electrical cable was relocated away from the work area to protect it from damage.",
            "Observation Date": "30/08/2026",
            "Responsible Person": "Mechanical Engineer",
            "Category": "Electrical",
            "Status": "Closed",
            "Date Closed": "30/08/2026",
            "Responsible Company": "HEISCO",
            "Assignee Company": "HEISCO",
            "Observation Type": "Minor"
        },
        {
            "S/N": 18,
            "Area": "GOSP - 03",
            "Finding": "Excavation was proceeding with a blunt pickaxe, which is an inefficient and unsafe tool for the task.",
            "Corrective Action": "The blunt hand tool was replaced with a new, sharp pickaxe suitable for the work.",
            "Observation Date": "30/08/2026",
            "Responsible Person": "Civil Engineer",
            "Category": "Hand & Power Tools",
            "Status": "Closed",
            "Date Closed": "30/08/2026",
            "Responsible Company": "HEISCO",
            "Assignee Company": "HEISCO",
            "Observation Type": "Minor"
        },
        {
            "S/N": 19,
            "Area": "GOSP - 03",
            "Finding": "High Heat Stress Index (Category IV) was observed, but work continued.",
            "Corrective Action": "The job was immediately stopped until the heat stress category returned to a safe level.",
            "Observation Date": "30/08/2026",
            "Responsible Person": "WPR/Foreman",
            "Category": "Worksite Medical Facilities and Related Requirements",
            "Status": "Closed",
            "Date Closed": "30/08/2026",
            "Responsible Company": "HEISCO",
            "Assignee Company": "HEISCO",
            "Observation Type": "Major"
        },
        {
            "S/N": 20,
            "Area": "GOSP - 03",
            "Finding": "Backfill material was unloaded less than 0.6 meters from the edge of an excavation.",
            "Corrective Action": "The operator was instructed to always maintain a minimum distance of 0.6 meters from the excavation edge when placing spoil or materials.",
            "Observation Date": "30/08/2026",
            "Responsible Person": "Site Supervisor",
            "Category": "Trenching/Excavating & Shoring",
            "Status": "Closed",
            "Date Closed": "30/08/2026",
            "Responsible Company": "HEISCO",
            "Assignee Company": "HEISCO",
            "Observation Type": "Major"
        },
        {
            "S/N": 21,
            "Area": "GOSP - 03",
            "Finding": "The location mentioned on a cable locating sketch for a cutting permit was incorrect.",
            "Corrective Action": "The permit was rescinded. The excavation point was re-scanned, and a new, correct sketch was made before work was re-authorized.",
            "Observation Date": "30/08/2026",
            "Responsible Person": "Electrician",
            "Category": "Permit To Work",
            "Status": "Closed",
            "Date Closed": "30/08/2026",
            "Responsible Company": "HEISCO",
            "Assignee Company": "HEISCO",
            "Observation Type": "Major"
        },
        {
            "S/N": 22,
            "Area": "GOSP - 03",
            "Finding": "No emergency eyewash station was available in the saltwater work area.",
            "Corrective Action": "A portable, shaded eyewash station was installed in a clearly marked, accessible location.",
            "Observation Date": "30/08/2026",
            "Responsible Person": "Civil Supervisor",
            "Category": "Worksite Medical Facilities and Related Requirements",
            "Status": "Closed",
            "Date Closed": "30/08/2026",
            "Responsible Company": "HEISCO",
            "Assignee Company": "HEISCO",
            "Observation Type": "Major"
        },
        {
            "S/N": 23,
            "Area": "LAYDOWN",
            "Finding": "Scaffolding material was stored at the edge of an excavation, creating a surcharge load.",
            "Corrective Action": "All materials were removed from the edge of the excavation to a designated storage area.",
            "Observation Date": "03/09/2026",
            "Responsible Person": "Scaffolding Supervisor",
            "Category": "Scaffolding",
            "Status": "Closed",
            "Date Closed": "03/09/2026",
            "Responsible Company": "HEISCO",
            "Assignee Company": "HEISCO",
            "Observation Type": "Major"
        },
        {
            "S/N": 24,
            "Area": "LAYDOWN",
            "Finding": "Wooden planks with protruding nails were found in the work area.",
            "Corrective Action": "All wood with protruding nails was collected, and the nails were removed or bent over.",
            "Observation Date": "03/09/2026",
            "Responsible Person": "Rigger 3",
            "Category": "Housekeeping & Material Arrangement",
            "Status": "Closed",
            "Date Closed": "03/09/2026",
            "Responsible Company": "HEISCO",
            "Assignee Company": "HEISCO",
            "Observation Type": "Minor"
        }
    ]
    df_obs = pd.DataFrame(observations_data)
    obs_file_path = os.path.join(sample_dir, "Weekly_Observation_Sample.xlsx")
    df_obs.to_excel(obs_file_path, index=False)
    print(f"Generated {obs_file_path}")

if __name__ == "__main__":
    generate_sample_files()
