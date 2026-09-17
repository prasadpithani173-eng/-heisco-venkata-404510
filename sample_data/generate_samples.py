import os
import pandas as pd

def generate_sample_files():
    sample_dir = os.path.join(os.path.dirname(__file__))
    os.makedirs(sample_dir, exist_ok=True)
    
    # 1. Names + Locations Excel matching sample PDF exactly
    names_data = [
        {
            "LOCATION": "GOSP - 03",
            "SAFETY OFFICER": "BILAL SAEED MUHAMM",
            "SUPERVISOR": "EMAD KAMAL",
            "ENGINEER": "AAFAQ AHMAD"
        },
        {
            "LOCATION": "GOSP-3",
            "SAFETY OFFICER": "BILAL SAEED MUHAMM",
            "SUPERVISOR": "EMAD KAMAL",
            "ENGINEER": "AAFAQ AHMAD"
        },
        {
            "LOCATION": "GOSP - 02",
            "SAFETY OFFICER": "RAKIB AJAMIN",
            "SUPERVISOR": "AHMED GADELKARIM",
            "ENGINEER": "OSAMA ABDELRAHMAN"
        },
        {
            "LOCATION": "GOSP-2",
            "SAFETY OFFICER": "RAKIB AJAMIN",
            "SUPERVISOR": "AHMED GADELKARIM",
            "ENGINEER": "OSAMA ABDELRAHMAN"
        },
        {
            "LOCATION": "GOSP - 05",
            "SAFETY OFFICER": "MOHAMED AARIF",
            "SUPERVISOR": "GAMAL SHAWKY",
            "ENGINEER": "AL MOATASEMBELLAH"
        },
        {
            "LOCATION": "GOSP-5",
            "SAFETY OFFICER": "MOHAMED AARIF",
            "SUPERVISOR": "GAMAL SHAWKY",
            "ENGINEER": "AL MOATASEMBELLAH"
        },
        {
            "LOCATION": "GOSP - 06",
            "SAFETY OFFICER": "SHEMEER SHAMSUDE",
            "SUPERVISOR": "AYMAN SALAHELDIN",
            "ENGINEER": "ALAA ABDELGHANY"
        },
        {
            "LOCATION": "GOSP-6",
            "SAFETY OFFICER": "SHEMEER SHAMSUDE",
            "SUPERVISOR": "AYMAN SALAHELDIN",
            "ENGINEER": "ALAA ABDELGHANY"
        },
        {
            "LOCATION": "LAYDOWN",
            "SAFETY OFFICER": "AAMIR SADDIQUE",
            "SUPERVISOR": "RASHEED MAKHMOOR",
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

    # 3. Weekly HSE Activity Daily Permit Data (GOSP-02, 03, 05, 06, Laydown)
    activity_data = [
        # GOSP - 05
        {
            "Date": "05-09-2026",
            "Site": "GOSP - 05",
            "Permit No": "PTW-G05-4401",
            "Permit Type": "Cold Work",
            "Permit Activity": "Formwork, rebar fixing, manual excavation, and concrete pouring/masonry activities."
        },
        {
            "Date": "07-09-2026",
            "Site": "GOSP - 05",
            "Permit No": "PTW-G05-4402",
            "Permit Type": "Cold Work",
            "Permit Activity": "Backfilling, soil compaction, and surface preparation works."
        },
        {
            "Date": "09-09-2026",
            "Site": "GOSP - 05",
            "Permit No": "PTW-G05-4403",
            "Permit Type": "Hot Work",
            "Permit Activity": "Fit-up, tack welding, pipe alignment, asset inspection, and steel structure assembly/erection."
        },
        # GOSP - 03
        {
            "Date": "05-09-2026",
            "Site": "GOSP - 03",
            "Permit No": "PTW-G03-3310",
            "Permit Type": "Cold Work",
            "Permit Activity": "Installation of concrete foundations, masonry works, concrete chipping, and survey layout."
        },
        {
            "Date": "07-09-2026",
            "Site": "GOSP - 03",
            "Permit No": "PTW-G03-3311",
            "Permit Type": "Cold Work",
            "Permit Activity": "Spoil removal, asphalt cutting/removal, manual excavation, backfilling, and compaction."
        },
        {
            "Date": "08-09-2026",
            "Site": "GOSP - 03",
            "Permit No": "PTW-G03-3312",
            "Permit Type": "Critical Lift / Hot Work",
            "Permit Activity": "Steel structure assembly, platform erection, crane setup (500-ton), material offloading, and RTR joint heating."
        },
        # GOSP - 02
        {
            "Date": "06-09-2026",
            "Site": "GOSP - 02",
            "Permit No": "PTW-G02-2201",
            "Permit Type": "Equipment Mobilization",
            "Permit Activity": "Loading, unloading, material shifting, and equipment alignment/inspection."
        },
        {
            "Date": "08-09-2026",
            "Site": "GOSP - 02",
            "Permit No": "PTW-G02-2202",
            "Permit Type": "Cold Work",
            "Permit Activity": "RTR pipe jointing, masonry works, concrete chipping, backfilling, and compaction around control buildings."
        },
        {
            "Date": "10-09-2026",
            "Site": "GOSP - 02",
            "Permit No": "PTW-G02-2203",
            "Permit Type": "Hot Work",
            "Permit Activity": "Steel structure assembly and erection, including bolt tightening and torquing."
        },
        # GOSP - 06
        {
            "Date": "05-09-2026",
            "Site": "GOSP - 06",
            "Permit No": "PTW-G06-6601",
            "Permit Type": "Scaffolding / Civil",
            "Permit Activity": "Scaffolding erection/dismantling, formwork, steel fixing, masonry, and tile installation."
        },
        {
            "Date": "07-09-2026",
            "Site": "GOSP - 06",
            "Permit No": "PTW-G06-6602",
            "Permit Type": "Hot Work / Hydrotest",
            "Permit Activity": "Pipe spool alignment, fit-up, welding, cutting, grinding, and RTR line water filling/testing."
        },
        {
            "Date": "09-09-2026",
            "Site": "GOSP - 06",
            "Permit No": "PTW-G06-6603",
            "Permit Type": "Civil & Housekeeping",
            "Permit Activity": "Asphalt cutting/removal, manual excavation, backfilling, compaction, painting, and general housekeeping."
        },
        # LAYDOWN
        {
            "Date": "06-09-2026",
            "Site": "LAYDOWN",
            "Permit No": "PTW-LAY-5501",
            "Permit Type": "Hot Work",
            "Permit Activity": "Pipe fabrication, welding, cutting, grinding, drilling, and gas cutting activities."
        },
        {
            "Date": "08-09-2026",
            "Site": "LAYDOWN",
            "Permit No": "PTW-LAY-5502",
            "Permit Type": "Civil",
            "Permit Activity": "Loading, unloading, material shifting, and crane pad form/rebar work."
        },
        {
            "Date": "10-09-2026",
            "Site": "LAYDOWN",
            "Permit No": "PTW-LAY-5503",
            "Permit Type": "General Operations",
            "Permit Activity": "Painting works, water/diesel refilling, and water sucking/dewatering operations."
        }
    ]
    df_act = pd.DataFrame(activity_data)
    act_file_path = os.path.join(sample_dir, "Weekly_HSE_Activity_Sample.xlsx")
    df_act.to_excel(act_file_path, index=False)
    print(f"Generated {act_file_path}")

if __name__ == "__main__":
    generate_sample_files()
