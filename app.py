from itertools import product
import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import requests
from requests.adapters import HTTPAdapter
import streamlit as st
from urllib3.util import Retry

# --- Streamlit Page Configuration ---
st.set_page_config(page_title="DHIS2 Tracker Exporter", page_icon="📊")

# --- Backend DHIS2 Processing Functions ---
def build_dhis2_metadata_map(session: requests.Session, base_url: str) -> dict[str, str]:
    meta_map = {}
    endpoints = {
        "trackedEntityAttributes": f"{base_url}/api/trackedEntityAttributes.json?fields=id,displayName&paging=false",
        "dataElements": f"{base_url}/api/dataElements.json?fields=id,displayName&paging=false",
        "programStages": f"{base_url}/api/programStages.json?fields=id,displayName&paging=false",
        "programs": f"{base_url}/api/programs.json?fields=id,displayName&paging=false",
        "organisationUnits": f"{base_url}/api/organisationUnits.json?fields=id,displayName&paging=false",
    }
    for resource, url in endpoints.items():
        try:
            res = session.get(url, timeout=30)
            if res.ok:
                items = res.json().get(resource, [])
                for item in items:
                    meta_map[item["id"]] = item.get("displayName", item["id"])
        except Exception:
            pass
    return meta_map

def get_data_dhis2(web: str, username: str, password: str, idprogram: list[str], idou: list[str]) -> pd.DataFrame:
    base_url = web.rsplit("/api/", 1)[0].rstrip("/")
    endpoint = f"{base_url}/api/trackedEntityInstances.json"

    session = requests.Session()
    session.auth = (username, password)
    session.headers.update({"Accept": "application/json", "User-Agent": "DHIS2-Python-Script/1.0"})

    retries = Retry(total=5, backoff_factor=2, status_forcelist=[500, 502, 503, 504])
    adapter = HTTPAdapter(max_retries=retries)
    session.mount("https://", adapter)
    session.mount("http://", adapter)

    meta_map = build_dhis2_metadata_map(session, base_url)
    all_records = []

    for prog, ou in product(idprogram, idou):
        page = 1
        page_size = 100

        while True:
            params = {
                "program": prog,
                "ou": ou,
                "ouMode": "DESCENDANTS",
                "pageSize": page_size,
                "page": page,
                "totalPages": "true",
                "fields": "trackedEntityInstance,orgUnit,created,lastUpdated,attributes[attribute,displayName,value],enrollments[enrollment,program,orgUnit,enrolledAt,occurredAt,status,events[event,programStage,occurredAt,status,dataValues[dataElement,value]]]",
            }
            try:
                response = session.get(endpoint, params=params, timeout=45)
                if not response.ok:
                    break

                data = response.json()
                instances = data.get("trackedEntityInstances", [])
                if not instances:
                    break

                for instance in instances:
                    ou_id = instance.get("orgUnit")
                    ou_name = meta_map.get(ou_id, ou_id)
                    prog_name = meta_map.get(prog, prog)

                    row = {
                        "program_id": prog,
                        "program_name": prog_name,
                        "trackedEntityInstance": instance.get("trackedEntityInstance"),
                        "orgUnit_id": ou_id,
                        "Org Unit Name": ou_name,
                        "created": instance.get("created"),
                        "lastUpdated": instance.get("lastUpdated"),
                    }

                    for attr in instance.get("attributes", []):
                        attr_id = attr.get("attribute")
                        col_name = attr.get("displayName") or meta_map.get(attr_id) or attr_id
                        row[col_name] = attr.get("value")

                    for enrollment in instance.get("enrollments", []):
                        if enrollment.get("program") == prog:
                            row["enrollment_date"] = enrollment.get("enrolledAt")
                            row["enrollment_status"] = enrollment.get("status")

                            for event in enrollment.get("events", []):
                                stage_id = event.get("programStage")
                                stage_name = meta_map.get(stage_id, f"Stage_{stage_id}")

                                for dv in event.get("dataValues", []):
                                    de_id = dv.get("dataElement")
                                    de_name = meta_map.get(de_id, f"Element_{de_id}")
                                    col_key = f"[{stage_name}] {de_name}"
                                    row[col_key] = dv.get("value")

                    all_records.append(row)

                pager = data.get("pager", {})
                if page >= pager.get("pageCount", 1):
                    break
                page += 1
            except Exception:
                break

    session.close()
    return pd.DataFrame(all_records)

def convert_df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Segoe UI", size=10)
    alt_fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    thin_border = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"), top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"))

    group_col = "program_name" if "program_name" in df.columns else "program_id"
    for prog_label, prog_df in df.groupby(group_col, dropna=False):
        sheet_title = str(prog_label)[:30] if pd.notna(prog_label) else "Unknown_Program"
        ws = wb.create_sheet(title=sheet_title)
        ws.views.sheetView[0].showGridLines = True

        prog_df_clean = prog_df.dropna(how="all", axis=1)
        headers = list(prog_df_clean.columns)
        ws.append(headers)

        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill, cell.font = header_fill, header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for row_idx, record in enumerate(prog_df_clean.to_dict(orient="records"), start=2):
            ws.append([record.get(col, "") for col in headers])
            is_even = row_idx % 2 == 0
            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font, cell.border = data_font, thin_border
                cell.alignment = Alignment(vertical="center")
                if is_even: cell.fill = alt_fill

        ws.freeze_panes = "A2"
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 45)

    wb.save(output)
    return output.getvalue()

# PROGRAM ID { Registration&Screening : UZF0HrTlps0 , 
#               DiagnosticEvaluation : GvywHD6crky , 
#               TBCaseSurveillence : Lt6P15ps7f6 ,
#               TBContactInvestigationTPT : cQsXTtAJ3HW }

# ORGANISATION UNIT ID { YTPMATA_HLG : KqjORlUe8Yc , 
#                       YTPMATA_KMD : rTTCKrLpxTB ,
#                       YTPMATA_SDG : XHz6CPxTAbR , 
#                       YTPMMA_TGG : MlBn9fEP74R ,
#                       YTPMATA_TGG : aBfPB9AwbF5 ,
#                       YTPMATA_SOK : OeZsFpNKLP5 ,
#                       YTPMATA_MYG : mPwLv1cjror ,
#                       YTPMATA_SPT : aMAEOgli6W8 }

# # --- Streamlit UI Components ---
# st.title("📊 DHIS2 Tracker Exporter")
# st.markdown("Extract tracked entity instances from DHIS2 and download formatted Excel files.")

# with st.form("dhis2_form"):
#     web_url = st.text_input("DHIS2 Base URL", value="https://hmistraining.mm.dhis2.net/train")
#     col1, col2 = st.columns(2)
#     with col1:
#         username = st.text_input("Username", value="Ygn_NTP1")
#     with col2:
#         password = st.text_input("Password", type="password", value="District@1")



#     prog_input = st.text_area("Program IDs (comma-separated)", value="UZF0HrTlps0, GvywHD6crky, Lt6P15ps7f6, cQsXTtAJ3HW")




#     ou_input = st.text_area("Org Unit IDs (comma-separated)", value="KqjORlUe8Yc, rTTCKrLpxTB, XHz6CPxTAbR, MlBn9fEP74R, aBfPB9AwbF5, OeZsFpNKLP5, mPwLv1cjror, aMAEOgli6W8")
    
#     submitted = st.form_submit_button("Extract & Process Data")

# if submitted:
#     prog_ids = [p.strip() for p in prog_input.split(",") if p.strip()]
#     ou_ids = [o.strip() for o in ou_input.split(",") if o.strip()]

#     with st.spinner("Connecting to DHIS2 and extracting data..."):
#         df_result = get_data_dhis2(web_url, username, password, prog_ids, ou_ids)

#     if df_result.empty:
#         st.error("No data extracted. Check your parameters or credentials.")
#     else:
#         st.success(f"Successfully extracted {len(df_result)} records!")
#         excel_data = convert_df_to_excel_bytes(df_result)
        
#         st.download_button(
#             label="💾 Download Excel File (.xlsx)",
#             data=excel_data,
#             file_name="DHIS2_Tracker_Export.xlsx",
#             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#             type="primary"
#         )


# --- Streamlit UI Components ---
st.title("📊 DHIS2 Tracker Exporter")
st.markdown("Extract tracked entity instances from DHIS2 and download formatted Excel files.")

# Define Mappings
PROGRAM_MAP = {
    "Registration & Screening": "UZF0HrTlps0",
    "Diagnostic Evaluation": "GvywHD6crky",
    "TB Case Surveillance": "Lt6P15ps7f6",
    "TB Contact Investigation TPT": "cQsXTtAJ3HW"
}

TOWNSHIP_OU_MAP = {
    "HLG (Hlaing)": ["KqjORlUe8Yc"],
    "KMD (Kyeemyindaing)": ["rTTCKrLpxTB"],
    "SDG (Dagon Myothit South)": ["XHz6CPxTAbR"],
    "TGG (Thingangyun)": ["aBfPB9AwbF5"],
    "SOK (South Okkalapa)": ["OeZsFpNKLP5"],
    "MYG (Mayangone)": ["mPwLv1cjror"],
    "SPT (Shwepyithar)": ["aMAEOgli6W8"]
}

with st.form("dhis2_form"):
    web_url = st.text_input("DHIS2 Base URL", value="https://hmistraining.mm.dhis2.net/train")
    col1, col2 = st.columns(2)
    with col1:
        username = st.text_input("Username", value="Ygn_NTP1")
    with col2:
        password = st.text_input("Password", type="password", value="District@1")

    # Program Selection UI
    selected_programs = st.multiselect(
        "Select Programs",
        options=list(PROGRAM_MAP.keys()),
        default=list(PROGRAM_MAP.keys())
    )

    # Township Selection UI
    selected_townships = st.multiselect(
        "Select Townships",
        options=list(TOWNSHIP_OU_MAP.keys()),
        default=list(TOWNSHIP_OU_MAP.keys())
    )

    submitted = st.form_submit_button("Extract & Process Data")

if submitted:
    # Extract selected Program IDs
    prog_ids = [PROGRAM_MAP[p] for p in selected_programs]

    # Extract and flatten selected Township Org Unit IDs
    ou_ids = []
    for township in selected_townships:
        ou_ids.extend(TOWNSHIP_OU_MAP[township])

    if not prog_ids or not ou_ids:
        st.warning("Please select at least one Program and one Township.")
    else:
        with st.spinner("Connecting to DHIS2 and extracting data..."):
            df_result = get_data_dhis2(web_url, username, password, prog_ids, ou_ids)

        if df_result.empty:
            st.error("No data extracted. Check your parameters or credentials.")
        else:
            st.success(f"Successfully extracted {len(df_result)} records!")
            excel_data = convert_df_to_excel_bytes(df_result)
            
            st.download_button(
                label="💾 Download Excel File (.xlsx)",
                data=excel_data,
                file_name="DHIS2_Tracker_Export.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )