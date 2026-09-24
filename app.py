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
st.set_page_config(page_title="DHIS2 Tracker Exporter", page_icon="📊", layout="centered")


# --- Helper & Backend Functions ---
def normalize_base_url(url: str) -> str:
    """Ensures base_url is cleanly formatted without trailing slashes or duplicate /api paths."""
    clean_url = url.strip().rstrip("/")
    if clean_url.endswith("/api"):
        clean_url = clean_url[:-4]
    return clean_url


def build_dhis2_metadata_map(session: requests.Session, base_url: str) -> dict[str, str]:
    """Fetches metadata mappings (UID -> Display Name) from DHIS2."""
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
    """Extracts tracked entity instances from legacy or modern DHIS2 endpoints."""
    base_url = normalize_base_url(web)
    legacy_endpoint = f"{base_url}/api/trackedEntityInstances.json"
    modern_endpoint = f"{base_url}/api/tracker/trackedEntities.json"

    session = requests.Session()
    session.auth = (username, password)
    session.headers.update({"Accept": "application/json", "User-Agent": "DHIS2-Streamlit-Exporter/2.0"})

    retries = Retry(total=5, backoff_factor=2, status_forcelist=[500, 502, 503, 504])
    adapter = HTTPAdapter(max_retries=retries)
    session.mount("https://", adapter)
    session.mount("http://", adapter)

    # Fetch metadata map
    meta_map = build_dhis2_metadata_map(session, base_url)
    all_records = []

    # Detect endpoint version (Legacy vs Tracker API v2)
    active_endpoint = legacy_endpoint
    test_res = session.get(legacy_endpoint, params={"pageSize": 1}, timeout=15)
    if test_res.status_code in (404, 410):
        active_endpoint = modern_endpoint

    for prog, ou in product(idprogram, idou):
        page = 1
        page_size = 100

        while True:
            if active_endpoint == legacy_endpoint:
                params = {
                    "program": prog,
                    "ou": ou,
                    "ouMode": "DESCENDANTS",
                    "pageSize": page_size,
                    "page": page,
                    "totalPages": "true",
                    "fields": "trackedEntityInstance,orgUnit,created,lastUpdated,attributes[attribute,displayName,value],enrollments[enrollment,program,orgUnit,enrolledAt,occurredAt,status,events[event,programStage,occurredAt,status,dataValues[dataElement,value]]]",
                }
            else:
                params = {
                    "program": prog,
                    "orgUnit": ou,
                    "ouMode": "DESCENDANTS",
                    "pageSize": page_size,
                    "page": page,
                    "fields": "*",
                }

            try:
                response = session.get(active_endpoint, params=params, timeout=45)
                if not response.ok:
                    break

                data = response.json()
                instances = data.get("trackedEntityInstances") or data.get("instances") or data.get("trackedEntities") or []
                if not instances:
                    break

                for instance in instances:
                    tei_id = instance.get("trackedEntityInstance") or instance.get("trackedEntity")
                    ou_id = instance.get("orgUnit")
                    ou_name = meta_map.get(ou_id, ou_id)
                    prog_name = meta_map.get(prog, prog)

                    row = {
                        "program_id": prog,
                        "program_name": prog_name,
                        "trackedEntityInstance": tei_id,
                        "orgUnit_id": ou_id,
                        "Org Unit Name": ou_name,
                        "created": instance.get("created") or instance.get("createdAt"),
                        "lastUpdated": instance.get("lastUpdated") or instance.get("updatedAt"),
                    }

                    # Parse Attributes
                    for attr in instance.get("attributes", []):
                        attr_id = attr.get("attribute")
                        col_name = attr.get("displayName") or meta_map.get(attr_id) or attr_id
                        row[col_name] = attr.get("value")

                    # Parse Enrollments & Stage Data Values
                    for enrollment in instance.get("enrollments", []):
                        if enrollment.get("program") == prog:
                            row["enrollment_date"] = enrollment.get("enrolledAt") or enrollment.get("enrollmentDate")
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

                # Pagination logic
                pager = data.get("pager", {})
                page_count = pager.get("pageCount")
                if page_count is not None:
                    if page >= page_count:
                        break
                elif len(instances) < page_size:
                    break

                page += 1
            except Exception:
                break

    session.close()
    df_dhis = pd.DataFrame(all_records)
    df_dhis = df_dhis.reindex(sorted(df_dhis.columns), axis=1)  # Sort columns alphabetically (A -> Z)
    # df_dhis = df_dhis.reindex(sorted(df_dhis.columns, reverse=True), axis=1) # Sort columns in reverse alphabetical order (Z -> A)
    return df_dhis

def convert_df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default active sheet

    # Define Column Selection for YgnTBPro
    SelectedColumnList = [
        'Org Unit Name', 'created', 'lastUpdated', 'Nationality', 'Home Address', 
        'GEN - Name', 'Father Name', 'District', 'GEN - Date of birth', 'Age', 
        'Ward', 'Unique ID (UPI)', 'GEN - Sex', 'Ward / Village tract', 'NRC No.', 
        'GEN - Contact phone number (local)', 'Township (T)', 'Region/State', 
        'enrollment_date', 'enrollment_status', '[TB Screening] Loss of appetite', 
        '[TB Screening] TB CS - Risk factor alcohol', '[TB Screening] CXR result category', 
        '[TB Screening] TB CS - Risk factor undernourishment', '[TB Screening] Cough more than 2 weeks', 
        '[TB Screening] Chest pain', '[TB Screening] CXR screening date', 
        '[TB Screening] TB CS - Risk factor smoking', '[TB Screening] TB CS - Risk factor diabetes', 
        '[TB Screening] CXR screening facility type', '[TB Screening] Referral organization', 
        '[TB Screening] Referral activity'
    ]

    # Style Definitions
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Segoe UI", size=10)
    alt_fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    def _format_and_populate_sheet(ws, sheet_df: pd.DataFrame):
        """Helper function to format headers, cells, zebra-striping, and auto-fit columns."""
        ws.views.sheetView[0].showGridLines = True
        
        # Clean up columns that are completely empty
        clean_df = sheet_df.dropna(how="all", axis=1)
        headers = list(clean_df.columns)
        
        if not headers:
            return

        # Append Header
        ws.append(headers)
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill, cell.font = header_fill, header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # Append Data Rows & Apply Styling
        for row_idx, record in enumerate(clean_df.to_dict(orient="records"), start=2):
            ws.append([record.get(col, "") for col in headers])
            is_even = row_idx % 2 == 0
            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font, cell.border = data_font, thin_border
                cell.alignment = Alignment(vertical="center")
                if is_even:
                    cell.fill = alt_fill

        # Freeze Headers & Set Column Widths
        ws.freeze_panes = "A2"
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 45)

    # 1. ADD CONSOLIDATED SHEET (All Data)
    ws_consolidated = wb.create_sheet(title="Consolidated")
    _format_and_populate_sheet(ws_consolidated, df)

    # 2. ADD YgnTBPro SHEET (Filter Selected Columns that exist in df)
    existing_cols = [col for col in SelectedColumnList if col in df.columns]
    df_ygntbpro = df[existing_cols] if existing_cols else pd.DataFrame()
    ws_ygntbpro = wb.create_sheet(title="YgnTBPro")
    _format_and_populate_sheet(ws_ygntbpro, df_ygntbpro)

    # 3. ADD PROGRAM-SPECIFIC SHEETS
    group_col = "program_name" if "program_name" in df.columns else "program_id"
    if group_col in df.columns:
        for prog_label, prog_df in df.groupby(group_col, dropna=False):
            # Clean sheet title (Max 30 chars, remove illegal characters)
            sheet_title = str(prog_label)[:30] if pd.notna(prog_label) else "Unknown_Program"
            for char in [":", "\\", "/", "?", "*", "[", "]"]:
                sheet_title = sheet_title.replace(char, "_")
            
            ws_prog = wb.create_sheet(title=sheet_title)
            _format_and_populate_sheet(ws_prog, prog_df)

    wb.save(output)
    return output.getvalue()

# def convert_df_to_excel_bytes(df: pd.DataFrame) -> bytes:
#     """Formats DataFrame into a styled Excel file in-memory."""
#     output = io.BytesIO()
#     wb = openpyxl.Workbook()
#     wb.remove(wb.active)

#     header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
#     header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
#     data_font = Font(name="Segoe UI", size=10)
#     alt_fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
#     thin_border = Border(
#         left=Side(style="thin", color="D9D9D9"),
#         right=Side(style="thin", color="D9D9D9"),
#         top=Side(style="thin", color="D9D9D9"),
#         bottom=Side(style="thin", color="D9D9D9"),
#     )

#     group_col = "program_name" if "program_name" in df.columns else "program_id"
#     for prog_label, prog_df in df.groupby(group_col, dropna=False):
#         sheet_title = str(prog_label)[:30] if pd.notna(prog_label) else "Unknown_Program"
#         ws = wb.create_sheet(title=sheet_title)
#         ws.views.sheetView[0].showGridLines = True

#         prog_df_clean = prog_df.dropna(how="all", axis=1)
#         headers = list(prog_df_clean.columns)
#         ws.append(headers)

#         for col_idx in range(1, len(headers) + 1):
#             cell = ws.cell(row=1, column=col_idx)
#             cell.fill, cell.font = header_fill, header_font
#             cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

#         for row_idx, record in enumerate(prog_df_clean.to_dict(orient="records"), start=2):
#             ws.append([record.get(col, "") for col in headers])
#             is_even = row_idx % 2 == 0
#             for col_idx in range(1, len(headers) + 1):
#                 cell = ws.cell(row=row_idx, column=col_idx)
#                 cell.font, cell.border = data_font, thin_border
#                 cell.alignment = Alignment(vertical="center")
#                 if is_even:
#                     cell.fill = alt_fill

#         ws.freeze_panes = "A2"
#         for col in ws.columns:
#             max_len = max(len(str(cell.value or "")) for cell in col)
#             col_letter = get_column_letter(col[0].column)
#             ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 45)

#     wb.save(output)
#     return output.getvalue()


# --- Streamlit UI Components ---
st.title("📊 DHIS2 Tracker Exporter")
st.markdown("Extract tracked entity instances from DHIS2 and download formatted Excel files.")

# Defined Mappings
PROGRAM_MAP = {
    "Registration & Screening": "UZF0HrTlps0",
    "Diagnostic Evaluation": "GvywHD6crky",
    "TB Case Surveillance": "Lt6P15ps7f6",
    "TB Contact Investigation TPT": "cQsXTtAJ3HW",
}

TOWNSHIP_OU_MAP = {
    "HLG (Hlaing)": ["KqjORlUe8Yc"],
    "KMD (Kyeemyindaing)": ["rTTCKrLpxTB"],
    "SDG (Dagon Myothit South)": ["XHz6CPxTAbR"],
    "TGG (Thingangyun)": ["aBfPB9AwbF5"],
    "SOK (South Okkalapa)": ["OeZsFpNKLP5"],
    "MYG (Mayangone)": ["mPwLv1cjror"],
    "SPT (Shwepyithar)": ["aMAEOgli6W8"],
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
        default=list(PROGRAM_MAP.keys()),
    )

    # Township Selection UI
    selected_townships = st.multiselect(
        "Select Townships",
        options=list(TOWNSHIP_OU_MAP.keys()),
        default=list(TOWNSHIP_OU_MAP.keys()),
    )

    submitted = st.form_submit_button("Extract & Process Data")

if submitted:
    # Map selected options to UIDs
    prog_ids = [PROGRAM_MAP[p] for p in selected_programs]
    ou_ids = []
    for township in selected_townships:
        ou_ids.extend(TOWNSHIP_OU_MAP[township])

    if not prog_ids or not ou_ids:
        st.warning("Please select at least one Program and one Township.")
    else:
        with st.spinner("Connecting to DHIS2 and extracting data..."):
            df_result = get_data_dhis2(web_url, username, password, prog_ids, ou_ids)

        if df_result.empty:
            st.error("No data extracted. Check your parameters, credentials, or selected Org Units.")
        else:
            st.success(f"Successfully extracted {len(df_result)} records!")
            excel_data = convert_df_to_excel_bytes(df_result)

            st.download_button(
                label="💾 Download Excel File (.xlsx)",
                data=excel_data,
                file_name="DHIS2_Tracker_Export.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )