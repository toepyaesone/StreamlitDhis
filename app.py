from functools import reduce
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
st.set_page_config(
    page_title="DHIS2 Tracker Exporter", 
    page_icon="📊", 
    layout="centered"
)


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

    meta_map = build_dhis2_metadata_map(session, base_url)
    all_records = []

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
    if not df_dhis.empty:
        df_dhis = df_dhis.reindex(sorted(df_dhis.columns), axis=1)
    return df_dhis


def merge_two_program_dfs(left: pd.DataFrame, right: pd.DataFrame, join_key: str) -> pd.DataFrame:
    """Safely outer merges two program DataFrames on a join key without duplicate column errors."""
    # Deduplicate columns in both dataframes prior to merging
    left = left.loc[:, ~left.columns.duplicated()].copy()
    right = right.loc[:, ~right.columns.duplicated()].copy()

    merged = pd.merge(left, right, on=join_key, how="outer", suffixes=("_left", "_right"))

    # Resolve duplicate column names created during merge
    final_cols = {}
    for col in merged.columns:
        if col == join_key:
            final_cols[col] = merged[col]
        elif col.endswith("_left"):
            base_name = col[:-5]
            right_col = f"{base_name}_right"
            if right_col in merged.columns:
                final_cols[base_name] = merged[col].combine_first(merged[right_col])
            else:
                final_cols[base_name] = merged[col]
        elif col.endswith("_right"):
            base_name = col[:-6]
            if base_name not in final_cols:
                final_cols[base_name] = merged[col]
        else:
            final_cols[col] = merged[col]

    return pd.DataFrame(final_cols)


def prepare_excel_sheets_data(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Processes the original DataFrame and splits it into named DataFrames for each Excel sheet."""
    sheets_data: dict[str, pd.DataFrame] = {}
    SELECTED_COLUMN_LIST = [
        'Unique ID (UPI)', 'GEN - Name', 'Age', 'GEN - Sex','Nationality', 'Home Address','GEN - Contact phone number (local)','NRC No.', 
        'GEN - Date of birth', 'GEN - Date of birth is estimated', 'Father Name', 'Region/State','District', 'Township (T)','Village', 'Ward', 'Ward / Village tract', 
        'Unique ID (UPI) - Index Case', 'Relationship with index','created', 'lastUpdated', 'enrollment_date', 'enrollment_status', 
        'orgUnit_id','Org Unit Name', 'program_id', 'program_name',
        '[TB Screening] Age (at screening)', '[TB Screening] Any TB drug resistance history?', '[TB Screening] BMI', 
        '[TB Screening] Breathlessness', '[TB Screening] CXR result', '[TB Screening] CXR result category', '[TB Screening] CXR screening date', 
        '[TB Screening] CXR screening done', '[TB Screening] CXR screening facility type', '[TB Screening] Chest pain', '[TB Screening] Cough more than 2 weeks', 
        '[TB Screening] Current activity', '[TB Screening] Enroll to Diagnostic Evaluation', '[TB Screening] Fatigue and Tiredness', '[TB Screening] Fever more than 2 weeks', 
        '[TB Screening] Haemoptysis', '[TB Screening] Healthcare worker population', '[TB Screening] Height (in inches)', '[TB Screening] Household Contact', 
        '[TB Screening] Loss of appetite', '[TB Screening] Migrant population', '[TB Screening] Night sweats', '[TB Screening] No symptoms related with TB', 
        '[TB Screening] Number of previous TB episodes', '[TB Screening] Other referral organisation (Specify)', '[TB Screening] Other symptoms related with TB', 
        '[TB Screening] Previous TB History', '[TB Screening] Previous TB regimen', '[TB Screening] Referral activity', '[TB Screening] Referral organization', 
        '[TB Screening] Specify other symptoms', '[TB Screening] TB CS - HIV infection', '[TB Screening] TB CS - HIV status date', '[TB Screening] TB CS - Registration - Type of patient in last TB Treatment', 
        '[TB Screening] TB CS - Risk factor alcohol', '[TB Screening] TB CS - Risk factor diabetes', '[TB Screening] TB CS - Risk factor smoking', 
        '[TB Screening] TB CS - Risk factor undernourishment', '[TB Screening] Type of CXR', '[TB Screening] Weight (in kgs)', '[TB Screening] Weight loss', '[TB Screening] Year of last TB Treatment',
        '[2. TB Treatment] TB CS - Diagnosis date', '[2. TB Treatment] TB CS - First-line treatment regimen composition', '[2. TB Treatment] TB CS - First-line treatment start date', 
        '[2. TB Treatment] TB CS - Manually assigned resistance classification', '[2. TB Treatment] TB CS - Outcome due date', '[2. TB Treatment] TB CS - Reassign resistance classification', 
        '[2. TB Treatment] TB CS - Resistance at diagnosis', '[2. TB Treatment] TB CS - Resistance classification', '[2. TB Treatment] TB CS - Treatment initiation delay (days)', 
        '[2. TB Treatment] TB CS - Treatment regimen', '[4. Outcome] TB CS - Treatment outcome', '[4. Outcome] TB CS - Treatment outcome delay (weeks)'
    ]

    # 1. DHIS2 Data Sheet
    sheets_data["DHIS2 Data"] = df.copy()

    # 2. Combined Sheet
    group_col = "program_name" if "program_name" in df.columns else "program_id"
    join_key = "Unique ID (UPI)" if "Unique ID (UPI)" in df.columns else "trackedEntityInstance"

    if join_key in df.columns and group_col in df.columns and not df.empty:
        program_dfs = [group_df.copy() for _, group_df in df.groupby(group_col, dropna=False)]

        if len(program_dfs) == 1:
            merged_df = program_dfs[0].copy()
        else:
            merged_df = reduce(
                lambda left, right: merge_two_program_dfs(left, right, join_key),
                program_dfs
            )

        existing_selected_cols = [col for col in SELECTED_COLUMN_LIST if col in merged_df.columns]
        sheets_data["Combined"] = merged_df[existing_selected_cols]
    else:
        existing_selected_cols = [col for col in SELECTED_COLUMN_LIST if col in df.columns]
        sheets_data["Combined"] = df[existing_selected_cols] if existing_selected_cols else pd.DataFrame()

    # 3. YgnTBPro Sheet
    existing_cols = [col for col in SELECTED_COLUMN_LIST if col in df.columns]
    sheets_data["YgnTBPro"] = df[existing_cols] if existing_cols else pd.DataFrame()

    # 4. Program-Specific Sheets
    if group_col in df.columns:
        for prog_label, prog_df in df.groupby(group_col, dropna=False):
            sheet_title = str(prog_label)[:30] if pd.notna(prog_label) else "Unknown_Program"
            for char in [":", "\\", "/", "?", "*", "[", "]"]:
                sheet_title = sheet_title.replace(char, "_")
            
            sheets_data[sheet_title] = prog_df

    return sheets_data


def convert_df_to_excel_bytes(data: pd.DataFrame | dict[str, pd.DataFrame]) -> bytes:
    """Converts a DataFrame or pre-prepared sheets dict into a formatted Excel workbook as bytes."""
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default active sheet

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
        ws.views.sheetView[0].showGridLines = True
        clean_df = sheet_df.dropna(how="all", axis=1)
        headers = list(clean_df.columns)
        
        if not headers:
            return

        ws.append(headers)
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill, cell.font = header_fill, header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for row_idx, record in enumerate(clean_df.to_dict(orient="records"), start=2):
            ws.append([record.get(col, "") for col in headers])
            is_even = row_idx % 2 == 0
            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font, cell.border = data_font, thin_border
                cell.alignment = Alignment(vertical="center")
                if is_even:
                    cell.fill = alt_fill

        ws.freeze_panes = "A2"
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 45)

    if isinstance(data, dict):
        sheets_dict = data
    else:
        sheets_dict = prepare_excel_sheets_data(data)

    for sheet_name, sheet_df in sheets_dict.items():
        ws = wb.create_sheet(title=sheet_name)
        _format_and_populate_sheet(ws, sheet_df)

    wb.save(output)
    return output.getvalue() 


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

    selected_programs = st.multiselect(
        "Select Programs",
        options=list(PROGRAM_MAP.keys()),
        default=list(PROGRAM_MAP.keys()),
    )

    selected_townships = st.multiselect(
        "Select Townships",
        options=list(TOWNSHIP_OU_MAP.keys()),
        default=list(TOWNSHIP_OU_MAP.keys()),
    )

    submitted = st.form_submit_button("Extract & Process Data")

if submitted:
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
            
            sheets_dict = prepare_excel_sheets_data(df_result)
            excel_data = convert_df_to_excel_bytes(sheets_dict)

            st.download_button(
                label="💾 Download Excel File (.xlsx)",
                data=excel_data,
                file_name="DHIS2_Tracker_Export.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )