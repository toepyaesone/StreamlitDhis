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
    page_title="DHIS2 TB Tracker Data Exporter",
    page_icon="📊",
    layout="centered",
)

# --- Constants & Mappings ---
REGISTRATION_PROGRAM_ID = "UZF0HrTlps0"

PROGRAM_MAP = {
    "Registration & Screening": "UZF0HrTlps0",
    "Diagnostic Evaluation": "GvywHD6crky",
    "TB Case Surveillance": "Lt6P15ps7f6",
    "TB Contact Investigation TPT": "cQsXTtAJ3HW",
}

# Includes both YTP/Partner OUs and corresponding Township NTP OUs for different user accounts/roles
TOWNSHIP_OU_MAP = {
    "12. YTP/Pyi Gyi Khin (Hlaing)": ["YGCPqlBamwD", "KqjORlUe8Yc"],
    "12. YTP/Pyi Gyi Khin (Kyeemyindaing)": ["dE0upZMc4xG", "rTTCKrLpxTB"],
    "12. YTP/Pyi Gyi Khin (Dagon Myothit (South))": ["mCgl8tnCaee", "XHz6CPxTAbR"],
    "12. YTP/Pyi Gyi Khin (Thingangyun)": ["lJ9AHubJTjg", "aBfPB9AwbF5"],
    "12. YTP/Pyi Gyi Khin (South Okkalapa)": ["KSRVZsryAKz", "OeZsFpNKLP5"],
    "12. YTP/Pyi Gyi Khin (Mayangone)": ["bctwifUcioE", "mPwLv1cjror"],
    "12. YTP/Pyi Gyi Khin (Shwepyithar)": ["Z7Qf4wgtHz3", "aMAEOgli6W8"],
}

SELECTED_COLUMN_LIST = [
    "Unique ID (UPI)", "GEN - Name", "Age", "GEN - Sex", "Nationality", "Home Address",
    "GEN - Contact phone number (local)", "NRC No.", "GEN - Date of birth",
    "GEN - Date of birth is estimated", "Father Name", "Region/State", "District",
    "Township (T)", "Village", "Ward", "Ward / Village tract",
    "Unique ID (UPI) - Index Case", "Relationship with index", "created", "lastUpdated",
    "enrollment_date", "enrollment_status", "orgUnit_id", "Org Unit Name",
    "program_id", "program_name",
    "[TB Screening] Age (at screening)", "[TB Screening] Any TB drug resistance history?",
    "[TB Screening] BMI", "[TB Screening] Breathlessness", "[TB Screening] CXR result",
    "[TB Screening] CXR result category", "[TB Screening] CXR screening date",
    "[TB Screening] CXR screening done", "[TB Screening] CXR screening facility type",
    "[TB Screening] Chest pain", "[TB Screening] Cough more than 2 weeks",
    "[TB Screening] Current activity", "[TB Screening] Enroll to Diagnostic Evaluation",
    "[TB Screening] Fatigue and Tiredness", "[TB Screening] Fever more than 2 weeks",
    "[TB Screening] Haemoptysis", "[TB Screening] Healthcare worker population",
    "[TB Screening] Height (in inches)", "[TB Screening] Household Contact",
    "[TB Screening] Loss of appetite", "[TB Screening] Migrant population",
    "[TB Screening] Night sweats", "[TB Screening] No symptoms related with TB",
    "[TB Screening] Number of previous TB episodes",
    "[TB Screening] Other referral organisation (Specify)",
    "[TB Screening] Other symptoms related with TB", "[TB Screening] Previous TB History",
    "[TB Screening] Previous TB regimen", "[TB Screening] Referral activity",
    "[TB Screening] Referral organization", "[TB Screening] Specify other symptoms",
    "[TB Screening] TB CS - HIV infection", "[TB Screening] TB CS - HIV status date",
    "[TB Screening] TB CS - Registration - Type of patient in last TB Treatment",
    "[TB Screening] TB CS - Risk factor alcohol", "[TB Screening] TB CS - Risk factor diabetes",
    "[TB Screening] TB CS - Risk factor smoking",
    "[TB Screening] TB CS - Risk factor undernourishment", "[TB Screening] Type of CXR",
    "[TB Screening] Weight (in kgs)", "[TB Screening] Weight loss",
    "[TB Screening] Year of last TB Treatment",
    "[2. TB Treatment] TB CS - Diagnosis date",
    "[2. TB Treatment] TB CS - First-line treatment regimen composition",
    "[2. TB Treatment] TB CS - First-line treatment start date",
    "[2. TB Treatment] TB CS - Manually assigned resistance classification",
    "[2. TB Treatment] TB CS - Outcome due date",
    "[2. TB Treatment] TB CS - Reassign resistance classification",
    "[2. TB Treatment] TB CS - Resistance at diagnosis",
    "[2. TB Treatment] TB CS - Resistance classification",
    "[2. TB Treatment] TB CS - Treatment initiation delay (days)",
    "[2. TB Treatment] TB CS - Treatment regimen",
    "[4. Outcome] TB CS - Treatment outcome",
    "[4. Outcome] TB CS - Treatment outcome delay (weeks)",
]


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
        "organisationUnits": f"{base_url}/api/organisationUnits.json?fields=id,displayName,ancestors[id,displayName]&paging=false",
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


def get_user_accessible_ous(session: requests.Session, base_url: str, fallback_ous: list[str]) -> list[str]:
    """Retrieves the organization units assigned to the current user's account/role."""
    try:
        url = (
            f"{base_url}/api/me.json?"
            "fields=organisationUnits[id],dataViewOrganisationUnits[id],teiSearchOrganisationUnits[id]"
        )
        res = session.get(url, timeout=15)
        if res.ok:
            me_data = res.json()
            user_ous = set()
            for key in ("teiSearchOrganisationUnits", "dataViewOrganisationUnits", "organisationUnits"):
                for item in me_data.get(key, []):
                    if "id" in item:
                        user_ous.add(item["id"])
            if user_ous:
                return list(user_ous)
    except Exception:
        pass
    return fallback_ous


def combine_unique_values(series: pd.Series):
    """Combines non-null unique values for a Series into a pipe-delimited string."""
    non_null = series.dropna().astype(str).str.strip()
    valid_vals = [val for val in non_null if val != "" and val.lower() not in ("nan", "none")]
    if not valid_vals:
        return None
    unique_vals = list(dict.fromkeys(valid_vals))
    if len(unique_vals) == 1:
        return unique_vals[0]
    return " | ".join(unique_vals)


def parse_tei_instances(instances: list[dict], prog: str, meta_map: dict[str, str]) -> list[dict]:
    """Parses raw DHIS2 Tracked Entity Instances into flat dictionary records."""
    records = []
    prog_name = meta_map.get(prog, prog)

    for instance in instances:
        tei_id = instance.get("trackedEntityInstance") or instance.get("trackedEntity")
        tei_ou_id = instance.get("orgUnit")

        # Extract base attributes
        base_attrs = {}
        for attr in instance.get("attributes", []):
            attr_id = attr.get("attribute")
            col_name = attr.get("displayName") or meta_map.get(attr_id) or attr_id
            val = attr.get("value")
            if val is not None and str(val).strip() != "":
                base_attrs[col_name] = val

        # Filter enrollments belonging to the target program
        enrollments = [
            enr for enr in instance.get("enrollments", [])
            if enr.get("program") == prog
        ]

        if not enrollments:
            ou_name = meta_map.get(tei_ou_id, tei_ou_id)
            row = {
                "program_id": prog,
                "program_name": prog_name,
                "trackedEntityInstance": tei_id,
                "orgUnit_id": tei_ou_id,
                "Org Unit Name": ou_name,
                "created": instance.get("created") or instance.get("createdAt"),
                "lastUpdated": instance.get("lastUpdated") or instance.get("updatedAt"),
                **base_attrs,
            }
            records.append(row)
        else:
            for enrollment in enrollments:
                # Prefer the enrollment's orgUnit if present so stage/program OU is accurate
                enr_ou_id = enrollment.get("orgUnit") or tei_ou_id
                enr_ou_name = enrollment.get("orgUnitName") or meta_map.get(enr_ou_id, enr_ou_id)

                row = {
                    "program_id": prog,
                    "program_name": prog_name,
                    "trackedEntityInstance": tei_id,
                    "orgUnit_id": enr_ou_id,
                    "tei_orgUnit_id": tei_ou_id,
                    "Org Unit Name": enr_ou_name,
                    "created": instance.get("created") or instance.get("createdAt"),
                    "lastUpdated": instance.get("lastUpdated") or instance.get("updatedAt"),
                    "enrollment_date": enrollment.get("enrolledAt") or enrollment.get("enrollmentDate"),
                    "enrollment_status": enrollment.get("status"),
                    **base_attrs,
                }

                # Parse Events & Stage Data Values (preserve multiple non-empty values across repeat events)
                for event in enrollment.get("events", []):
                    stage_id = event.get("programStage")
                    stage_name = meta_map.get(stage_id, f"Stage_{stage_id}")

                    for dv in event.get("dataValues", []):
                        val = dv.get("value")
                        if val is None or str(val).strip() == "":
                            continue
                        de_id = dv.get("dataElement")
                        de_name = meta_map.get(de_id, f"Element_{de_id}")
                        col_key = f"[{stage_name}] {de_name}"

                        val_str = str(val).strip()
                        if col_key in row and row[col_key] is not None:
                            existing_vals = [v.strip() for v in str(row[col_key]).split(" | ")]
                            if val_str not in existing_vals:
                                row[col_key] = f"{row[col_key]} | {val_str}"
                        else:
                            row[col_key] = val_str

                records.append(row)

    return records


def fetch_paginated_tei(
    session: requests.Session,
    active_endpoint: str,
    is_legacy: bool,
    prog: str,
    meta_map: dict[str, str],
    ou: str | None = None,
    ou_mode: str = "ACCESSIBLE",
) -> list[dict]:
    """Paginates through all pages for a given program and organization unit configuration."""
    records = []
    page = 1
    page_size = 100

    while True:
        if is_legacy:
            params = {
                "program": prog,
                "ouMode": ou_mode,
                "pageSize": page_size,
                "page": page,
                "totalPages": "true",
                "fields": (
                    "trackedEntityInstance,orgUnit,created,lastUpdated,"
                    "attributes[attribute,displayName,value],"
                    "enrollments[enrollment,program,orgUnit,orgUnitName,enrolledAt,occurredAt,status,"
                    "events[event,programStage,occurredAt,status,dataValues[dataElement,value]]]"
                ),
            }
            if ou:
                params["ou"] = ou
        else:
            params = {
                "program": prog,
                "ouMode": ou_mode,
                "pageSize": page_size,
                "page": page,
                "totalPages": "true",
                "fields": "*",
            }
            if ou:
                params["orgUnit"] = ou

        try:
            response = session.get(active_endpoint, params=params, timeout=45)
            # Handle DHIS2 2.41+ parameter naming differences (`orgUnitMode` vs `ouMode`)
            if not response.ok and not is_legacy and "ouMode" in params:
                params["orgUnitMode"] = params.pop("ouMode")
                if "orgUnit" in params:
                    params["orgUnits"] = params.pop("orgUnit")
                response = session.get(active_endpoint, params=params, timeout=45)

            if not response.ok:
                break

            data = response.json()
            instances = (
                data.get("trackedEntityInstances")
                or data.get("instances")
                or data.get("trackedEntities")
                or []
            )
            if not instances:
                break

            records.extend(parse_tei_instances(instances, prog, meta_map))

            # Check pagination across legacy (`data["pager"]["pageCount"]`) and modern (`data["pageCount"]`)
            pager = data.get("pager", {})
            page_count = pager.get("pageCount") or data.get("pageCount")
            if page_count is not None:
                if page >= int(page_count):
                    break
            elif len(instances) < page_size:
                break

            page += 1
        except Exception:
            break

    return records


def get_data_dhis2(
    web: str,
    username: str,
    password: str,
    idprogram: list[str],
    idou: list[str],
) -> pd.DataFrame:
    """
    Extracts all tracked entity instances across multiple pages and organization units
    from DHIS2 servers to form `df_dhis2`.
    """
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
    all_configured_ous = list({ou for ous in TOWNSHIP_OU_MAP.values() for ou in ous} | set(idou))
    user_ous = get_user_accessible_ous(session, base_url, all_configured_ous)

    active_endpoint = legacy_endpoint
    is_legacy = True
    test_res = session.get(legacy_endpoint, params={"pageSize": 1, "ouMode": "ACCESSIBLE"}, timeout=15)
    if test_res.status_code in (404, 410):
        active_endpoint = modern_endpoint
        is_legacy = False

    all_records = []

    for prog in idprogram:
        # 1. Try fetching all accessible records for the user account/role across all OUs and pages
        prog_records = fetch_paginated_tei(
            session=session,
            active_endpoint=active_endpoint,
            is_legacy=is_legacy,
            prog=prog,
            meta_map=meta_map,
            ou=None,
            ou_mode="ACCESSIBLE",
        )

        # 2. Fallback: If server requires explicit orgUnit IDs, iterate over candidate OUs with DESCENDANTS
        if not prog_records:
            candidate_ous = list(dict.fromkeys(idou + all_configured_ous + user_ous))
            for ou in candidate_ous:
                ou_records = fetch_paginated_tei(
                    session=session,
                    active_endpoint=active_endpoint,
                    is_legacy=is_legacy,
                    prog=prog,
                    meta_map=meta_map,
                    ou=ou,
                    ou_mode="DESCENDANTS",
                )
                prog_records.extend(ou_records)

        all_records.extend(prog_records)

    session.close()

    df_dhis2 = pd.DataFrame(all_records)
    if not df_dhis2.empty:
        # Remove exact duplicate rows if overlapping OU hierarchies returned the same enrollment
        dedup_cols = [c for c in ["program_id", "trackedEntityInstance", "orgUnit_id", "enrollment_date"] if c in df_dhis2.columns]
        if dedup_cols:
            df_dhis2 = df_dhis2.drop_duplicates(subset=dedup_cols, keep="first").reset_index(drop=True)
        df_dhis2 = df_dhis2.reindex(sorted(df_dhis2.columns), axis=1)

    return df_dhis2


def merge_two_program_dfs(left: pd.DataFrame, right: pd.DataFrame, join_key: str) -> pd.DataFrame:
    """Safely outer merges two program DataFrames on a join key without duplicate column errors."""
    left = left.loc[:, ~left.columns.duplicated()].copy()
    right = right.loc[:, ~right.columns.duplicated()].copy()

    merged = pd.merge(left, right, on=join_key, how="outer", suffixes=("_left", "_right"))

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


def process_dhis2_data(
    df_dhis2: pd.DataFrame,
    selected_ou_ids: list[str],
    selected_township_names: list[str] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, dict[str, pd.DataFrame]]:
    """
    Processes `df_dhis2` according to the required workflow:
    1. Filter Registration page by selected organization units and extract unique UPI -> `df_registration`.
    2. Extract all UPI of `df_registration` regardless of organization units for the rest of the pages.
    3. Create `df_combine` and `df_ygntbpro`.
    """
    if df_dhis2.empty:
        empty = pd.DataFrame()
        return empty, empty, empty, {"DHIS2 Data": empty}

    join_key = "Unique ID (UPI)" if "Unique ID (UPI)" in df_dhis2.columns else "trackedEntityInstance"

    # Clean UPI values (strip whitespace and remove empty strings)
    df_dhis2 = df_dhis2.copy()
    if join_key in df_dhis2.columns:
        df_dhis2[join_key] = df_dhis2[join_key].astype(str).str.strip()
        df_dhis2.loc[df_dhis2[join_key].isin(["", "None", "nan", "NaN"]), join_key] = None

    # --- Step 1: Registration Page (`df_registration`) ---
    if "program_id" in df_dhis2.columns:
        reg_mask = df_dhis2["program_id"] == REGISTRATION_PROGRAM_ID
    elif "program_name" in df_dhis2.columns:
        reg_mask = df_dhis2["program_name"].astype(str).str.contains("Registration", case=False, na=False)
    else:
        reg_mask = pd.Series(True, index=df_dhis2.index)

    df_reg_all = df_dhis2[reg_mask].copy()

    # Filter Registration Page by Selected Organization Units
    if selected_ou_ids and not df_reg_all.empty:
        ou_mask = pd.Series(False, index=df_reg_all.index)
        if "orgUnit_id" in df_reg_all.columns:
            ou_mask |= df_reg_all["orgUnit_id"].isin(selected_ou_ids)
        if "tei_orgUnit_id" in df_reg_all.columns:
            ou_mask |= df_reg_all["tei_orgUnit_id"].isin(selected_ou_ids)
        if selected_township_names and "Org Unit Name" in df_reg_all.columns:
            ou_mask |= df_reg_all["Org Unit Name"].isin(selected_township_names)

        df_reg_filtered = df_reg_all[ou_mask].copy()
    else:
        df_reg_filtered = df_reg_all.copy()

    # Extract unique UPI based on organization units in registration page -> `df_registration`
    if join_key in df_reg_filtered.columns and not df_reg_filtered.empty:
        df_reg_valid = df_reg_filtered[df_reg_filtered[join_key].notna()].copy()
        df_registration = (
            df_reg_valid.groupby(join_key, as_index=False, dropna=False)
            .agg(combine_unique_values)
            .reindex(columns=df_dhis2.columns)
        )
        registration_upis = set(df_registration[join_key].dropna().unique())
    else:
        df_registration = df_reg_filtered.copy()
        registration_upis = set()

    # --- Step 2: Rest of the Pages (Filter by `df_registration` UPIs regardless of Org Units) ---
    df_rest_all = df_dhis2[~reg_mask].copy()
    if join_key in df_rest_all.columns and registration_upis:
        df_rest_filtered = df_rest_all[df_rest_all[join_key].isin(registration_upis)].copy()
    else:
        df_rest_filtered = pd.DataFrame(columns=df_dhis2.columns)

    # --- Step 3: Create `df_combine` ---
    df_relevant = pd.concat([df_registration, df_rest_filtered], ignore_index=True)
    if join_key in df_relevant.columns and not df_relevant.empty:
        df_combine = (
            df_relevant.groupby(join_key, as_index=False, dropna=False)
            .agg(combine_unique_values)
            .reindex(columns=df_dhis2.columns)
        )
    else:
        df_combine = df_relevant.copy()

    # --- Step 4: Create `df_ygntbpro` ---
    existing_cols = [col for col in SELECTED_COLUMN_LIST if col in df_combine.columns]
    df_ygntbpro = df_combine[existing_cols].copy() if existing_cols else pd.DataFrame()

    # --- Step 5: Prepare Excel Sheets Dictionary ---
    sheets_data: dict[str, pd.DataFrame] = {
        "DHIS2 Data": df_dhis2,
        "Registration": df_registration,
        "Combined": df_combine,
        "YgnTBPro": df_ygntbpro,
    }

    # Add sheets for the rest of the programs (filtered by `df_registration` UPIs regardless of OU)
    group_col = "program_name" if "program_name" in df_rest_filtered.columns else "program_id"
    if group_col in df_rest_filtered.columns and not df_rest_filtered.empty:
        for prog_label, prog_df in df_rest_filtered.groupby(group_col, dropna=False):
            sheet_title = str(prog_label)[:30] if pd.notna(prog_label) else "Unknown_Program"
            for char in [":", "\\", "/", "?", "*", "[", "]"]:
                sheet_title = sheet_title.replace(char, "_")
            if sheet_title not in sheets_data:
                sheets_data[sheet_title] = prog_df

    return df_registration, df_combine, df_ygntbpro, sheets_data


def convert_df_to_excel_bytes(sheets_dict: dict[str, pd.DataFrame]) -> bytes:
    """Converts a dictionary of named DataFrames into a formatted Excel workbook as bytes."""
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
        bottom=Side(style="thin", color="D9D9D9"),
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

    for sheet_name, sheet_df in sheets_dict.items():
        ws = wb.create_sheet(title=sheet_name)
        _format_and_populate_sheet(ws, sheet_df)

    wb.save(output)
    return output.getvalue()


# --- Streamlit UI Components ---
st.title("📊 DHIS2 TB Tracker Data Exporter")
st.markdown("Extract tracked entity instances from DHIS2 and download formatted Excel files.")

with st.form("dhis2_form"):
    web_url = st.text_input("DHIS2 Base URL", value="https://mbdrtraining.moh.gov.mm/train")
    col1, col2 = st.columns(2)
    with col1:
        username = st.text_input("Username", value="NTP_YTP_Central_S2")
    with col2:
        password = st.text_input("Password", type="password", value="District@1")

    selected_programs = st.multiselect(
        "Select Programs",
        options=list(PROGRAM_MAP.keys()),
        default=list(PROGRAM_MAP.keys()),
    )

    selected_townships = st.multiselect(
        "Select Townships (Applied to Registration Page)",
        options=list(TOWNSHIP_OU_MAP.keys()),
        default=list(TOWNSHIP_OU_MAP.keys()),
    )

    submitted = st.form_submit_button("Extract & Process Data")

if submitted:
    # Ensure Registration & Screening is always included so `df_registration` UPIs can be extracted
    prog_ids = list(dict.fromkeys([REGISTRATION_PROGRAM_ID] + [PROGRAM_MAP[p] for p in selected_programs]))
    selected_ou_ids = []
    for township in selected_townships:
        selected_ou_ids.extend(TOWNSHIP_OU_MAP[township])

    if not selected_ou_ids:
        st.warning("Please select at least one Township.")
    else:
        with st.spinner("Connecting to DHIS2 and extracting all pages (`df_dhis2`)..."):
            df_dhis2 = get_data_dhis2(web_url, username, password, prog_ids, selected_ou_ids)

        if df_dhis2.empty:
            st.error("No data extracted. Check your URL, credentials, or user Org Unit permissions.")
        else:
            with st.spinner("Processing `df_registration`, `df_combine`, and `df_ygntbpro`..."):
                df_registration, df_combine, df_ygntbpro, sheets_dict = process_dhis2_data(
                    df_dhis2=df_dhis2,
                    selected_ou_ids=selected_ou_ids,
                    selected_township_names=selected_townships,
                )
                excel_data = convert_df_to_excel_bytes(sheets_dict)

            st.success(
                f"Extraction complete! "
                f"Total DHIS2 records (`df_dhis2`): {len(df_dhis2)} | "
                f"Unique Registration UPIs (`df_registration`): {len(df_registration)} | "
                f"Combined records (`df_combine`): {len(df_combine)}"
            )

            st.download_button(
                label="💾 Download Excel File (.xlsx)",
                data=excel_data,
                file_name="DHIS2_Tracker_Export.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
 