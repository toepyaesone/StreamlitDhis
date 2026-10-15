from concurrent.futures import ThreadPoolExecutor, as_completed
import gc
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
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- Constants & Mappings ---
REGISTRATION_PROGRAM_ID = "UZF0HrTlps0"

PROGRAM_MAP = {
    "Registration & Screening": "UZF0HrTlps0",
    "Diagnostic Evaluation": "GvywHD6crky",
    "TB Case Surveillance": "Lt6P15ps7f6",
    "TB Contact Investigation TPT": "cQsXTtAJ3HW",
}

# Includes both Partner/YTP OUs and Township NTP OUs for different user accounts and roles
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


def create_http_session(username: str, password: str) -> requests.Session:
    """Creates a pooled HTTP session with automatic retries for DHIS2 requests."""
    session = requests.Session()
    session.auth = (username, password)
    session.headers.update({"Accept": "application/json", "User-Agent": "DHIS2-Streamlit-Exporter/3.0"})
    retries = Retry(total=3, backoff_factor=1, status_forcelist=[500, 502, 503, 504])
    adapter = HTTPAdapter(pool_connections=20, pool_maxsize=20, max_retries=retries)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_cached_metadata(base_url: str, username: str, password: str) -> tuple[dict[str, str], list[str], bool]:
    """
    Fetches and caches metadata mappings, accessible user OUs, and endpoint type for 1 hour.
    Avoids downloading the entire national organisationUnits tree on every run.
    """
    session = create_http_session(username, password)
    meta_map: dict[str, str] = {}

    # Seed known Township OU names immediately
    for t_name, uids in TOWNSHIP_OU_MAP.items():
        for uid in uids:
            meta_map[uid] = t_name

    endpoints = {
        "trackedEntityAttributes": f"{base_url}/api/trackedEntityAttributes.json?fields=id,displayName&paging=false",
        "dataElements": f"{base_url}/api/dataElements.json?fields=id,displayName&paging=false",
        "programStages": f"{base_url}/api/programStages.json?fields=id,displayName&paging=false",
        "programs": f"{base_url}/api/programs.json?fields=id,displayName&paging=false",
    }

    def _fetch_meta(resource_url: tuple[str, str]):
        resource, url = resource_url
        try:
            res = session.get(url, timeout=25)
            if res.ok:
                return res.json().get(resource, [])
        except Exception:
            pass
        return []

    with ThreadPoolExecutor(max_workers=4) as executor:
        for items in executor.map(_fetch_meta, endpoints.items()):
            for item in items:
                meta_map[item["id"]] = item.get("displayName", item["id"])

    # Fetch user's accessible OUs based on account/role
    user_ous = set()
    try:
        me_url = (
            f"{base_url}/api/me.json?"
            "fields=organisationUnits[id,displayName],"
            "dataViewOrganisationUnits[id,displayName],"
            "teiSearchOrganisationUnits[id,displayName]"
        )
        res = session.get(me_url, timeout=15)
        if res.ok:
            me_data = res.json()
            for key in ("teiSearchOrganisationUnits", "dataViewOrganisationUnits", "organisationUnits"):
                for item in me_data.get(key, []):
                    uid = item.get("id")
                    if uid:
                        user_ous.add(uid)
                        if item.get("displayName"):
                            meta_map[uid] = item["displayName"]
    except Exception:
        pass

    # Detect legacy vs modern tracker endpoint once
    is_legacy = True
    try:
        test_res = session.get(
            f"{base_url}/api/trackedEntityInstances.json",
            params={"pageSize": 1, "ouMode": "ACCESSIBLE"},
            timeout=15,
        )
        if test_res.status_code in (404, 410):
            is_legacy = False
    except Exception:
        pass

    session.close()
    return meta_map, list(user_ous), is_legacy


def parse_tei_instances(instances: list[dict], prog: str, meta_map: dict[str, str]) -> list[dict]:
    """Parses raw DHIS2 Tracked Entity Instances into flat dictionary records."""
    records = []
    prog_name = meta_map.get(prog, prog)

    for instance in instances:
        tei_id = instance.get("trackedEntityInstance") or instance.get("trackedEntity")
        tei_ou_id = instance.get("orgUnit")
        tei_ou_name = instance.get("orgUnitName") or meta_map.get(tei_ou_id, tei_ou_id)

        base_attrs = {}
        for attr in instance.get("attributes", []):
            val = attr.get("value")
            if val is not None and str(val).strip() != "":
                attr_id = attr.get("attribute")
                col_name = attr.get("displayName") or meta_map.get(attr_id) or attr_id
                base_attrs[col_name] = str(val).strip()

        enrollments = [
            enr for enr in instance.get("enrollments", [])
            if enr.get("program") == prog
        ]

        if not enrollments:
            records.append({
                "program_id": prog,
                "program_name": prog_name,
                "trackedEntityInstance": tei_id,
                "orgUnit_id": tei_ou_id,
                "tei_orgUnit_id": tei_ou_id,
                "Org Unit Name": tei_ou_name,
                "created": instance.get("created") or instance.get("createdAt"),
                "lastUpdated": instance.get("lastUpdated") or instance.get("updatedAt"),
                **base_attrs,
            })
        else:
            for enrollment in enrollments:
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
                        existing = row.get(col_key)
                        if existing:
                            if val_str not in existing.split(" | "):
                                row[col_key] = f"{existing} | {val_str}"
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
    """
    Paginates through DHIS2 pages using pageSize=250 and totalPages=false
    to avoid expensive database count queries on every page.
    """
    records = []
    page = 1
    page_size = 250

    while True:
        if is_legacy:
            params = {
                "program": prog,
                "ouMode": ou_mode,
                "pageSize": page_size,
                "page": page,
                "totalPages": "false",
                "fields": (
                    "trackedEntityInstance,orgUnit,orgUnitName,created,lastUpdated,"
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
                "totalPages": "false",
                "fields": "*",
            }
            if ou:
                params["orgUnit"] = ou

        try:
            response = session.get(active_endpoint, params=params, timeout=40)
            if not response.ok and not is_legacy and "ouMode" in params:
                params["orgUnitMode"] = params.pop("ouMode")
                if "orgUnit" in params:
                    params["orgUnits"] = params.pop("orgUnit")
                response = session.get(active_endpoint, params=params, timeout=40)

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

            # If fewer instances than page_size are returned, we reached the final page
            if len(instances) < page_size:
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
    status_container=None,
) -> pd.DataFrame:
    """
    Extracts all tracked entity instances across multiple pages and organization units
    concurrently to build `df_dhis2`.
    """
    base_url = normalize_base_url(web)
    if status_container:
        status_container.write("🔄 Fetching & caching DHIS2 metadata and user OU permissions...")

    meta_map, user_ous, is_legacy = fetch_cached_metadata(base_url, username, password)
    active_endpoint = (
        f"{base_url}/api/trackedEntityInstances.json"
        if is_legacy
        else f"{base_url}/api/tracker/trackedEntities.json"
    )

    all_configured_ous = list({ou for ous in TOWNSHIP_OU_MAP.values() for ou in ous} | set(idou))
    candidate_ous = list(dict.fromkeys(idou + all_configured_ous + user_ous))

    session = create_http_session(username, password)
    all_records = []

    def _fetch_program(prog: str) -> list[dict]:
        # 1. Try single-pass extraction across all accessible OUs for this account/role
        prog_records = fetch_paginated_tei(
            session=session,
            active_endpoint=active_endpoint,
            is_legacy=is_legacy,
            prog=prog,
            meta_map=meta_map,
            ou=None,
            ou_mode="ACCESSIBLE",
        )
        # 2. Fallback: If DHIS2 server requires explicit OU UIDs, fetch candidate OUs in parallel
        if not prog_records and candidate_ous:
            with ThreadPoolExecutor(max_workers=6) as ou_executor:
                futures = [
                    ou_executor.submit(
                        fetch_paginated_tei,
                        session,
                        active_endpoint,
                        is_legacy,
                        prog,
                        meta_map,
                        ou,
                        "DESCENDANTS",
                    )
                    for ou in candidate_ous
                ]
                for fut in as_completed(futures):
                    prog_records.extend(fut.result())
        return prog_records

    # Fetch all programs concurrently
    with ThreadPoolExecutor(max_workers=len(idprogram)) as prog_executor:
        future_to_prog = {prog_executor.submit(_fetch_program, prog): prog for prog in idprogram}
        for future in as_completed(future_to_prog):
            prog_id = future_to_prog[future]
            prog_label = meta_map.get(prog_id, prog_id)
            res_records = future.result()
            all_records.extend(res_records)
            if status_container:
                status_container.write(f"✅ Extracted **{len(res_records):,}** records from *{prog_label}*")

    session.close()

    df_dhis2 = pd.DataFrame(all_records)
    if not df_dhis2.empty:
        # Resolve missing Org Unit Names using non-empty names in the dataset
        if "orgUnit_id" in df_dhis2.columns and "Org Unit Name" in df_dhis2.columns:
            ou_lookup = (
                df_dhis2.dropna(subset=["orgUnit_id", "Org Unit Name"])
                .loc[lambda d: d["orgUnit_id"] != d["Org Unit Name"]]
                .set_index("orgUnit_id")["Org Unit Name"]
                .to_dict()
            )
            if ou_lookup:
                df_dhis2["Org Unit Name"] = df_dhis2["orgUnit_id"].map(ou_lookup).fillna(df_dhis2["Org Unit Name"])

        dedup_cols = [
            c for c in ["program_id", "trackedEntityInstance", "orgUnit_id", "enrollment_date"]
            if c in df_dhis2.columns
        ]
        if dedup_cols:
            df_dhis2 = df_dhis2.drop_duplicates(subset=dedup_cols, keep="first").reset_index(drop=True)
        df_dhis2 = df_dhis2.reindex(sorted(df_dhis2.columns), axis=1)

    return df_dhis2


def fast_combine_by_key(df: pd.DataFrame, join_key: str, column_order: list[str]) -> pd.DataFrame:
    """
    High-speed vectorized replacement for `.groupby(join_key).agg(combine_unique_values)`.
    Avoids running slow Python loops on empty/single-value cells.
    """
    if df.empty or join_key not in df.columns:
        return df.copy()

    valid_df = df[df[join_key].notna()].copy()
    if valid_df.empty:
        return pd.DataFrame(columns=column_order)

    # If each key already appears only once, skip heavy aggregation
    if valid_df[join_key].is_unique:
        return valid_df.reindex(columns=column_order).reset_index(drop=True)

    # Identify columns that actually contain non-null values
    non_empty_cols = [c for c in valid_df.columns if c != join_key and valid_df[c].notna().any()]
    if not non_empty_cols:
        return valid_df[[join_key]].drop_duplicates().reindex(columns=column_order).reset_index(drop=True)

    # Convert to long format, drop nulls/duplicates, and join unique values per (join_key, column)
    melted = (
        valid_df[[join_key] + non_empty_cols]
        .melt(id_vars=join_key, var_name="_col", value_name="_val")
        .dropna(subset=["_val"])
    )
    melted["_val"] = melted["_val"].astype(str).str.strip()
    melted = melted[~melted["_val"].isin(["", "nan", "None", "NaN"])]

    # Split already pipe-separated strings so duplicates across rows are cleanly deduplicated
    if melted["_val"].str.contains(" | ", regex=False).any():
        melted["_val"] = melted["_val"].str.split(r"\s*\|\s*")
        melted = melted.explode("_val")

    melted = melted.drop_duplicates(subset=[join_key, "_col", "_val"], keep="first")

    grouped = (
        melted.groupby([join_key, "_col"], sort=False)["_val"]
        .agg(" | ".join)
        .unstack("_col")
        .reset_index()
    )

    return grouped.reindex(columns=column_order)


def process_dhis2_data(
    df_dhis2: pd.DataFrame,
    selected_ou_ids: list[str],
    selected_township_names: list[str] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, dict[str, pd.DataFrame]]:
    """
    Maintains exact data architecture:
    1. Extract unique UPI based on organization units in registration page -> `df_registration`
       and apply filter for selected organization units.
    2. Extract all UPI of `df_registration` regardless of organization units for the rest of the pages.
    3. Create `df_combine` and `df_ygntbpro`.
    """
    if df_dhis2.empty:
        empty = pd.DataFrame()
        return empty, empty, empty, {"DHIS2 Data": empty}

    join_key = "Unique ID (UPI)" if "Unique ID (UPI)" in df_dhis2.columns else "trackedEntityInstance"

    df_dhis2 = df_dhis2.copy()
    if join_key in df_dhis2.columns:
        df_dhis2[join_key] = df_dhis2[join_key].astype(str).str.strip()
        df_dhis2.loc[df_dhis2[join_key].isin(["", "None", "nan", "NaN"]), join_key] = None

    # --- 1. Registration Page (`df_registration`) ---
    if "program_id" in df_dhis2.columns:
        reg_mask = df_dhis2["program_id"] == REGISTRATION_PROGRAM_ID
    elif "program_name" in df_dhis2.columns:
        reg_mask = df_dhis2["program_name"].astype(str).str.contains("Registration", case=False, na=False)
    else:
        reg_mask = pd.Series(True, index=df_dhis2.index)

    df_reg_all = df_dhis2[reg_mask].copy()

    # Apply filter for selected organization units on registration page
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

    # Extract unique UPI in registration page -> `df_registration`
    df_registration = fast_combine_by_key(df_reg_filtered, join_key, list(df_dhis2.columns))
    registration_upis = (
        set(df_registration[join_key].dropna().unique())
        if join_key in df_registration.columns and not df_registration.empty
        else set()
    )

    # --- 2. Rest of the Pages (All UPI of `df_registration` regardless of Org Units) ---
    df_rest_all = df_dhis2[~reg_mask].copy()
    if join_key in df_rest_all.columns and registration_upis:
        df_rest_filtered = df_rest_all[df_rest_all[join_key].isin(registration_upis)].copy()
    else:
        df_rest_filtered = pd.DataFrame(columns=df_dhis2.columns)

    # --- 3. Create `df_combine` ---
    df_relevant = pd.concat([df_registration, df_rest_filtered], ignore_index=True)
    df_combine = fast_combine_by_key(df_relevant, join_key, list(df_dhis2.columns))

    # --- 4. Create `df_ygntbpro` ---
    existing_cols = [col for col in SELECTED_COLUMN_LIST if col in df_combine.columns]
    df_ygntbpro = df_combine[existing_cols].copy() if existing_cols else pd.DataFrame()

    # --- 5. Assemble Excel Sheets Dictionary ---
    sheets_data: dict[str, pd.DataFrame] = {
        "YgnTBPro": df_ygntbpro,
        "Combined": df_combine,
        "Registration": df_registration,
        "DHIS2 Data": df_dhis2,
    }

    group_col = "program_name" if "program_name" in df_rest_filtered.columns else "program_id"
    if group_col in df_rest_filtered.columns and not df_rest_filtered.empty:
        for prog_label, prog_df in df_rest_filtered.groupby(group_col, dropna=False):
            sheet_title = str(prog_label)[:30] if pd.notna(prog_label) else "Unknown_Program"
            for char in [":", "\\", "/", "?", "*", "[", "]"]:
                sheet_title = sheet_title.replace(char, "_")
            if sheet_title not in sheets_data:
                sheets_data[sheet_title] = prog_df

    return df_registration, df_combine, df_ygntbpro, sheets_data


def convert_df_to_excel_bytes(sheets_dict: dict[str, pd.DataFrame], lightweight_mode: bool = True) -> bytes:
    """
    Memory-optimized Excel builder for Streamlit Community Cloud (1 GB RAM limit).
    - Calculates column widths via vectorized Pandas operations instead of looping over openpyxl cells.
    - Uses `write_only=True` when `lightweight_mode` is enabled or when total cells exceed 150,000.
    """
    output = io.BytesIO()
    total_cells = sum(df.shape[0] * df.shape[1] for df in sheets_dict.values())
    use_write_only = lightweight_mode or (total_cells > 150_000)

    wb = openpyxl.Workbook(write_only=use_write_only)
    if not use_write_only and wb.active:
        wb.remove(wb.active)

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

    for sheet_name, sheet_df in sheets_dict.items():
        ws = wb.create_sheet(title=sheet_name)
        clean_df = sheet_df.dropna(how="all", axis=1).fillna("")
        headers = list(clean_df.columns)
        if not headers:
            continue

        # Fast vectorized column width calculation in Pandas (avoids scanning cells in openpyxl)
        for col_idx, col_name in enumerate(headers, start=1):
            sample_max = (
                clean_df[col_name].iloc[:500].astype(str).str.len().max()
                if not clean_df.empty
                else 0
            )
            max_len = max(len(str(col_name)), int(sample_max if pd.notna(sample_max) else 0))
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 45)

        if use_write_only:
            # Write-only mode keeps RAM under 20 MB even for tens of thousands of rows
            styled_headers = []
            for h in headers:
                cell = openpyxl.cell.WriteOnlyCell(ws, value=h)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                styled_headers.append(cell)
            ws.append(styled_headers)

            for row_tuple in clean_df.itertuples(index=False, name=None):
                ws.append(list(row_tuple))
        else:
            ws.views.sheetView[0].showGridLines = True
            ws.freeze_panes = "A2"
            ws.append(headers)
            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=1, column=col_idx)
                cell.fill, cell.font = header_fill, header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            for row_idx, row_tuple in enumerate(clean_df.itertuples(index=False, name=None), start=2):
                ws.append(list(row_tuple))
                is_even = row_idx % 2 == 0
                for col_idx in range(1, len(headers) + 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    cell.font, cell.border = data_font, thin_border
                    cell.alignment = Alignment(vertical="center")
                    if is_even:
                        cell.fill = alt_fill

    wb.save(output)
    gc.collect()
    return output.getvalue()


# --- Streamlit Sidebar & UI ---
st.title("📊 DHIS2 TB Tracker Data Exporter")
st.caption("High-speed multi-page & multi-OrgUnit extractor optimized for Streamlit Community Cloud.")

with st.sidebar:
    st.header("⚙️ Connection & Filters")
    with st.form("dhis2_form"):
        web_url = st.text_input("DHIS2 Base URL", value="https://mbdrtraining.moh.gov.mm/train")
        username = st.text_input("Username", value="NTP_YTP_Central_S2")
        password = st.text_input("Password", type="password", value="District@1")

        st.divider()
        selected_programs = st.multiselect(
            "Select Programs",
            options=list(PROGRAM_MAP.keys()),
            default=list(PROGRAM_MAP.keys()),
        )

        selected_townships = st.multiselect(
            "Select Townships (Registration Filter)",
            options=list(TOWNSHIP_OU_MAP.keys()),
            default=list(TOWNSHIP_OU_MAP.keys()),
        )

        lightweight_excel = st.checkbox(
            "Fast/Low-RAM Excel Mode (Recommended for Cloud)",
            value=True,
            help="Uses openpyxl write_only stream mode to prevent 1GB RAM crashes on large exports.",
        )

        submitted = st.form_submit_button("🚀 Extract & Process Data", type="primary", use_container_width=True)

if submitted:
    prog_ids = list(dict.fromkeys([REGISTRATION_PROGRAM_ID] + [PROGRAM_MAP[p] for p in selected_programs]))
    selected_ou_ids = []
    for township in selected_townships:
        selected_ou_ids.extend(TOWNSHIP_OU_MAP[township])

    if not selected_ou_ids:
        st.warning("Please select at least one Township.")
    else:
        with st.status("Extracting and processing DHIS2 data...", expanded=True) as status:
            df_dhis2 = get_data_dhis2(
                web_url, username, password, prog_ids, selected_ou_ids, status_container=status
            )

            if df_dhis2.empty:
                status.update(label="Extraction failed or returned 0 records.", state="error", expanded=True)
                st.error("No data extracted. Please verify your URL, credentials, or user Org Unit permissions.")
            else:
                status.write("⚙️ Filtering `df_registration` by selected Org Units and matching UPIs across stages...")
                df_registration, df_combine, df_ygntbpro, sheets_dict = process_dhis2_data(
                    df_dhis2=df_dhis2,
                    selected_ou_ids=selected_ou_ids,
                    selected_township_names=selected_townships,
                )

                status.write("📦 Building Excel workbook in memory...")
                excel_bytes = convert_df_to_excel_bytes(sheets_dict, lightweight_mode=lightweight_excel)

                # Store in session_state so UI interactions don't re-run the extraction
                st.session_state["dhis2_results"] = {
                    "df_dhis2": df_dhis2,
                    "df_registration": df_registration,
                    "df_combine": df_combine,
                    "df_ygntbpro": df_ygntbpro,
                    "excel_bytes": excel_bytes,
                }
                status.update(label="✅ Extraction & Processing Complete!", state="complete", expanded=False)

# --- Display Results from Session State ---
if "dhis2_results" in st.session_state:
    res = st.session_state["dhis2_results"]
    df_dhis2 = res["df_dhis2"]
    df_registration = res["df_registration"]
    df_combine = res["df_combine"]
    df_ygntbpro = res["df_ygntbpro"]

    # KPI Summary Cards
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total DHIS2 Rows (df_dhis2)", f"{len(df_dhis2):,}")
    m2.metric("Registered UPIs (df_registration)", f"{len(df_registration):,}")
    m3.metric("Combined UPIs (df_combine)", f"{len(df_combine):,}")
    m4.metric("YgnTBPro Columns", f"{df_ygntbpro.shape[1]:,}")

    st.download_button(
        label="💾 Download Excel Workbook (.xlsx)",
        data=res["excel_bytes"],
        file_name="DHIS2_Tracker_Export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    st.divider()

    # Interactive Preview Tabs
    tab1, tab2, tab3, tab4 = st.tabs([
        "📋 YgnTBPro (df_ygntbpro)",
        "🔗 Combined (df_combine)",
        "🏥 Registration (df_registration)",
        "🗄️ All Server Data (df_dhis2)",
    ])

    with tab1:
        st.dataframe(df_ygntbpro, use_container_width=True, height=420)
    with tab2:
        st.dataframe(df_combine, use_container_width=True, height=420)
    with tab3:
        st.dataframe(df_registration, use_container_width=True, height=420)
    with tab4:
        st.dataframe(df_dhis2, use_container_width=True, height=420)
