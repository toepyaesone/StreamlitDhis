from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date
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

# --- Constants & Program / OrgUnit Mappings ---
REGISTRATION_PROGRAM_ID = "UZF0HrTlps0"
DIAGNOSTIC_PROGRAM_ID = "GvywHD6crky"
TBCS_PROGRAM_ID = "Lt6P15ps7f6"
TPT_PROGRAM_ID = "cQsXTtAJ3HW"

PROGRAM_MAP = {
    "Registration & Screening": REGISTRATION_PROGRAM_ID,
    "Diagnostic Evaluation": DIAGNOSTIC_PROGRAM_ID,
    "TB Case Surveillance": TBCS_PROGRAM_ID,
    "TB Contact Investigation TPT": TPT_PROGRAM_ID,
}

# Reverse lookup for clean display names
PROGRAM_ID_TO_LABEL = {v: k for k, v in PROGRAM_MAP.items()}

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
    session.headers.update({"Accept": "application/json", "User-Agent": "DHIS2-Streamlit-Exporter/4.0"})
    retries = Retry(total=3, backoff_factor=1, status_forcelist=[500, 502, 503, 504])
    adapter = HTTPAdapter(pool_connections=20, pool_maxsize=20, max_retries=retries)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_cached_metadata(base_url: str, username: str, password: str) -> tuple[dict[str, str], list[str], bool]:
    """Fetches and caches DHIS2 metadata mappings, accessible user OUs, and endpoint type."""
    session = create_http_session(username, password)
    meta_map: dict[str, str] = {}

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
    prog_name = PROGRAM_ID_TO_LABEL.get(prog) or meta_map.get(prog, prog)

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
    """Paginates through all pages for a given program and organization unit scope."""
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
                    "enrollments[enrollment,program,orgUnit,orgUnitName,enrolledAt,enrollmentDate,occurredAt,status,"
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
    Extracts all data across multiple pages and organization units from the DHIS2 server
    to build the complete `df_dhis2` DataFrame.
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
        prog_records = fetch_paginated_tei(
            session=session,
            active_endpoint=active_endpoint,
            is_legacy=is_legacy,
            prog=prog,
            meta_map=meta_map,
            ou=None,
            ou_mode="ACCESSIBLE",
        )
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

    with ThreadPoolExecutor(max_workers=len(idprogram)) as prog_executor:
        future_to_prog = {prog_executor.submit(_fetch_program, prog): prog for prog in idprogram}
        for future in as_completed(future_to_prog):
            prog_id = future_to_prog[future]
            prog_label = PROGRAM_ID_TO_LABEL.get(prog_id) or meta_map.get(prog_id, prog_id)
            res_records = future.result()
            all_records.extend(res_records)
            if status_container:
                status_container.write(f"✅ Extracted **{len(res_records):,}** raw server records from *{prog_label}*")

    session.close()

    df_dhis2 = pd.DataFrame(all_records)
    if not df_dhis2.empty:
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
    """Vectorized combination of unique non-empty cell values grouped by `join_key`."""
    if df.empty or join_key not in df.columns:
        return pd.DataFrame(columns=column_order)

    valid_df = df[df[join_key].notna()].copy()
    if valid_df.empty:
        return pd.DataFrame(columns=column_order)

    if valid_df[join_key].is_unique:
        return valid_df.reindex(columns=column_order).reset_index(drop=True)

    non_empty_cols = [c for c in valid_df.columns if c != join_key and valid_df[c].notna().any()]
    if not non_empty_cols:
        return valid_df[[join_key]].drop_duplicates().reindex(columns=column_order).reset_index(drop=True)

    melted = (
        valid_df[[join_key] + non_empty_cols]
        .melt(id_vars=join_key, var_name="_col", value_name="_val")
        .dropna(subset=["_val"])
    )
    melted["_val"] = melted["_val"].astype(str).str.strip()
    melted = melted[~melted["_val"].isin(["", "nan", "None", "NaN"])]

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


def count_unique_upis(df: pd.DataFrame, join_key: str) -> int:
    """Returns the count of unique non-null UPIs in a DataFrame."""
    if df.empty or join_key not in df.columns:
        return 0
    return int(df[join_key].dropna().nunique())


def process_dhis2_data(
    df_dhis2: pd.DataFrame,
    selected_ou_ids: list[str],
    selected_township_names: list[str] | None = None,
    start_date_str: str | None = None,
    end_date_str: str | None = None,
) -> dict:
    """
    Processes `df_dhis2` according to the required workflow:
    1. Creates Registration & Screening (`df_registration`) based on user OU & date selection and extracts unique UPIs.
    2. Processes Diagnostic Evaluation (`df_diagnostic`) and remaining pages (`df_tbcs`, `df_tpt`) using `df_registration` UPIs regardless of Org Units.
    3. Creates `df_combine` and `df_ygntbpro`.
    4. Builds page-by-page summary statistics (`df_summary`) and `excel_sheets_dict` (excluding `df_dhis2`).
    """
    all_cols = list(df_dhis2.columns)
    join_key = "Unique ID (UPI)" if "Unique ID (UPI)" in df_dhis2.columns else "trackedEntityInstance"

    df_dhis2 = df_dhis2.copy()
    if join_key in df_dhis2.columns:
        df_dhis2[join_key] = df_dhis2[join_key].astype(str).str.strip()
        df_dhis2.loc[df_dhis2[join_key].isin(["", "None", "nan", "NaN"]), join_key] = None

    # --- 1. Split Raw Server Data by Program ---
    df_reg_raw = df_dhis2[df_dhis2["program_id"] == REGISTRATION_PROGRAM_ID].copy()
    df_diag_raw = df_dhis2[df_dhis2["program_id"] == DIAGNOSTIC_PROGRAM_ID].copy()
    df_tbcs_raw = df_dhis2[df_dhis2["program_id"] == TBCS_PROGRAM_ID].copy()
    df_tpt_raw = df_dhis2[df_dhis2["program_id"] == TPT_PROGRAM_ID].copy()

    # --- 2. Filter Registration & Screening by User Selection & Criteria ---
    df_reg_filtered = df_reg_raw.copy()

    # Criterion A: Selected Organization Units
    if selected_ou_ids and not df_reg_filtered.empty:
        ou_mask = pd.Series(False, index=df_reg_filtered.index)
        if "orgUnit_id" in df_reg_filtered.columns:
            ou_mask |= df_reg_filtered["orgUnit_id"].isin(selected_ou_ids)
        if "tei_orgUnit_id" in df_reg_filtered.columns:
            ou_mask |= df_reg_filtered["tei_orgUnit_id"].isin(selected_ou_ids)
        if selected_township_names and "Org Unit Name" in df_reg_filtered.columns:
            ou_mask |= df_reg_filtered["Org Unit Name"].isin(selected_township_names)
        df_reg_filtered = df_reg_filtered[ou_mask].copy()

    # Criterion B: Enrollment Date Range (if enabled by user)
    if (start_date_str or end_date_str) and "enrollment_date" in df_reg_filtered.columns and not df_reg_filtered.empty:
        enr_dates = df_reg_filtered["enrollment_date"].astype(str).str.slice(0, 10)
        if start_date_str:
            df_reg_filtered = df_reg_filtered[enr_dates >= start_date_str]
            enr_dates = df_reg_filtered["enrollment_date"].astype(str).str.slice(0, 10)
        if end_date_str:
            df_reg_filtered = df_reg_filtered[enr_dates <= end_date_str]

    # Create `df_registration` (unique UPI per row) and extract valid UPI set
    df_registration = fast_combine_by_key(df_reg_filtered, join_key, all_cols)
    registration_upis = (
        set(df_registration[join_key].dropna().unique())
        if join_key in df_registration.columns and not df_registration.empty
        else set()
    )

    # --- 3. Process Diagnostic Evaluation and Remaining Pages (Matched by UPI Regardless of OU) ---
    if registration_upis and join_key in df_dhis2.columns:
        df_diagnostic = df_diag_raw[df_diag_raw[join_key].isin(registration_upis)].copy()
        df_tbcs = df_tbcs_raw[df_tbcs_raw[join_key].isin(registration_upis)].copy()
        df_tpt = df_tpt_raw[df_tpt_raw[join_key].isin(registration_upis)].copy()
    else:
        df_diagnostic = pd.DataFrame(columns=all_cols)
        df_tbcs = pd.DataFrame(columns=all_cols)
        df_tpt = pd.DataFrame(columns=all_cols)

    # --- 4. Create `df_combine` and `df_ygntbpro` ---
    df_relevant = pd.concat([df_registration, df_diagnostic, df_tbcs, df_tpt], ignore_index=True)
    df_combine = fast_combine_by_key(df_relevant, join_key, all_cols)

    existing_cols = [col for col in SELECTED_COLUMN_LIST if col in df_combine.columns]
    df_ygntbpro = df_combine[existing_cols].copy() if existing_cols else pd.DataFrame()

    # --- 5. Build Data Summary Table (Record Counts & Unique UPIs per Page) ---
    summary_rows = [
        {
            "Page / Dataset": "1. Registration & Screening",
            "Raw Server Records (All OUs)": len(df_reg_raw),
            "Filtered Cohort Records": len(df_reg_filtered),
            "Processed / Deduplicated Rows": len(df_registration),
            "Unique UPIs": count_unique_upis(df_registration, join_key),
            "Export Destination": "Excel Sheet (Registration & Screening)",
        },
        {
            "Page / Dataset": "2. Diagnostic Evaluation",
            "Raw Server Records (All OUs)": len(df_diag_raw),
            "Filtered Cohort Records": len(df_diagnostic),
            "Processed / Deduplicated Rows": len(df_diagnostic),
            "Unique UPIs": count_unique_upis(df_diagnostic, join_key),
            "Export Destination": "Excel Sheet (Diagnostic Evaluation)",
        },
        {
            "Page / Dataset": "3. TB Case Surveillance",
            "Raw Server Records (All OUs)": len(df_tbcs_raw),
            "Filtered Cohort Records": len(df_tbcs),
            "Processed / Deduplicated Rows": len(df_tbcs),
            "Unique UPIs": count_unique_upis(df_tbcs, join_key),
            "Export Destination": "Excel Sheet (TB Case Surveillance)",
        },
        {
            "Page / Dataset": "4. TB Contact Investigation TPT",
            "Raw Server Records (All OUs)": len(df_tpt_raw),
            "Filtered Cohort Records": len(df_tpt),
            "Processed / Deduplicated Rows": len(df_tpt),
            "Unique UPIs": count_unique_upis(df_tpt, join_key),
            "Export Destination": "Excel Sheet (TB Contact Investigation TPT)",
        },
        {
            "Page / Dataset": "5. Combined (df_combine)",
            "Raw Server Records (All OUs)": len(df_dhis2),
            "Filtered Cohort Records": len(df_relevant),
            "Processed / Deduplicated Rows": len(df_combine),
            "Unique UPIs": count_unique_upis(df_combine, join_key),
            "Export Destination": "Excel Sheet (Combined)",
        },
        {
            "Page / Dataset": "6. YgnTBPro (df_ygntbpro)",
            "Raw Server Records (All OUs)": len(df_dhis2),
            "Filtered Cohort Records": len(df_relevant),
            "Processed / Deduplicated Rows": len(df_ygntbpro),
            "Unique UPIs": count_unique_upis(df_ygntbpro, join_key),
            "Export Destination": f"Excel Sheet (YgnTBPro - {len(existing_cols)} cols)",
        },
        {
            "Page / Dataset": "7. Raw DHIS2 Server Data (df_dhis2)",
            "Raw Server Records (All OUs)": len(df_dhis2),
            "Filtered Cohort Records": len(df_dhis2),
            "Processed / Deduplicated Rows": len(df_dhis2),
            "Unique UPIs": count_unique_upis(df_dhis2, join_key),
            "Export Destination": "Optional CSV Download Only (Excluded from Excel)",
        },
    ]
    df_summary = pd.DataFrame(summary_rows)

    # --- 6. Prepare Excel Sheets Dictionary (Excludes `df_dhis2` to keep Excel fast & small) ---
    excel_sheets_dict: dict[str, pd.DataFrame] = {
        "Summary": df_summary,
        "YgnTBPro": df_ygntbpro,
        "Combined": df_combine,
        "Registration & Screening": df_registration,
        "Diagnostic Evaluation": df_diagnostic,
        "TB Case Surveillance": df_tbcs,
        "TB Contact Investigation TPT": df_tpt,
    }

    return {
        "df_dhis2": df_dhis2,
        "df_registration": df_registration,
        "df_diagnostic": df_diagnostic,
        "df_tbcs": df_tbcs,
        "df_tpt": df_tpt,
        "df_combine": df_combine,
        "df_ygntbpro": df_ygntbpro,
        "df_summary": df_summary,
        "excel_sheets_dict": excel_sheets_dict,
    }


def convert_df_to_excel_bytes(sheets_dict: dict[str, pd.DataFrame], lightweight_mode: bool = False) -> bytes:
    """
    Converts the processed sheets dictionary (without raw `df_dhis2`) into a formatted Excel workbook.
    """
    output = io.BytesIO()
    total_cells = sum(df.shape[0] * df.shape[1] for df in sheets_dict.values())
    use_write_only = lightweight_mode or (total_cells > 180_000)

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
        ws = wb.create_sheet(title=sheet_name[:31])
        clean_df = sheet_df.dropna(how="all", axis=1).fillna("") if not sheet_df.empty else sheet_df.fillna("")
        headers = list(clean_df.columns)
        if not headers:
            continue

        # Vectorized column width calculation in Pandas
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


# --- Streamlit Sidebar & User Selection ---
st.title("📊 DHIS2 TB Tracker Data Exporter")
st.caption("Extracts all DHIS2 server records, filters Registration & follow-up pages by UPI, and exports processed sheets to Excel with optional CSV export for raw server data.")

with st.sidebar:
    st.header("⚙️ Server & Criteria Selection")
    with st.form("dhis2_form"):
        web_url = st.text_input("DHIS2 Base URL", value="https://mbdrtraining.moh.gov.mm/train")
        username = st.text_input("Username", value="NTP_YTP_Central_S2")
        password = st.text_input("Password", type="password", value="District@1")

        st.divider()
        st.subheader("🏥 Registration Criteria")
        selected_townships = st.multiselect(
            "Select Townships (Org Units)",
            options=list(TOWNSHIP_OU_MAP.keys()),
            default=list(TOWNSHIP_OU_MAP.keys()),
        )

        use_date_filter = st.checkbox("Filter Registration by Enrollment Date", value=False)
        d_col1, d_col2 = st.columns(2)
        with d_col1:
            start_date_input = st.date_input("Start Date", value=date(date.today().year, 1, 1))
        with d_col2:
            end_date_input = st.date_input("End Date", value=date.today())

        st.divider()
        st.subheader("📄 Programs & Export Options")
        selected_programs = st.multiselect(
            "Select Programs to Extract",
            options=list(PROGRAM_MAP.keys()),
            default=list(PROGRAM_MAP.keys()),
        )

        enable_raw_csv = st.checkbox(
            "Prepare Optional CSV for Raw DHIS2 Data (df_dhis2)",
            value=True,
            help="Generates a UTF-8 CSV download for the full df_dhis2 dataset instead of writing it to Excel.",
        )

        lightweight_excel = st.checkbox(
            "Fast Stream Mode for Excel",
            value=False,
            help="Enable if your filtered cohort exceeds 10,000 rows.",
        )

        submitted = st.form_submit_button("🚀 Extract & Process Data", type="primary", use_container_width=True)

if submitted:
    # Always include Registration & Screening so cohort UPIs can be built
    prog_ids = list(dict.fromkeys([REGISTRATION_PROGRAM_ID] + [PROGRAM_MAP[p] for p in selected_programs]))
    selected_ou_ids = []
    for township in selected_townships:
        selected_ou_ids.extend(TOWNSHIP_OU_MAP[township])

    if not selected_ou_ids:
        st.warning("Please select at least one Township for Registration & Screening.")
    elif use_date_filter and start_date_input > end_date_input:
        st.error("Start Date cannot be later than End Date.")
    else:
        start_str = start_date_input.strftime("%Y-%m-%d") if use_date_filter else None
        end_str = end_date_input.strftime("%Y-%m-%d") if use_date_filter else None

        with st.status("Extracting all data from DHIS2 server and processing pages...", expanded=True) as status:
            df_dhis2 = get_data_dhis2(
                web=web_url,
                username=username,
                password=password,
                idprogram=prog_ids,
                idou=selected_ou_ids,
                status_container=status,
            )

            if df_dhis2.empty:
                status.update(label="Extraction returned 0 records.", state="error", expanded=True)
                st.error("No data extracted from DHIS2 server. Check your URL, credentials, or user permissions.")
            else:
                status.write(f"📊 Total raw records extracted into `df_dhis2`: **{len(df_dhis2):,}**")
                status.write("⚙️ Processing Registration & Screening, Diagnostic Evaluation, and remaining pages...")

                processed = process_dhis2_data(
                    df_dhis2=df_dhis2,
                    selected_ou_ids=selected_ou_ids,
                    selected_township_names=selected_townships,
                    start_date_str=start_str,
                    end_date_str=end_str,
                )

                status.write("📦 Generating formatted Excel workbook (excluding raw `df_dhis2`)...")
                excel_bytes = convert_df_to_excel_bytes(
                    processed["excel_sheets_dict"],
                    lightweight_mode=lightweight_excel,
                )

                csv_bytes = None
                if enable_raw_csv:
                    status.write("📄 Preparing optional UTF-8 CSV for raw `df_dhis2`...")
                    csv_bytes = df_dhis2.to_csv(index=False).encode("utf-8-sig")

                st.session_state["dhis2_results"] = {
                    **processed,
                    "excel_bytes": excel_bytes,
                    "csv_bytes": csv_bytes,
                }
                status.update(label="✅ Extraction & Processing Complete!", state="complete", expanded=False)

# --- Display Summary, Downloads & Data Previews ---
if "dhis2_results" in st.session_state:
    res = st.session_state["dhis2_results"]
    df_dhis2 = res["df_dhis2"]
    df_registration = res["df_registration"]
    df_diagnostic = res["df_diagnostic"]
    df_tbcs = res["df_tbcs"]
    df_tpt = res["df_tpt"]
    df_combine = res["df_combine"]
    df_ygntbpro = res["df_ygntbpro"]
    df_summary = res["df_summary"]

    # 1. Top-Level Record Count Metrics Across All Pages
    st.subheader("📈 Data Summary & Record Counts by Page")
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Total DHIS2 Server (df_dhis2)", f"{len(df_dhis2):,}", help="All records extracted from DHIS2 server across all programs & OUs")
    c2.metric("1. Registration & Screening", f"{len(df_registration):,}", help="Unique UPIs matching selected OUs & criteria")
    c3.metric("2. Diagnostic Evaluation", f"{len(df_diagnostic):,}", help="Records matching Registration UPIs across all OUs")
    c4.metric("3. TB Case Surveillance", f"{len(df_tbcs):,}", help="Records matching Registration UPIs across all OUs")
    c5.metric("4. TPT Investigation", f"{len(df_tpt):,}", help="Records matching Registration UPIs across all OUs")
    c6.metric("Combined / YgnTBPro", f"{len(df_ygntbpro):,}", help="Final deduplicated cohort rows")

    # 2. Detailed Summary Table
    st.dataframe(df_summary, use_container_width=True, hide_index=True)

    # 3. Export & Download Section (Excel + Optional Raw CSV)
    st.subheader("💾 Export & Download")
    dl_col1, dl_col2 = st.columns(2)
    with dl_col1:
        st.download_button(
            label="📥 Download Processed Excel Workbook (.xlsx)",
            data=res["excel_bytes"],
            file_name="DHIS2_Processed_TB_Tracker.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
            help="Contains Summary, YgnTBPro, Combined, Registration & Screening, Diagnostic Evaluation, TB Case Surveillance, and TPT sheets.",
        )
    with dl_col2:
        if res.get("csv_bytes") is not None:
            st.download_button(
                label=f"📄 Download Raw DHIS2 Server Data (.csv) [{len(df_dhis2):,} rows]",
                data=res["csv_bytes"],
                file_name="DHIS2_Raw_Server_Data.csv",
                mime="text/csv",
                use_container_width=True,
                help="Full raw df_dhis2 dataset encoded in UTF-8-SIG for Excel compatibility.",
            )
        else:
            if st.button("⚙️ Generate Raw DHIS2 Data CSV on Demand", use_container_width=True):
                st.session_state["dhis2_results"]["csv_bytes"] = df_dhis2.to_csv(index=False).encode("utf-8-sig")
                st.rerun()

    st.divider()

    # 4. Interactive Page Preview Tabs
    st.subheader("🔍 Preview Processed Pages")
    tabs = st.tabs([
        f"📋 YgnTBPro ({len(df_ygntbpro):,})",
        f"🔗 Combined ({len(df_combine):,})",
        f"🏥 1. Registration ({len(df_registration):,})",
        f"🔬 2. Diagnostic Eval ({len(df_diagnostic):,})",
        f"💊 3. TB Case Surveillance ({len(df_tbcs):,})",
        f"🛡️ 4. TPT ({len(df_tpt):,})",
        f"🗄️ Raw Server Data ({len(df_dhis2):,})",
    ])

    with tabs[0]:
        st.dataframe(df_ygntbpro, use_container_width=True, height=400)
    with tabs[1]:
        st.dataframe(df_combine, use_container_width=True, height=400)
    with tabs[2]:
        st.dataframe(df_registration, use_container_width=True, height=400)
    with tabs[3]:
        st.dataframe(df_diagnostic, use_container_width=True, height=400)
    with tabs[4]:
        st.dataframe(df_tbcs, use_container_width=True, height=400)
    with tabs[5]:
        st.dataframe(df_tpt, use_container_width=True, height=400)
    with tabs[6]:
        st.caption("Showing top 1,000 rows of raw `df_dhis2` for browser performance. Use the CSV button above to download all rows.")
        st.dataframe(df_dhis2.head(1000), use_container_width=True, height=400)
