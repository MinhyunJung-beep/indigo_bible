import sys
from pathlib import Path
import gspread
import numpy as np
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

BASE_DIR = Path(__file__).resolve().parent
FALLBACK_EXCEL_PATH = BASE_DIR / "tcr_diagnosis.xlsx"
ERROR_TOKENS = {"#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#ERROR!", "Err:509"}

# 기존 엑셀 탭 이름 -> 구글 시트 실제 탭 이름 후보 자동 매칭
TAB_ALIASES = {
    "Race_Result_DB": ["RACE_RESULT", "Race_Result_DB"],
    "Tire_Analysis_DB": ["TIRE_ANALYSIS_87", "Tire_Analysis_DB"],
    "PART_STATUS V2": ["PART_STATUS V2", "PART_STATUS"],
    "Performance_Trend_Analysis": ["Performance Trend Analysis", "Performance_Trend_Analysis"],
}


def _norm(name: str) -> str:
    return str(name).strip().lower().replace("_", " ").replace("-", " ")


def _find_key_file() -> Path | None:
    streamlit_dir = BASE_DIR / ".streamlit"
    if not streamlit_dir.exists():
        return None
    candidates = sorted(streamlit_dir.glob("google_key*.json"))
    return candidates[0] if candidates else None


def _deduplicate_headers(raw_headers: list) -> list[str]:
    seen: dict[str, int] = {}
    clean_headers: list[str] = []
    for idx, col in enumerate(raw_headers):
        name = str(col).strip() if col is not None and str(col).strip() != "" else f"Unnamed: {idx}"
        if name in seen:
            seen[name] += 1
            clean_headers.append(f"{name}.{seen[name]}")
        else:
            seen[name] = 0
            clean_headers.append(name)
    return clean_headers


def _clean_cell(val):
    if val is None:
        return np.nan
    if isinstance(val, str):
        s = val.strip()
        if s == "" or s in ERROR_TOKENS:
            return np.nan
        return s
    return val


def _clean_dataframe(rows: list[list]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame()

    headers = _deduplicate_headers(rows[0])
    width = len(headers)
    normalized_rows = []
    for r in rows[1:]:
        cleaned_row = [_clean_cell(c) for c in (r + [None] * (width - len(r)))[:width]]
        if any(pd.notna(c) for c in cleaned_row):
            normalized_rows.append(cleaned_row)

    df = pd.DataFrame(normalized_rows, columns=headers)
    for col in df.columns:
        if not pd.api.types.is_numeric_dtype(df[col]):
            non_null = df[col].dropna()
            if not non_null.empty:
                numeric_col = pd.to_numeric(non_null, errors="coerce")
                if numeric_col.notna().all():
                    df[col] = pd.to_numeric(df[col], errors="coerce")
    return df


@st.cache_resource
def get_gspread_client() -> gspread.Client | None:
    key_path = _find_key_file()
    if key_path and key_path.exists():
        creds = Credentials.from_service_account_file(str(key_path), scopes=SCOPES)
        return gspread.authorize(creds)
    if "gcp_service_account" in st.secrets:
        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]), scopes=SCOPES
        )
        return gspread.authorize(creds)
    return None


@st.cache_data(ttl=600, show_spinner=False)
def get_tab_directory() -> dict[str, tuple[str, str, str]]:
    """공유된 모든 구글 시트를 자동 탐색해 탭 위치를 인덱싱."""
    client = get_gspread_client()
    if client is None:
        return {}

    directory: dict[str, tuple[str, str, str]] = {}
    for f in client.list_spreadsheet_files():
        sheet_id, sheet_title = f["id"], f["name"]
        try:
            sh = client.open_by_key(sheet_id)
            meta = sh.fetch_sheet_metadata()
            for sheet_prop in meta.get("sheets", []):
                tab_title = sheet_prop["properties"]["title"]
                key = _norm(tab_title)
                if key not in directory:
                    directory[key] = (sheet_id, sheet_title, tab_title)
        except Exception:
            continue
    return directory


@st.cache_data(ttl=300, show_spinner=False)
def _fetch_google_sheet_tab(sheet_id: str, tab_name: str, header_row: int = 0) -> pd.DataFrame:
    client = get_gspread_client()
    sh = client.open_by_key(sheet_id)
    resp = sh.values_get(
        tab_name,
        params={
            "valueRenderOption": "UNFORMATTED_VALUE",
            "dateTimeRenderOption": "FORMATTED_STRING",
        },
    )
    rows = resp.get("values", [])
    if header_row > 0 and len(rows) > header_row:
        rows = rows[header_row:]
    return _clean_dataframe(rows)


def load_tess_data(tab_name: str, header_row: int = 0) -> tuple[pd.DataFrame, str]:
    """실시간 구글 시트 우선 조회 및 오프라인 시 로컬 엑셀 자동 전환."""
    try:
        directory = get_tab_directory()
        candidates = list(TAB_ALIASES.get(tab_name, [tab_name]))
        if tab_name not in candidates:
            candidates.append(tab_name)

        for cand in candidates:
            norm_cand = _norm(cand)
            if norm_cand in directory:
                sheet_id, sheet_title, real_tab = directory[norm_cand]
                df = _fetch_google_sheet_tab(sheet_id, real_tab, header_row=header_row)
                return df, f"LIVE ({sheet_title} > {real_tab})"
    except Exception as exc:
        st.sidebar.warning(f"[{tab_name}] 실시간 읽기 실패 -> 로컬 백업 전환 ({exc})")

    if FALLBACK_EXCEL_PATH.exists():
        df_local = pd.read_excel(FALLBACK_EXCEL_PATH, sheet_name=tab_name, header=header_row)
        df_local = df_local.replace(list(ERROR_TOKENS), np.nan)
        return df_local, "OFFLINE_BACKUP (tcr_diagnosis.xlsx)"

    return pd.DataFrame(), "EMPTY"


def refresh_all_sheets() -> None:
    get_tab_directory.clear()
    _fetch_google_sheet_tab.clear()


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    key_file = _find_key_file()
    if not key_file:
        print("[ERROR] .streamlit 폴더 안에서 google_key 파일을 찾지 못했습니다.")
        sys.exit(1)

    print(f"[OK] 인증키 파일 인식 성공: {key_file.name}")
    creds = Credentials.from_service_account_file(str(key_file), scopes=SCOPES)
    gc = gspread.authorize(creds)
    shared_files = gc.list_spreadsheet_files()

    if not shared_files:
        print("[WARN] 아직 공유된 구글 시트가 없습니다. 시트 '공유'에서 완료 버튼을 눌렀는지 확인해 주세요.")
    else:
        print(f"[OK] 실시간 연결된 구글 시트 총 {len(shared_files)}개 발견:")
        for idx, sf in enumerate(shared_files, 1):
            sh = gc.open_by_key(sf["id"])
            tabs = [s["properties"]["title"] for s in sh.fetch_sheet_metadata().get("sheets", [])]
            print(f"  {idx}. {sf['name']} (탭 {len(tabs)}개: {', '.join(tabs[:8])}{' ...' if len(tabs) > 8 else ''})")