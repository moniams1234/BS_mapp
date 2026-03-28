"""
xlsx_parser.py
Heuristic parser for trial balance XLSX files.
Supports Polish and English column names, variable header rows,
multiple sheet formats.
"""
from __future__ import annotations

import re
from typing import Optional

import numpy as np
import pandas as pd


# ---------------------------------------------------------------------------
# Column name normalization map
# ---------------------------------------------------------------------------
COLUMN_ALIASES: dict[str, list[str]] = {
    "account_number": [
        "numer", "konto", "account", "account_number", "account number",
        "nr konta", "numer konta", "account no", "acc",
    ],
    "account_name2": [
        "nazwa 2", "name2", "nazwa2", "account name 2", "label",
    ],
    "account_name": [
        "nazwa", "name", "account_name", "account name", "opis",
        "description", "account desc",
    ],
    "bo_dt": [
        "bo dt", "bo dt ", "opening debit", "od dt", "saldo poczatkowe dt",
    ],
    "bo_ct": [
        "bo ct", "bo ct ", "opening credit", "od ct", "saldo poczatkowe ct",
    ],
    "obroty_dt": [
        "obroty dt", "obroty dt ", "turnover debit", "debit turnover",
        "debit", "wn", "dt",
    ],
    "obroty_ct": [
        "obroty ct", "obroty ct ", "turnover credit", "credit turnover",
        "credit", "ma", "ct",
    ],
    "obroty_n_dt": [
        "obroty n. dt", "obroty n. dt ", "net debit", "net dt",
    ],
    "obroty_n_ct": [
        "obroty n. ct", "obroty n. ct ", "net credit", "net ct",
    ],
    "saldo_dt": [
        "saldo dt", "saldo dt ", "closing debit", "balance debit",
    ],
    "saldo_ct": [
        "saldo ct", "saldo ct ", "closing credit", "balance credit",
    ],
    "persaldo": [
        "persaldo", "saldo", "balance", "net balance", "net saldo",
    ],
    "monthly": ["monthly", "miesięczne", "miesieczne"],
    "bs_mapp": [
        "bs mapp", "bs_mapp", "mapp", "mapping", "mapowanie", "group",
    ],
}

# Minimum columns required to consider a sheet a trial balance
REQUIRED_COLS = {"account_number"}
NUMERIC_COLS = ["bo_dt", "bo_ct", "obroty_dt", "obroty_ct",
                "obroty_n_dt", "obroty_n_ct", "saldo_dt", "saldo_ct", "persaldo"]

# Score keywords that indicate a trial balance sheet
TB_SHEET_KEYWORDS = [
    "zois", "trial balance", "balance", "zestawienie", "obroty", "saldo",
    "trial", "tb", "ksiega", "ledger", "accounts",
]


def _strip(s: object) -> str:
    """Lower-strip a cell value to string."""
    if pd.isna(s):
        return ""
    return str(s).strip().lower()


def _score_sheet(name: str) -> int:
    """Score a sheet name for likelihood of containing trial balance data."""
    name_l = name.lower()
    score = 0
    for kw in TB_SHEET_KEYWORDS:
        if kw in name_l:
            score += 2
    # Penalize sheets that are clearly not TB
    for bad in ["mapp", "hyperion", "check", "pivot", "summary", "chart"]:
        if bad in name_l:
            score -= 3
    return score


def _normalize_col_name(raw: str) -> str:
    """Map raw column name to a canonical internal name."""
    raw_l = raw.strip().lower()
    for canonical, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if raw_l == alias:
                return canonical
    return raw_l.replace(" ", "_")


def _detect_header_row(df_raw: pd.DataFrame) -> int:
    """
    Scan up to 20 rows to find the row that contains the most
    recognizable column names.
    """
    best_row = 0
    best_score = 0
    for i in range(min(20, len(df_raw))):
        row = df_raw.iloc[i]
        score = 0
        for cell in row:
            normalized = _normalize_col_name(_strip(cell))
            for canonical, aliases in COLUMN_ALIASES.items():
                if _strip(cell).lower() in aliases or normalized == canonical:
                    score += 1
        if score > best_score:
            best_score = score
            best_row = i
    return best_row


def _parse_number(val: object) -> float:
    """Coerce any value to float; return 0.0 on failure."""
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(" ", "").replace("\xa0", "")
    s = s.replace(",", ".")
    # Remove trailing / leading garbage
    match = re.search(r"-?\d+(?:\.\d+)?", s)
    if match:
        return float(match.group())
    return 0.0


def _select_sheet(xls: pd.ExcelFile) -> str:
    """Heuristically select the best sheet for trial balance data."""
    names = xls.sheet_names
    if len(names) == 1:
        return names[0]
    scores = {name: _score_sheet(name) for name in names}
    # Also check size: prefer larger sheets
    for name in names:
        try:
            df = xls.parse(name, header=None, nrows=5)
            scores[name] += min(df.shape[1], 10)  # wider = more likely to be TB
        except Exception:
            pass
    return max(scores, key=lambda n: scores[n])


def _build_dataframe(xls: pd.ExcelFile, sheet: str) -> pd.DataFrame:
    """
    Read a sheet, detect header row, normalize column names,
    clean data, compute persaldo.
    """
    df_raw = xls.parse(sheet, header=None, dtype=str)
    header_row = _detect_header_row(df_raw)

    # Use detected row as header
    df = xls.parse(sheet, header=header_row, dtype=str)
    df.columns = [_normalize_col_name(str(c)) for c in df.columns]

    # Remove completely empty rows
    df.dropna(how="all", inplace=True)
    df.reset_index(drop=True, inplace=True)

    # Parse numeric columns
    for col in NUMERIC_COLS:
        if col in df.columns:
            df[col] = df[col].apply(_parse_number)
        else:
            df[col] = 0.0

    # Ensure account columns exist
    if "account_number" not in df.columns:
        # Try to use the first column
        df.rename(columns={df.columns[0]: "account_number"}, inplace=True)

    if "account_name" not in df.columns and "account_name2" in df.columns:
        df["account_name"] = df["account_name2"]
    elif "account_name" not in df.columns:
        df["account_name"] = df["account_number"].astype(str)

    # Remove rows where account_number is empty / NaN
    df = df[df["account_number"].notna() & (df["account_number"].str.strip() != "")]

    # Compute persaldo if not present or all zero
    if "persaldo" not in df.columns or df["persaldo"].sum() == 0:
        # persaldo = saldo_dt - saldo_ct  (positive = debit balance)
        df["persaldo"] = df["saldo_dt"] - df["saldo_ct"]
        # If saldo_dt/ct also zero, try BO - based saldo
        mask_zero = df["persaldo"] == 0
        if mask_zero.all():
            df["persaldo"] = df["bo_dt"] - df["bo_ct"]

    # Keep bs_mapp column if present
    if "bs_mapp" not in df.columns:
        df["bs_mapp"] = ""

    df["account_number"] = df["account_number"].astype(str).str.strip()
    df["account_name"] = df["account_name"].astype(str).str.strip()

    return df


def parse_xlsx(file_obj) -> dict:
    """
    Main entry point.
    Returns dict with keys:
      - df: cleaned DataFrame
      - sheet_used: str
      - warnings: list[str]
      - columns_found: list[str]
    """
    warnings: list[str] = []
    try:
        xls = pd.ExcelFile(file_obj, engine="openpyxl")
    except Exception as e:
        return {"error": f"Cannot open file: {e}"}

    sheet = _select_sheet(xls)
    try:
        df = _build_dataframe(xls, sheet)
    except Exception as e:
        return {"error": f"Cannot parse sheet '{sheet}': {e}"}

    if df.empty:
        warnings.append("Parsed DataFrame is empty after cleaning.")

    missing_numeric = [c for c in ["saldo_dt", "saldo_ct", "persaldo"] if df[c].sum() == 0]
    if missing_numeric:
        warnings.append(f"All-zero numeric columns detected: {missing_numeric}. Check column mapping.")

    return {
        "df": df,
        "sheet_used": sheet,
        "all_sheets": xls.sheet_names,
        "warnings": warnings,
        "columns_found": list(df.columns),
    }
