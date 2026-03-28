"""
xlsx_parser.py

Parses two types of XLSX files:
  1. Trial balance (ZOiS) – returns cleaned DataFrame
  2. Mapping file – returns dict {account_number -> {side, group}}

MAPPING FILE FORMATS:

  Format A – "Mapp sheet":
    Sheet named 'Mapp', col 0 = 'A'|'P'|'X'|group-name, col 2 = account number

  Format B – "ZOiS + BS Mapp column" (default Mapping.xlsx):
    Columns: Numer | ... | BS Mapp
    BS Mapp = group label or 'x' (excluded from BS)
    Side derived from _GROUP_SIDE_MAP (fully enumerated, no heuristics)
"""
from __future__ import annotations

import re
from pathlib import Path

import pandas as pd


# ─── column normalization ─────────────────────────────────────────────────────

_COL_MAP: dict[str, list[str]] = {
    "account_number": ["numer", "konto", "account", "account_number", "account number",
                       "nr konta", "numer konta", "account no"],
    "account_name2": ["nazwa 2", "name2", "label", "nazwa2"],
    "account_name":  ["nazwa", "name", "account_name", "account name", "opis", "description"],
    "bo_dt":         ["bo dt", "bo dt "],
    "bo_ct":         ["bo ct", "bo ct "],
    "obroty_dt":     ["obroty dt", "obroty dt ", "debit", "wn", "dt"],
    "obroty_ct":     ["obroty ct", "obroty ct ", "credit", "ma", "ct"],
    "obroty_n_dt":   ["obroty n. dt", "obroty n. dt "],
    "obroty_n_ct":   ["obroty n. ct", "obroty n. ct "],
    "saldo_dt":      ["saldo dt", "saldo dt ", "closing debit", "balance debit"],
    "saldo_ct":      ["saldo ct", "saldo ct ", "closing credit", "balance credit"],
    "persaldo":      ["persaldo", "saldo", "balance"],
    "bs_mapp":       ["bs mapp", "bs_mapp", "mapp", "bs mapp "],
}


def _canonical(raw: str) -> str:
    r = raw.strip().lower()
    for canon, aliases in _COL_MAP.items():
        if r in aliases:
            return canon
    return r.replace(" ", "_")


def _find_header_row(df_raw: pd.DataFrame) -> int:
    best, best_score = 0, 0
    for i in range(min(15, len(df_raw))):
        score = sum(
            1 for cell in df_raw.iloc[i]
            if any(str(cell).strip().lower() in aliases
                   for aliases in _COL_MAP.values())
        )
        if score > best_score:
            best_score, best = score, i
    return best


def _to_float(val: object) -> float:
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(" ", "").replace("\xa0", "").replace(",", ".")
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    return float(m.group()) if m else 0.0


def _select_tb_sheet(xls: pd.ExcelFile) -> str:
    good = ["zois", "trial", "balance", "zestawienie", "obroty", "tb"]
    bad  = ["mapp", "hyperion", "check", "pivot", "summary", "bs "]
    scores: dict[str, int] = {}
    for name in xls.sheet_names:
        nl = name.lower()
        s = sum(2 for k in good if k in nl) - sum(3 for k in bad if k in nl)
        try:
            df = xls.parse(name, header=None, nrows=3)
            s += min(df.shape[1], 8)
        except Exception:
            pass
        scores[name] = s
    return max(scores, key=lambda n: scores[n])


# ─── Trial Balance parser ─────────────────────────────────────────────────────

def parse_trial_balance(file_obj) -> dict:
    """
    Returns: {df, sheet_used, warnings, error}
    df columns: account_number, account_name, account_name2,
                bo_dt, bo_ct, obroty_dt, obroty_ct,
                obroty_n_dt, obroty_n_ct, saldo_dt, saldo_ct, persaldo, bs_mapp
    """
    warnings: list[str] = []
    try:
        xls = pd.ExcelFile(file_obj, engine="openpyxl")
    except Exception as e:
        return {"error": str(e)}

    sheet = _select_tb_sheet(xls)
    try:
        raw = xls.parse(sheet, header=None, dtype=str)
    except Exception as e:
        return {"error": f"Cannot read sheet '{sheet}': {e}"}

    hrow = _find_header_row(raw)
    df = xls.parse(sheet, header=hrow, dtype=str)
    df.columns = [_canonical(str(c)) for c in df.columns]
    df.dropna(how="all", inplace=True)
    df.reset_index(drop=True, inplace=True)

    num_cols = ["bo_dt","bo_ct","obroty_dt","obroty_ct",
                "obroty_n_dt","obroty_n_ct","saldo_dt","saldo_ct","persaldo"]
    for c in num_cols:
        if c not in df.columns:
            df[c] = 0.0
        else:
            df[c] = df[c].apply(_to_float)

    if "account_number" not in df.columns:
        df.rename(columns={df.columns[0]: "account_number"}, inplace=True)
    if "account_name" not in df.columns:
        df["account_name"] = df.get("account_name2", df["account_number"].astype(str))
    if "account_name2" not in df.columns:
        df["account_name2"] = ""
    if "bs_mapp" not in df.columns:
        df["bs_mapp"] = ""

    df["account_number"] = df["account_number"].astype(str).str.strip()
    df["account_name"]   = df["account_name"].astype(str).str.strip()
    df["bs_mapp"]        = df["bs_mapp"].astype(str).str.strip()

    df = df[
        df["account_number"].notna() &
        (df["account_number"] != "") &
        (df["account_number"] != "nan")
    ]
    df.reset_index(drop=True, inplace=True)

    if df.empty:
        warnings.append("Parsed DataFrame is empty after cleaning.")

    return {"df": df, "sheet_used": sheet, "warnings": warnings}


# ─── Complete group→side mapping (enumerated, no heuristics) ─────────────────
#
# Every group that appears in Mapping.xlsx BS Mapp column is listed here
# with its correct side. This eliminates all heuristic/prefix inference.

_GROUP_SIDE_MAP: dict[str, str] = {
    # ASSETS (A)
    "CASH - Bank Accounts":                              "A",
    "CASH - Restricted":                                 "A",
    "AR_TRADE - Current Invoices 0 to 30 Days":          "A",
    "A/R - Allowance (Bad Debt / Aging)":                "A",
    "I/C Accounts Receivable, Trade":                    "A",
    "I/C Accounts Receivable, Non-Trade":                "A",
    "INVENTORY - Raw Materials":                         "A",
    "INVENTORY - Work-in-Process":                       "A",
    "INVENTORY - Finished Goods-In Stock":               "A",
    "INVENTORY - Finished Goods-Offsite":                "A",
    "INVENTORY - Raw Material Reserves":                 "A",
    "INVENTORY - Finished Goods Reserves":               "A",
    "PREPAIDS - Property Taxes":                         "A",
    "PREPAIDS - Insurance":                              "A",
    "PREPAIDS - Other":                                  "A",
    "Gross Property Plant and Equipment":                "A",
    "PPE_GROSS.LAND - Additions":                        "A",
    "PPE_GROSS.MACHINERY - Beginning Balance":           "A",
    "PPE_GROSS.F_F - Beginning Balance":                 "A",
    "PPE_GROSS.LEASES_IMPR - Beginning Balance":         "A",
    "PPE_GROSS.AUTO - Beginning Balance":                "A",
    "PPE_GROSS.CONSTRUCT - Additions":                   "A",
    "PPE_GROSS.HW - Beginning Balance":                  "A",
    "PPE_GROSS.SW - Beginning Balance":                  "A",
    "Accumulated Depreciation - PPE":                    "A",
    "PPE_ACCUM.MACHINERY - Beginning Balance":           "A",
    "PPE_ACCUM.F_F - Asset Class Correction - Transfer In": "A",
    "PPE_ACCUM.LEASES_IMPR - Beginning Balance":         "A",
    "PPE_ACCUM.AUTO - Beginning Balance":                "A",
    "PPE_ACCUM.HW - Beginning Balance":                  "A",
    "PPE_ACCUM.SW - Beginning Balance":                  "A",
    "Gross Property Plant and Equipment - Right of Use": "A",
    "PPE_GROSS_Right-of-use_Building - Beginning Balance": "A",
    "PPE_GROSS_Right-of-use_M&E - Beginning Balance":    "A",
    "PPE_GROSS_Right-of-use_Auto - Beginning Balance":   "A",
    "Accumulated Depreciation - PPE - Right of Use":     "A",
    "PPE_ACCUM_Right-of-use_Building - Beginning Balance": "A",
    "PPE_ACCUM_Right-of-use_M&E - Beginning Balance":    "A",
    "PPE_ACCUM_Right-of-use_Auto - Beginning Balance":   "A",
    "Deferred Tax Asset - Cost":                         "A",

    # LIABILITIES / EQUITY (P)
    "Accounts Payable - Trade":                          "P",
    "I/C Accounts Payable-Trade":                        "P",
    "I/C Accounts Payable-Non-Trade":                    "P",
    "ACCRUALS - Audit Reserve":                          "P",
    "ACCRUALS - Vacation Pay":                           "P",
    "ACCRUALS - Accrued Payroll":                        "P",
    "ACCRUALS - Accrued Bonus":                          "P",
    "ACCRUALS - Accrued Payables":                       "P",
    "ACCRUALS - Utilities":                              "P",
    "ACCRUALS - Other Short-Term Accruals (over 50K ea)": "P",
    "Taxes Payable":                                     "P",
    "Other Taxes Payable":                               "P",
    "VAT Taxes Payable":                                 "P",
    "Right-of-use Liability - Current":                  "P",
    "Right-of-use Liability - Long-term":                "P",
    "Capital Stock":                                     "P",
    "RET_EARN - Opening Retained Earnings":              "P",
    "LT_LIABILITIES - Post Employment":                  "P",

    # EXCLUDED from BS (X) — P&L / off-balance / control accounts
    "COS_MATERIAL - Intercompany":                       "X",
    "Material COGS - Substrate":                         "X",
    "Intercompany Sales":                                "X",
    "SALES - Other Sales":                               "X",
    "0":                                                 "X",
}


def _get_side(group: str) -> str:
    """Return side for a group. Unknown groups → X (safe default)."""
    return _GROUP_SIDE_MAP.get(group.strip(), "X")


# ─── Mapping file parsers ─────────────────────────────────────────────────────

def parse_mapping_file(file_obj) -> dict:
    """
    Auto-detects Format A or B.
    Returns: {mapping: dict[acc -> {side, group}], sheet_used, error}
    """
    try:
        xls = pd.ExcelFile(file_obj, engine="openpyxl")
    except Exception as e:
        return {"error": str(e), "mapping": {}}

    for name in xls.sheet_names:
        if name.strip().lower() == "mapp":
            return _parse_format_a(xls, name)

    return _parse_format_b(xls)


def _parse_format_a(xls: pd.ExcelFile, sheet: str) -> dict:
    """Format A: Mapp sheet — col 0 = A/P/X/group, col 2 = account."""
    try:
        df = xls.parse(sheet, header=None, dtype=str)
    except Exception as e:
        return {"error": str(e), "mapping": {}}

    mapping: dict[str, dict] = {}
    current_group = ""

    for _, row in df.iterrows():
        col0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        col2 = str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else ""

        if col0 in ("A", "P", "X"):
            if col2 and col2 != "nan":
                mapping[col2] = {"side": col0, "group": current_group}
        elif col0 and col0 != "nan":
            current_group = col0

    return {"mapping": mapping, "sheet_used": sheet}


def _parse_format_b(xls: pd.ExcelFile) -> dict:
    """
    Format B: ZOiS-style file with Numer + BS Mapp columns.
    'x' (lowercase) → excluded from BS.
    All other values → group label; side from _GROUP_SIDE_MAP.
    """
    sheet = None
    for name in xls.sheet_names:
        try:
            df_test = xls.parse(name, nrows=5, dtype=str)
            cols = [str(c).strip().lower() for c in df_test.columns]
            if any("numer" in c or "account" in c for c in cols):
                sheet = name
                break
        except Exception:
            pass
    if sheet is None:
        sheet = xls.sheet_names[0]

    try:
        df = xls.parse(sheet, dtype=str)
    except Exception as e:
        return {"error": str(e), "mapping": {}}

    df.columns = [_canonical(str(c)) for c in df.columns]

    if "account_number" not in df.columns:
        return {"error": "No 'Numer'/account_number column found.", "mapping": {}}
    if "bs_mapp" not in df.columns:
        return {"error": "No 'BS Mapp' column found.", "mapping": {}}

    mapping: dict[str, dict] = {}
    for _, row in df.iterrows():
        acc   = str(row.get("account_number", "")).strip()
        group = str(row.get("bs_mapp", "")).strip()

        if not acc or acc == "nan":
            continue
        if not group or group == "nan":
            continue

        # Lowercase 'x' = excluded from BS
        if group.lower() == "x":
            mapping[acc] = {"side": "X", "group": "X"}
            continue

        side = _get_side(group)
        mapping[acc] = {"side": side, "group": group}

    return {"mapping": mapping, "sheet_used": sheet}


# ─── Default mapping loader ───────────────────────────────────────────────────

_DEFAULT_MAPPING_PATH = Path(__file__).parent.parent / "data" / "default_mapping.xlsx"


def load_default_mapping() -> dict:
    """Load the built-in default mapping from data/default_mapping.xlsx."""
    if not _DEFAULT_MAPPING_PATH.exists():
        return {"error": "Default mapping file not found in data/", "mapping": {}}
    try:
        with open(_DEFAULT_MAPPING_PATH, "rb") as f:
            return parse_mapping_file(f)
    except Exception as e:
        return {"error": str(e), "mapping": {}}
