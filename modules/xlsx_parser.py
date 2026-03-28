"""
xlsx_parser.py

Parses two types of XLSX files:
  1. Trial balance (ZOiS) - returns cleaned DataFrame
  2. Mapping file - returns dict {account_number -> {side, group}}

MAPPING FILE FORMATS SUPPORTED:

  Format A – "Mapp sheet" format (from Mapp_BS_Hyperion file):
    Sheet named 'Mapp':
      Col 0: 'A' | 'P' | 'X' | group-header-name
      Col 2: account number (when col 0 is A/P/X)

  Format B – "ZOiS + BS Mapp column" format (from Mapping.xlsx):
    Sheet: any (heuristically detected)
    Columns: Numer | ... | BS Mapp
    The BS Mapp column contains the group label; side is derived from
    the Mapp sheet of the reference Mapp_BS file.
    Since the BS Mapp values in Mapping.xlsx correspond to group labels,
    and side is already encoded in the reference Mapp sheet, we use a
    lookup: group label -> side (A/P/X) derived from the Mapp sheet groups.

    For Mapping.xlsx we derive side from the canonical group ordering:
      - Groups whose accounts had side A in the Mapp reference → A
      - Groups whose accounts had side P → P
      - 'x' value → X (excluded)

TRIAL BALANCE FORMAT:
  Standard ZOiS: Numer, Nazwa 2, Nazwa, BO Dt, BO Ct, Obroty Dt, Obroty Ct,
                 Obroty n. Dt, Obroty n. Ct, Saldo Dt, Saldo Ct, Persaldo
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Optional

import pandas as pd


# ─── column name normalization ───────────────────────────────────────────────

_COL_MAP: dict[str, list[str]] = {
    "account_number": ["numer", "konto", "account", "account_number", "account number",
                       "nr konta", "numer konta", "account no"],
    "account_name2": ["nazwa 2", "name2", "label", "nazwa2"],
    "account_name":  ["nazwa", "name", "account_name", "account name", "opis", "description"],
    "bo_dt":         ["bo dt", "bo dt ", "opening debit"],
    "bo_ct":         ["bo ct", "bo ct ", "opening credit"],
    "obroty_dt":     ["obroty dt", "obroty dt ", "debit", "wn", "dt", "turnover debit"],
    "obroty_ct":     ["obroty ct", "obroty ct ", "credit", "ma", "ct", "turnover credit"],
    "obroty_n_dt":   ["obroty n. dt", "obroty n. dt "],
    "obroty_n_ct":   ["obroty n. ct", "obroty n. ct "],
    "saldo_dt":      ["saldo dt", "saldo dt ", "closing debit", "balance debit"],
    "saldo_ct":      ["saldo ct", "saldo ct ", "closing credit", "balance credit"],
    "persaldo":      ["persaldo", "saldo", "balance", "net balance", "net saldo"],
    "bs_mapp":       ["bs mapp", "bs_mapp", "mapp", "bs mapp "],
}


def _norm(s: object) -> str:
    return str(s).strip().lower() if not pd.isna(s) else ""


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
            if any(_norm(cell) in aliases for aliases in _COL_MAP.values())
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


# ─── public API – trial balance ───────────────────────────────────────────────

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
        if "account_name2" in df.columns:
            df["account_name"] = df["account_name2"]
        else:
            df["account_name"] = df["account_number"].astype(str)
    if "account_name2" not in df.columns:
        df["account_name2"] = ""
    if "bs_mapp" not in df.columns:
        df["bs_mapp"] = ""

    df["account_number"] = df["account_number"].astype(str).str.strip()
    df["account_name"]   = df["account_name"].astype(str).str.strip()
    df["bs_mapp"]        = df["bs_mapp"].astype(str).str.strip()

    df = df[df["account_number"].notna() & (df["account_number"] != "") & (df["account_number"] != "nan")]
    df.reset_index(drop=True, inplace=True)

    if df.empty:
        warnings.append("Parsed DataFrame is empty after cleaning.")

    return {"df": df, "sheet_used": sheet, "warnings": warnings}


# ─── public API – mapping file ────────────────────────────────────────────────

def parse_mapping_file(file_obj) -> dict:
    """
    Parse mapping XLSX. Supports two formats:

    Format A: has a 'Mapp' sheet with A/P/X in col 0, account in col 2.
    Format B: has 'Numer' and 'BS Mapp' columns (like Mapping.xlsx).
              Side is inferred from the BS Mapp group label using the
              reference side map embedded below.

    Returns: {mapping: dict[acc -> {side, group}], sheet_used, error}
    """
    try:
        xls = pd.ExcelFile(file_obj, engine="openpyxl")
    except Exception as e:
        return {"error": str(e), "mapping": {}}

    # Try Format A: 'Mapp' sheet
    mapp_sheet = None
    for name in xls.sheet_names:
        if name.strip().lower() == "mapp":
            mapp_sheet = name
            break

    if mapp_sheet:
        return _parse_format_a(xls, mapp_sheet)
    else:
        # Format B: ZOiS-style file with BS Mapp column
        return _parse_format_b(xls)


def _parse_format_a(xls: pd.ExcelFile, sheet: str) -> dict:
    """Format A: Mapp sheet with A/P/X col0, account col2."""
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
            acc = col2
            if acc and acc != "nan":
                mapping[acc] = {"side": col0, "group": current_group}
        elif col0 and col0 != "nan":
            current_group = col0

    return {"mapping": mapping, "sheet_used": sheet}


# Side lookup for Format B: inferred from the canonical Mapp file analysis
# Groups that are ASSETS (side A) in the reference Mapp sheet:
_ASSET_GROUPS: frozenset = frozenset({
    "CASH - Bank Accounts", "CASH - Restricted",
    "AR_TRADE", "A/R - Allowance (Bad Debt / Aging)",
    "I/C Accounts Receivable, Trade", "I/C Accounts Receivable, Non-Trade",
    "INVENTORY - Raw Materials", "INVENTORY - Work-in-Process",
    "INVENTORY - Finished Goods-In Stock", "INVENTORY - Finished Goods-Offsite",
    "INVENTORY - Raw Material Reserves", "INVENTORY - Finished Goods Reserves",
    "PREPAIDS - Property Tax", "PREPAIDS - Insurance", "PREPAIDS - Other",
    "Gross Property Plant and Equipment",
    "PPE_GROSS.LAND", "PPE_GROSS.MACHINERY", "PPE_GROSS.F_F",
    "PPE_GROSS.LEASES_IMPR", "PPE_GROSS.AUTO", "PPE_GROSS.CONSTRUCT", "PPE_GROSS.SW",
    "Accumulated Depreciation - PPE",
    "PPE_ACCUM.MACHINERY", "PPE_ACCUM.LEASES_IMPR", "PPE_ACCUM.AUTO", "PPE_ACCUM.SW",
    "Gross Property Plant and Equipment - Right of Use",
    "PPE_GROSS_Right-of-use_Land", "PPE_GROSS_Right-of-use_M&E", "PPE_GROSS_Right-of-use_Auto",
    "Accumulated Depreciation - PPE - Right of Use",
    "PPE_ACCUM_Right-of-use_Building", "PPE_ACCUM_Right-of-use_M&E", "PPE_ACCUM_Right-of-use_Auto",
    "Deferred Tax Asset",
})

_LIABILITY_GROUPS: frozenset = frozenset({
    "Accounts Payable - Trade", "I/C Accounts Payable-Trade",
    "ACCRUALS - Audit Reserve", "ACCRUALS - Vacation Pay",
    "ACCRUALS - Accrued Payroll", "ACCRUALS - Accrued Bonus",
    "ACCRUALS - Accrued payables", "ACCRUALS - Utilities",
    "ACCRUALS - Other Short-Term Accruals (under 50K ea",
    "I/C Accounts Payable-Non-Trade",
    "Taxes Payable", "Other Taxes Payable", "VAT Taxes Payable",
    "Right-of-use Liability - Current", "Right-of-use Liability - Long-term",
    "Capital Stock", "RET_EARN - Opening Retained Earnings",
})


def _infer_side(group: str) -> str:
    """Infer A/P/X from group name using reference lookup."""
    if group in _ASSET_GROUPS:
        return "A"
    if group in _LIABILITY_GROUPS:
        return "P"
    # fallback: try heuristics
    gl = group.lower()
    if any(k in gl for k in ["cash", "receivable", "inventory", "prepaid", "ppe", "asset", "deferred tax asset"]):
        return "A"
    if any(k in gl for k in ["payable", "accrual", "tax payab", "liability", "capital", "earn", "equity", "debt"]):
        return "P"
    return "A"  # default to A if unknown


def _parse_format_b(xls: pd.ExcelFile) -> dict:
    """
    Format B: ZOiS file with Numer and BS Mapp columns.
    account_number = Numer, group = BS Mapp, side = inferred.
    """
    # Find the best sheet
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

    # Normalize column names
    df.columns = [_canonical(str(c)) for c in df.columns]

    if "account_number" not in df.columns:
        return {"error": "No 'Numer' or account_number column found in mapping file.", "mapping": {}}
    if "bs_mapp" not in df.columns:
        return {"error": "No 'BS Mapp' column found in mapping file.", "mapping": {}}

    mapping: dict[str, dict] = {}
    for _, row in df.iterrows():
        acc   = str(row.get("account_number", "")).strip()
        group = str(row.get("bs_mapp", "")).strip()

        if not acc or acc == "nan":
            continue
        if not group or group == "nan":
            continue

        # 'x' means excluded
        if group.lower() == "x":
            mapping[acc] = {"side": "X", "group": "X"}
            continue

        side = _infer_side(group)
        mapping[acc] = {"side": side, "group": group}

    return {"mapping": mapping, "sheet_used": sheet}


# ─── default mapping loader ───────────────────────────────────────────────────

_DEFAULT_MAPPING_PATH = Path(__file__).parent.parent / "data" / "default_mapping.xlsx"


def load_default_mapping() -> dict:
    """Load the built-in default mapping from data/default_mapping.xlsx."""
    if not _DEFAULT_MAPPING_PATH.exists():
        return {"error": "Default mapping file not found.", "mapping": {}}
    try:
        with open(_DEFAULT_MAPPING_PATH, "rb") as f:
            return parse_mapping_file(f)
    except Exception as e:
        return {"error": str(e), "mapping": {}}
