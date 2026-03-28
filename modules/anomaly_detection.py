"""
anomaly_detection.py
Builds red flag list from parsed and mapped data.
Each flag: {type: error|warning|success, category: str, message: str}
"""
from __future__ import annotations

from typing import Any

import numpy as np
import pandas as pd


def build_red_flags(
    df: pd.DataFrame,
    mapp_df: pd.DataFrame,
    bs: dict,
    warnings: list[str],
) -> list[dict[str, str]]:
    flags: list[dict[str, str]] = []

    def flag(ftype: str, category: str, message: str):
        flags.append({"type": ftype, "category": category, "message": message})

    # --- Source data checks ---
    if df is None or df.empty:
        flag("error", "Data", "No data parsed from file.")
        return flags

    # Required columns
    for col in ["account_number", "saldo_dt", "saldo_ct"]:
        if col not in df.columns:
            flag("error", "Missing Column", f"Required column '{col}' not found in trial balance.")

    # Numeric issues
    for col in ["obroty_dt", "obroty_ct", "saldo_dt", "saldo_ct"]:
        if col in df.columns and df[col].isna().any():
            count = df[col].isna().sum()
            flag("warning", "Data Quality", f"Column '{col}' has {count} non-numeric / missing values.")

    # Parser warnings passthrough
    for w in warnings:
        flag("warning", "Parser", w)

    if mapp_df is None or mapp_df.empty:
        flag("error", "Mapping", "Mapping produced empty result.")
        return flags

    # --- Mapping checks ---
    total = len(mapp_df)
    unmapped = mapp_df[mapp_df["mapping_status"] == "heuristic"]
    n_unmapped = len(unmapped)
    pct_unmapped = n_unmapped / total * 100 if total else 0

    if n_unmapped == 0:
        flag("success", "Mapping", f"All {total} accounts mapped successfully.")
    elif pct_unmapped < 10:
        flag("warning", "Mapping", f"{n_unmapped} accounts ({pct_unmapped:.1f}%) used heuristic mapping.")
    elif pct_unmapped < 30:
        flag("warning", "Mapping", f"{n_unmapped} accounts ({pct_unmapped:.1f}%) unmapped – review mapping table.")
    else:
        flag("error", "Mapping", f"{n_unmapped} accounts ({pct_unmapped:.1f}%) could not be mapped. Check account prefixes.")

    # Duplicate accounts
    dupes = mapp_df[mapp_df.duplicated(subset=["account_number"], keep=False)]
    if not dupes.empty:
        flag("warning", "Data Quality", f"{len(dupes)} duplicate account numbers detected.")

    # Empty groups
    empty_groups = mapp_df.groupby("group")["persaldo"].sum()
    zero_groups = empty_groups[empty_groups == 0].index.tolist()
    if zero_groups:
        flag("warning", "Mapping", f"{len(zero_groups)} groups have zero net balance (may be correct).")

    # Large unidentified bucket
    unidentified = mapp_df[mapp_df["group"] == "UNIDENTIFIED"]
    if not unidentified.empty:
        flag("error", "Mapping", f"{len(unidentified)} accounts in UNIDENTIFIED group – check account number format.")

    # --- Balance sheet checks ---
    if bs:
        diff = bs.get("difference", 0)
        total_assets = bs.get("total_assets", 0)
        total_liabilities = bs.get("total_liabilities", 0)

        if abs(diff) < 1.0:
            flag("success", "Balance", "Balance sheet balances (Assets = Liabilities + Equity).")
        elif abs(diff) < total_assets * 0.01:
            flag("warning", "Balance", f"Small imbalance detected: {diff:,.2f}. May be rounding.")
        else:
            flag("error", "Balance", f"Balance sheet does NOT balance. Difference: {diff:,.2f}")

        if total_assets == 0:
            flag("error", "Balance", "Total assets = 0. Check trial balance data or mapping.")

        if total_liabilities < 0:
            flag("warning", "Balance", f"Total liabilities is negative ({total_liabilities:,.2f}). Check P-side signs.")

    # --- P&L checks ---
    r_accounts = mapp_df[mapp_df["side"] == "R"]
    if r_accounts.empty:
        flag("warning", "P&L", "No P&L accounts found. All accounts mapped to balance sheet sides.")

    # Very large debit / credit imbalance per account (potential data error)
    mapp_df_copy = mapp_df.copy()
    mapp_df_copy["db_cr_ratio"] = mapp_df_copy.apply(
        lambda r: abs(r["debit"] / r["credit"]) if r["credit"] != 0 else None, axis=1
    )
    weird = mapp_df_copy[mapp_df_copy["db_cr_ratio"].notna() & (mapp_df_copy["db_cr_ratio"] > 100)]
    if not weird.empty:
        flag("warning", "Data Quality", f"{len(weird)} accounts have debit/credit ratio > 100x. Verify data.")

    return flags
