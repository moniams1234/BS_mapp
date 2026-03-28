"""
mapping_engine.py

Business rules:
  - EXACT MATCH on full account_number — no prefix, no heuristics
  - Accounts starting with '9' excluded UNLESS in _9XX_WHITELIST
  - Group 'X' accounts excluded from balance sheet
  - heuristic_mapped always = 0

Persaldo sign convention:
  A (assets):  persaldo = saldo_dt - saldo_ct  (positive = debit balance)
  P (liab/eq): persaldo = saldo_ct - saldo_dt  (positive = credit balance)
  X / other:   persaldo = saldo_dt - saldo_ct  (reference only)
"""
from __future__ import annotations

import pandas as pd


# 9xx accounts that are valid BS accounts (PPE / off-balance positions)
_9XX_WHITELIST: frozenset = frozenset({
    "902-03",
    "907", "907-01", "907-02", "907-03",
    "910", "910-01", "910-02", "910-03",
    "912", "912-01",
    "972", "972-01",
    "973", "973-03",
})


def _is_excluded_9xx(acc: str) -> bool:
    """True if account starts with '9' and is NOT in the whitelist."""
    return acc.startswith("9") and acc not in _9XX_WHITELIST


def run_mapping(df: pd.DataFrame, mapping: dict) -> pd.DataFrame:
    """
    Exact-match mapping of trial balance to groups/sides.

    Parameters
    ----------
    df      : cleaned trial balance DataFrame
    mapping : {account_number -> {side, group}}

    Returns DataFrame with columns:
      account_number, account_name, group, side,
      debit, credit, saldo_dt, saldo_ct,
      persaldo, persaldo_group, mapping_status
    """
    records = []

    for _, row in df.iterrows():
        acc = str(row.get("account_number", "")).strip()

        # Auto-exclude 9xx (except whitelist)
        if _is_excluded_9xx(acc):
            records.append(_make_record(row, acc, "excluded", "", "excluded"))
            continue

        match = mapping.get(acc)
        if match:
            records.append(_make_record(row, acc, match["side"], match["group"], "mapped"))
        else:
            records.append(_make_record(row, acc, "unmapped", "", "unmapped"))

    result = pd.DataFrame(records)

    # persaldo_group = group-level sum (for mapped A/P rows only)
    mapped_mask = result["mapping_status"] == "mapped"
    if mapped_mask.any():
        group_sums = (
            result[mapped_mask]
            .groupby("group")["persaldo"]
            .sum()
            .rename("persaldo_group")
        )
        result = result.merge(group_sums, on="group", how="left")
    else:
        result["persaldo_group"] = 0.0

    result["persaldo_group"] = result["persaldo_group"].fillna(0.0)
    return result


def _make_record(row: pd.Series, acc: str, side: str, group: str, status: str) -> dict:
    debit    = float(row.get("obroty_dt", 0) or 0)
    credit   = float(row.get("obroty_ct", 0) or 0)
    saldo_dt = float(row.get("saldo_dt",  0) or 0)
    saldo_ct = float(row.get("saldo_ct",  0) or 0)
    name     = str(row.get("account_name", "")).strip()

    persaldo = (saldo_ct - saldo_dt) if side == "P" else (saldo_dt - saldo_ct)

    return {
        "account_number": acc,
        "account_name":   name,
        "group":          group,
        "side":           side,
        "debit":          debit,
        "credit":         credit,
        "saldo_dt":       saldo_dt,
        "saldo_ct":       saldo_ct,
        "persaldo":       persaldo,
        "mapping_status": status,
    }


def compute_kpis(mapp_df: pd.DataFrame) -> dict:
    """Compute summary KPIs from the mapped DataFrame."""
    if mapp_df is None or mapp_df.empty:
        return {}

    total    = len(mapp_df)
    mapped   = int((mapp_df["mapping_status"] == "mapped").sum())
    unmapped = int((mapp_df["mapping_status"] == "unmapped").sum())
    excluded = int((mapp_df["mapping_status"] == "excluded").sum())

    assets_df = mapp_df[(mapp_df["side"] == "A") & (mapp_df["mapping_status"] == "mapped")]
    liab_df   = mapp_df[(mapp_df["side"] == "P") & (mapp_df["mapping_status"] == "mapped")]

    return {
        "total_accounts":    total,
        "mapped_accounts":   mapped,
        "heuristic_mapped":  0,          # always 0 — no heuristics used
        "unmapped_accounts": unmapped,
        "excluded_accounts": excluded,
        "total_assets":      float(assets_df["persaldo"].sum()),
        "total_liabilities": float(liab_df["persaldo"].sum()),
        "difference":        float(assets_df["persaldo"].sum() - liab_df["persaldo"].sum()),
    }
