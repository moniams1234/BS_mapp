"""
balance_sheet.py
Builds structured balance sheet from Mapp DataFrame.
Only uses accounts with side A or P (mapped rows only).
X accounts are excluded from BS.

BS ordering:
  Groups are presented in the order they appear in the Mapp file
  (which mirrors the BS reference file).
  Any group present in results but missing from the reference order
  is appended at the end.
"""
from __future__ import annotations
from pathlib import Path

import pandas as pd


# Canonical group order extracted from the Mapp sheet (col 0 non-A/P rows)
# This is derived from the Mapp file used as input and also cached here
# as fallback from data/bs_order.xlsx.

_BS_ORDER_FILE = Path(__file__).parent.parent / "data" / "bs_order.xlsx"


def load_bs_group_order() -> list[str]:
    """
    Load the canonical BS group order from bs_order.xlsx (sheet BS, col B)
    and from the Mapp sheet group headers.
    Returns an ordered list of group name strings.
    """
    order: list[str] = []
    seen: set[str] = set()

    # Try bs_order.xlsx first (col B = index 1)
    try:
        df = pd.read_excel(_BS_ORDER_FILE, sheet_name="BS", header=None)
        for val in df.iloc[:, 1]:
            s = str(val).strip() if pd.notna(val) else ""
            if s and s != "nan" and s not in seen:
                seen.add(s)
                order.append(s)
    except Exception:
        pass  # fallback below

    return order


def _sort_groups(df: pd.DataFrame, ref_order: list[str]) -> pd.DataFrame:
    """
    Sort a DataFrame with a 'group' column according to ref_order.
    Groups not in ref_order are appended at the end (alphabetically).
    """
    if df.empty:
        return df

    order_map = {g: i for i, g in enumerate(ref_order)}
    n = len(ref_order)

    df = df.copy()
    df["_sort_key"] = df["group"].apply(lambda g: order_map.get(g, n))
    df = df.sort_values("_sort_key").drop(columns=["_sort_key"])
    df.reset_index(drop=True, inplace=True)
    return df


def build_balance_sheet(mapp_df: pd.DataFrame, ref_order: list[str] | None = None) -> dict:
    """
    Returns dict:
      assets_by_group       pd.DataFrame  [group, amount]
      liabilities_by_group  pd.DataFrame  [group, amount]
      total_assets          float
      total_liabilities     float
      difference            float
    """
    if mapp_df is None or mapp_df.empty:
        return _empty()

    if ref_order is None:
        ref_order = load_bs_group_order()

    mapped = mapp_df[mapp_df["mapping_status"] == "mapped"]

    assets_df = mapped[mapped["side"] == "A"]
    liab_df   = mapped[mapped["side"] == "P"]

    assets_by_group = (
        assets_df.groupby("group")["persaldo"]
        .sum()
        .reset_index()
        .rename(columns={"persaldo": "amount"})
    )
    liab_by_group = (
        liab_df.groupby("group")["persaldo"]
        .sum()
        .reset_index()
        .rename(columns={"persaldo": "amount"})
    )

    # Apply reference ordering instead of sorting by value
    assets_by_group = _sort_groups(assets_by_group, ref_order)
    liab_by_group   = _sort_groups(liab_by_group, ref_order)

    total_assets = float(assets_by_group["amount"].sum())
    total_liab   = float(liab_by_group["amount"].sum())
    difference   = total_assets - total_liab

    return {
        "assets_by_group":      assets_by_group,
        "liabilities_by_group": liab_by_group,
        "total_assets":         total_assets,
        "total_liabilities":    total_liab,
        "difference":           difference,
    }


def _empty() -> dict:
    cols = ["group", "amount"]
    return {
        "assets_by_group":      pd.DataFrame(columns=cols),
        "liabilities_by_group": pd.DataFrame(columns=cols),
        "total_assets":         0.0,
        "total_liabilities":    0.0,
        "difference":           0.0,
    }
