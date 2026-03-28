"""
balance_sheet.py

Builds structured balance sheet from the mapped DataFrame.
Only uses accounts with side A (assets) or P (liabilities/equity).
X and excluded accounts are not shown in BS.

BS GROUP ORDERING:
  Groups are ordered according to the reference file data/bs_order.xlsx
  (sheet 'BS', column B). This is the canonical Hyperion/CCL BS order.
  Groups not found in the reference list are appended at the end.
  NEVER sort by amount.
"""
from __future__ import annotations
from pathlib import Path

import pandas as pd


_BS_ORDER_FILE = Path(__file__).parent.parent / "data" / "bs_order.xlsx"


def load_bs_group_order() -> list[str]:
    """
    Load the canonical BS group order from bs_order.xlsx,
    sheet 'BS', column B (index 1).
    Returns an ordered list of group name strings.
    """
    order: list[str] = []
    seen: set[str] = set()
    try:
        df = pd.read_excel(_BS_ORDER_FILE, sheet_name="BS", header=None)
        for val in df.iloc[:, 1]:
            s = str(val).strip() if pd.notna(val) else ""
            if s and s != "nan" and s not in seen:
                seen.add(s)
                order.append(s)
    except Exception:
        pass
    return order


def apply_bs_order(df: pd.DataFrame, ref_order: list[str]) -> pd.DataFrame:
    """
    Sort DataFrame by 'group' column according to ref_order.
    Groups not in ref_order are appended at the end in alphabetical order.
    """
    if df.empty:
        return df

    order_map = {g: i for i, g in enumerate(ref_order)}
    n = len(ref_order)

    df = df.copy()
    df["_sort_key"] = df["group"].apply(lambda g: order_map.get(g, n))
    df = df.sort_values(["_sort_key", "group"]).drop(columns=["_sort_key"])
    df.reset_index(drop=True, inplace=True)
    return df


def build_balance_sheet(mapp_df: pd.DataFrame, ref_order: list[str] | None = None) -> dict:
    """
    Build balance sheet from mapped DataFrame.

    Returns dict:
      assets_by_group       pd.DataFrame  [group, amount]  — ordered by ref_order
      liabilities_by_group  pd.DataFrame  [group, amount]  — ordered by ref_order
      total_assets          float
      total_liabilities     float
      difference            float
    """
    if mapp_df is None or mapp_df.empty:
        return _empty()

    if ref_order is None:
        ref_order = load_bs_group_order()

    mapped = mapp_df[mapp_df["mapping_status"] == "mapped"]

    assets_by_group = (
        mapped[mapped["side"] == "A"]
        .groupby("group")["persaldo"]
        .sum()
        .reset_index()
        .rename(columns={"persaldo": "amount"})
    )

    liab_by_group = (
        mapped[mapped["side"] == "P"]
        .groupby("group")["persaldo"]
        .sum()
        .reset_index()
        .rename(columns={"persaldo": "amount"})
    )

    # Apply reference ordering — never sort by amount
    assets_by_group = apply_bs_order(assets_by_group, ref_order)
    liab_by_group   = apply_bs_order(liab_by_group, ref_order)

    total_assets = float(assets_by_group["amount"].sum())
    total_liab   = float(liab_by_group["amount"].sum())

    return {
        "assets_by_group":      assets_by_group,
        "liabilities_by_group": liab_by_group,
        "total_assets":         total_assets,
        "total_liabilities":    total_liab,
        "difference":           total_assets - total_liab,
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
