"""
balance_sheet.py
Builds structured balance sheet DataFrames for display.
"""
from __future__ import annotations

import pandas as pd
import numpy as np


# Groups that are treated as equity (to split from liabilities)
EQUITY_GROUPS = {
    "Equity - Share Capital",
    "Equity - Retained Earnings",
    "Equity - Reserves",
    "Equity - Current Year Profit",
    "Equity - Other",
}


def build_balance_sheet_table(mapp_df: pd.DataFrame) -> dict:
    """
    Returns:
      assets_detail, liabilities_detail, equity_detail, pl_detail
      Each is a DataFrame: group | amount
    Plus scalar summaries.
    """
    if mapp_df is None or mapp_df.empty:
        return {}

    assets = mapp_df[mapp_df["side"] == "A"].groupby("group")["persaldo"].sum().reset_index()
    assets.columns = ["Line Item", "Amount"]
    assets = assets[assets["Amount"] != 0].sort_values("Amount", ascending=False)

    all_p = mapp_df[mapp_df["side"] == "P"].groupby("group")["persaldo"].sum().reset_index()
    all_p.columns = ["Line Item", "Amount"]

    equity_detail = all_p[all_p["Line Item"].isin(EQUITY_GROUPS)].copy()
    liabilities_detail = all_p[~all_p["Line Item"].isin(EQUITY_GROUPS)].copy()
    liabilities_detail = liabilities_detail[liabilities_detail["Amount"] != 0].sort_values("Amount", ascending=False)
    equity_detail = equity_detail[equity_detail["Amount"] != 0].sort_values("Amount", ascending=False)

    pl_detail = mapp_df[mapp_df["side"] == "R"].groupby("group")["persaldo"].sum().reset_index()
    pl_detail.columns = ["Line Item", "Amount"]
    pl_detail = pl_detail[pl_detail["Amount"] != 0].sort_values("Amount", ascending=False)

    total_assets = assets["Amount"].sum()
    total_liab = liabilities_detail["Amount"].sum()
    total_equity = equity_detail["Amount"].sum()
    total_pl = pl_detail["Amount"].sum()
    total_liab_equity = total_liab + total_equity
    difference = total_assets - total_liab_equity

    return {
        "assets_detail": assets,
        "liabilities_detail": liabilities_detail,
        "equity_detail": equity_detail,
        "pl_detail": pl_detail,
        "total_assets": total_assets,
        "total_liabilities": total_liab,
        "total_equity": total_equity,
        "total_pl": total_pl,
        "total_liab_equity": total_liab_equity,
        "difference": difference,
    }
