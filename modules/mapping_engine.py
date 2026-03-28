"""
mapping_engine.py
Maps trial balance accounts to reporting groups.

Logic:
- Exact match on account_number first.
- Prefix match (longest prefix wins).
- Uses built-in sample_data/sample_mapping.json unless user supplies another.
- Assigns side: A / P / R (assets / liabilities+equity / P&L).
- Computes persaldo per account and per group.

Persaldo sign convention:
  Assets (A):  positive persaldo = debit balance (normal for assets)
  Liabilities (P): positive persaldo = credit balance (normal for liabilities)
    → stored as negative debit surplus = saldo_ct - saldo_dt
  P&L (R): positive = credit (revenue/income); negative = debit (expense)
    → stored as saldo_ct - saldo_dt
"""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd


_MAPPING_PATH = Path(__file__).parent.parent / "sample_data" / "sample_mapping.json"

# Fallback if file missing
_BUILTIN_ACCOUNTS: list[dict] = []


def _load_mapping(mapping_file: Optional[str] = None) -> list[dict]:
    path = Path(mapping_file) if mapping_file else _MAPPING_PATH
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("accounts", [])
    except Exception:
        return []


def _build_prefix_index(accounts: list[dict]) -> dict[str, dict]:
    """Return a dict of prefix -> account_def, sorted longest first."""
    return {a["prefix"]: a for a in accounts}


def _match_account(account_no: str, prefix_index: dict[str, dict]) -> Optional[dict]:
    """Find best matching account def for a given account number."""
    account_no = str(account_no).strip()
    # Exact match
    if account_no in prefix_index:
        return prefix_index[account_no]
    # Prefix match: try stripping sub-account suffixes
    # Also try numeric prefix (first 3 digits)
    for length in range(len(account_no), 0, -1):
        prefix = account_no[:length]
        if prefix in prefix_index:
            return prefix_index[prefix]
    # Try numeric-only prefix for accounts like "131-03"
    numeric_only = "".join(c for c in account_no if c.isdigit())
    for length in range(3, len(numeric_only) + 1):
        prefix = numeric_only[:length]
        if prefix in prefix_index:
            return prefix_index[prefix]
    return None


def run_mapping(df: pd.DataFrame, mapping_file: Optional[str] = None) -> pd.DataFrame:
    """
    Take cleaned trial balance DataFrame and return a Mapp DataFrame.

    Returns columns:
      side, account_number, account_name, group, debit, credit,
      saldo_dt, saldo_ct, persaldo, persaldo_group, mapping_status, bs_mapp
    """
    accounts = _load_mapping(mapping_file)
    prefix_index = _build_prefix_index(accounts)

    records = []
    for _, row in df.iterrows():
        acc_no = str(row.get("account_number", "")).strip()
        acc_name = str(row.get("account_name", "")).strip()

        debit = float(row.get("obroty_dt", 0) or 0)
        credit = float(row.get("obroty_ct", 0) or 0)
        saldo_dt = float(row.get("saldo_dt", 0) or 0)
        saldo_ct = float(row.get("saldo_ct", 0) or 0)
        persaldo_raw = float(row.get("persaldo", 0) or 0)
        bs_mapp = str(row.get("bs_mapp", "")).strip()

        matched = _match_account(acc_no, prefix_index)

        if matched:
            side = matched["side"]
            group = matched["group"]
            status = "mapped"
        else:
            # Fallback heuristic
            first3 = acc_no[:3] if len(acc_no) >= 3 else acc_no
            try:
                num = int(first3)
            except ValueError:
                num = -1

            if 0 <= num <= 99:
                side, group = "A", "Non-Current Assets"
            elif 100 <= num <= 199:
                side, group = "A", "Cash & Short-Term Financial Assets"
            elif 200 <= num <= 299:
                side, group = "A", "Receivables & Payables"
            elif 300 <= num <= 399:
                side, group = "A", "Inventories"
            elif 400 <= num <= 799:
                side, group = "R", "Operating Costs / Revenue"
            elif 800 <= num <= 899:
                side, group = "P", "Liabilities"
            elif 900 <= num <= 999:
                side, group = "P", "Equity"
            else:
                side, group = "A", "UNIDENTIFIED"
            status = "heuristic"

        # Compute persaldo by side convention
        if side == "A":
            persaldo = saldo_dt - saldo_ct  # positive = debit (normal for assets)
        else:
            # P and R: positive = credit balance (normal for liabilities/equity/revenue)
            persaldo = saldo_ct - saldo_dt

        records.append({
            "side": side,
            "account_number": acc_no,
            "account_name": acc_name,
            "group": group,
            "debit": debit,
            "credit": credit,
            "saldo_dt": saldo_dt,
            "saldo_ct": saldo_ct,
            "persaldo": persaldo,
            "mapping_status": status,
            "bs_mapp": bs_mapp,
        })

    mapp_df = pd.DataFrame(records)

    if mapp_df.empty:
        return mapp_df

    # Add group-level persaldo (sum of persaldo within same group)
    group_sums = mapp_df.groupby("group")["persaldo"].sum().rename("persaldo_group")
    mapp_df = mapp_df.merge(group_sums, on="group", how="left")

    return mapp_df


def compute_balance_sheet(mapp_df: pd.DataFrame) -> dict:
    """
    From Mapp DataFrame, compute balance sheet summary.
    Returns dict with assets_df, liabilities_df, total_assets, total_liabilities,
    total_equity, difference.
    """
    if mapp_df.empty:
        return {}

    assets = mapp_df[mapp_df["side"] == "A"].copy()
    liabilities = mapp_df[mapp_df["side"] == "P"].copy()
    pl = mapp_df[mapp_df["side"] == "R"].copy()

    # Group summaries
    assets_grouped = (
        assets.groupby("group")["persaldo"].sum().reset_index()
        .rename(columns={"persaldo": "amount"})
        .sort_values("amount", ascending=False)
    )
    liabilities_grouped = (
        liabilities.groupby("group")["persaldo"].sum().reset_index()
        .rename(columns={"persaldo": "amount"})
        .sort_values("amount", ascending=False)
    )
    pl_grouped = (
        pl.groupby("group")["persaldo"].sum().reset_index()
        .rename(columns={"persaldo": "amount"})
        .sort_values("amount", ascending=False)
    )

    total_assets = assets["persaldo"].sum()
    total_liabilities = liabilities[liabilities["group"] != "Equity - Share Capital"]["persaldo"].sum()
    total_equity = liabilities[liabilities["group"].str.contains("Equity", na=False)]["persaldo"].sum()
    total_pl = pl["persaldo"].sum()

    difference = total_assets - (total_liabilities + total_equity)

    return {
        "assets_df": assets_grouped,
        "liabilities_df": liabilities_grouped,
        "pl_df": pl_grouped,
        "total_assets": total_assets,
        "total_liabilities": total_liabilities,
        "total_equity": total_equity,
        "total_pl": total_pl,
        "difference": difference,
    }
