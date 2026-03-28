"""
charts.py
Builds Plotly figures for the financial dashboard.
"""
from __future__ import annotations

import plotly.graph_objects as go
import plotly.express as px
import pandas as pd
import numpy as np


_COLORS = {
    "asset": "#2563EB",
    "liability": "#DC2626",
    "equity": "#059669",
    "pl": "#D97706",
    "bg": "#0F172A",
    "card": "#1E293B",
    "text": "#F1F5F9",
    "grid": "#334155",
}

_LAYOUT_DEFAULTS = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(color=_COLORS["text"], family="Arial, sans-serif"),
    margin=dict(l=40, r=20, t=50, b=40),
)


def balance_sheet_waterfall(bs: dict) -> go.Figure:
    """Waterfall showing assets vs liabilities+equity."""
    total_assets = bs.get("total_assets", 0)
    total_liab = bs.get("total_liabilities", 0)
    total_equity = bs.get("total_equity", 0)
    total_pl = bs.get("total_pl", 0)

    labels = ["Total Assets", "Liabilities", "Equity", "P&L Net"]
    values = [total_assets, -total_liab, -total_equity, total_pl]
    colors = [_COLORS["asset"], _COLORS["liability"], _COLORS["equity"], _COLORS["pl"]]

    fig = go.Figure(go.Bar(
        x=labels,
        y=[abs(v) for v in values],
        marker_color=colors,
        text=[f"{v/1e6:.2f}M" if abs(v) >= 1e5 else f"{v:,.0f}" for v in values],
        textposition="auto",
    ))
    fig.update_layout(
        title="Balance Sheet Structure",
        **_LAYOUT_DEFAULTS,
        xaxis=dict(showgrid=False),
        yaxis=dict(showgrid=True, gridcolor=_COLORS["grid"]),
    )
    return fig


def assets_breakdown_pie(assets_df: pd.DataFrame) -> go.Figure:
    """Pie chart of asset groups."""
    if assets_df is None or assets_df.empty:
        return go.Figure()
    df = assets_df[assets_df["amount"] > 0].copy()
    fig = go.Figure(go.Pie(
        labels=df["group"],
        values=df["amount"],
        hole=0.4,
        marker=dict(colors=px.colors.qualitative.Set3),
        textinfo="label+percent",
    ))
    fig.update_layout(title="Asset Structure", **_LAYOUT_DEFAULTS)
    return fig


def liabilities_breakdown_pie(liab_df: pd.DataFrame) -> go.Figure:
    """Pie chart of liability/equity groups."""
    if liab_df is None or liab_df.empty:
        return go.Figure()
    df = liab_df[liab_df["amount"] > 0].copy()
    fig = go.Figure(go.Pie(
        labels=df["group"],
        values=df["amount"],
        hole=0.4,
        marker=dict(colors=px.colors.qualitative.Pastel),
        textinfo="label+percent",
    ))
    fig.update_layout(title="Liabilities & Equity Structure", **_LAYOUT_DEFAULTS)
    return fig


def mapp_group_bar(mapp_df: pd.DataFrame) -> go.Figure:
    """Grouped bar chart of persaldo per group, split by side."""
    if mapp_df is None or mapp_df.empty:
        return go.Figure()

    group_summary = (
        mapp_df.groupby(["group", "side"])["persaldo"].sum().reset_index()
    )

    fig = px.bar(
        group_summary,
        x="group",
        y="persaldo",
        color="side",
        color_discrete_map={"A": _COLORS["asset"], "P": _COLORS["liability"], "R": _COLORS["pl"]},
        title="Net Balance by Reporting Group",
        labels={"persaldo": "Net Balance", "group": "Group"},
    )
    fig.update_layout(
        **_LAYOUT_DEFAULTS,
        xaxis=dict(tickangle=-45, showgrid=False),
        yaxis=dict(showgrid=True, gridcolor=_COLORS["grid"]),
        legend_title="Side",
    )
    return fig


def top_accounts_bar(mapp_df: pd.DataFrame, top_n: int = 15) -> go.Figure:
    """Horizontal bar of top N accounts by absolute persaldo."""
    if mapp_df is None or mapp_df.empty:
        return go.Figure()

    df = mapp_df.copy()
    df["abs_persaldo"] = df["persaldo"].abs()
    top = df.nlargest(top_n, "abs_persaldo")

    colors = [_COLORS["asset"] if s == "A" else _COLORS["liability"] if s == "P" else _COLORS["pl"]
              for s in top["side"]]

    fig = go.Figure(go.Bar(
        y=top["account_number"] + " " + top["account_name"].str[:30],
        x=top["persaldo"],
        orientation="h",
        marker_color=colors,
        text=top["persaldo"].apply(lambda v: f"{v/1e3:.1f}K" if abs(v) >= 1000 else f"{v:.0f}"),
        textposition="auto",
    ))
    fig.update_layout(
        title=f"Top {top_n} Accounts by Net Balance",
        **_LAYOUT_DEFAULTS,
        height=500,
        yaxis=dict(showgrid=False),
        xaxis=dict(showgrid=True, gridcolor=_COLORS["grid"]),
    )
    return fig


def mapping_status_donut(mapp_df: pd.DataFrame) -> go.Figure:
    """Donut of mapping status distribution."""
    if mapp_df is None or mapp_df.empty:
        return go.Figure()

    counts = mapp_df["mapping_status"].value_counts().reset_index()
    counts.columns = ["status", "count"]
    color_map = {"mapped": "#059669", "heuristic": "#D97706", "unmapped": "#DC2626"}
    colors = [color_map.get(s, "#64748B") for s in counts["status"]]

    fig = go.Figure(go.Pie(
        labels=counts["status"],
        values=counts["count"],
        hole=0.5,
        marker=dict(colors=colors),
    ))
    fig.update_layout(title="Mapping Coverage", **_LAYOUT_DEFAULTS)
    return fig
