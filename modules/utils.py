"""
utils.py
Shared utility functions.
"""
from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st


FEEDBACK_FILE = Path(__file__).parent.parent / "user_feedback.json"


def fmt_currency(v: float, decimals: int = 0) -> str:
    """Format number as currency string."""
    if v is None:
        return "N/A"
    try:
        v = float(v)
    except (TypeError, ValueError):
        return "N/A"
    if abs(v) >= 1_000_000:
        return f"{v/1_000_000:,.2f}M"
    if abs(v) >= 1_000:
        return f"{v/1_000:,.1f}K"
    return f"{v:,.{decimals}f}"


def render_kpi_card(label: str, value: str, delta: str = "", color: str = "#2563EB"):
    """Render a styled KPI card using st.metric."""
    st.metric(label=label, value=value, delta=delta if delta else None)


def render_flag_section(flags: list[dict]):
    """Render red flags as colored Streamlit messages."""
    if not flags:
        st.success("No issues detected.")
        return

    errors = [f for f in flags if f.get("type") == "error"]
    warnings = [f for f in flags if f.get("type") == "warning"]
    successes = [f for f in flags if f.get("type") == "success"]

    for f in errors:
        st.error(f"🔴 **[{f.get('category', 'Error')}]** {f.get('message', '')}")
    for f in warnings:
        st.warning(f"🟡 **[{f.get('category', 'Warning')}]** {f.get('message', '')}")
    for f in successes:
        st.success(f"🟢 **[{f.get('category', 'OK')}]** {f.get('message', '')}")


def save_feedback(rating: int, comment: str):
    """Append feedback to local JSON file and session_state."""
    entry = {
        "timestamp": datetime.now().isoformat(),
        "rating": rating,
        "comment": comment,
    }
    # Save to session_state
    if "feedback_log" not in st.session_state:
        st.session_state["feedback_log"] = []
    st.session_state["feedback_log"].append(entry)

    # Try to save to local file
    try:
        existing = []
        if FEEDBACK_FILE.exists():
            with open(FEEDBACK_FILE, "r", encoding="utf-8") as f:
                existing = json.load(f)
        existing.append(entry)
        with open(FEEDBACK_FILE, "w", encoding="utf-8") as f:
            json.dump(existing, f, ensure_ascii=False, indent=2)
    except Exception:
        pass  # Fail silently – session_state already updated


def empty_state_message(msg: str = "Upload and analyze a file to see results."):
    st.info(f"📂 {msg}")


def check_session_data() -> bool:
    """Return True if analyzed data is available in session_state."""
    return bool(st.session_state.get("analyzed"))


def get_session(key: str, default: Any = None) -> Any:
    return st.session_state.get(key, default)
