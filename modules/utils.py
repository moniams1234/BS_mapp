"""utils.py — shared helpers."""
from __future__ import annotations

import streamlit as st

MONTHS = [
    "Styczeń","Luty","Marzec","Kwiecień","Maj","Czerwiec",
    "Lipiec","Sierpień","Wrzesień","Październik","Listopad","Grudzień",
]


def fmt(value: float, decimals: int = 2) -> str:
    """Format a float as a Polish-style number string."""
    try:
        return f"{value:,.{decimals}f}"
    except (ValueError, TypeError):
        return str(value)


def render_flags(flags: list[dict]) -> None:
    """Render red-flag messages in Streamlit."""
    if not flags:
        st.success("Brak red flags.")
        return
    for f in flags:
        t = f.get("type", "info")
        msg = f.get("message", "")
        if t == "error":
            st.error(f"🔴 {msg}")
        elif t == "warning":
            st.warning(f"🟡 {msg}")
        else:
            st.success(f"✅ {msg}")
