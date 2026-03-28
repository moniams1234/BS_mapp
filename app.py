"""
app.py – Financial Analyzer v4
Run: python -m streamlit run app.py
"""
from __future__ import annotations

from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st

from modules.xlsx_parser    import (parse_trial_balance, parse_mapping_file,
                                    load_default_mapping)
from modules.mapping_engine import run_mapping, compute_kpis
from modules.balance_sheet  import build_balance_sheet, load_bs_group_order
from modules.pnl            import compute_pnl
from modules.anomaly_detection import build_red_flags
from modules.charts         import (balance_bar, assets_pie, liabilities_pie,
                                    mapp_group_bar, pnl_waterfall, mapping_donut)
from modules.export_utils   import build_excel_export, build_json_export
from modules.utils          import fmt, render_flags, save_feedback, MONTHS

# ─── page config ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Financial Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS – dark theme with improved contrast ─────────────────────────────────
st.markdown("""
<style>
/* ── Base ── */
html, body, [data-testid="stAppViewContainer"] {
    background-color: #0A0F1E;
    color: #E2E8F0;
}

/* ── Sidebar ── */
[data-testid="stSidebar"] {
    background: #0F172A !important;
    border-right: 1px solid #1E3A5F;
}
[data-testid="stSidebar"] * { color: #CBD5E1 !important; }
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 { color: #93C5FD !important; }

/* ── File uploader – dark bg, visible border ── */
[data-testid="stFileUploader"] {
    background: #1E293B !important;
    border: 1.5px dashed #3B82F6 !important;
    border-radius: 8px !important;
    padding: 8px !important;
}
[data-testid="stFileUploader"] * { color: #CBD5E1 !important; }
[data-testid="stFileUploader"]:hover {
    border-color: #60A5FA !important;
    background: #1E3A5F !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] { color: #94A3B8 !important; }
[data-testid="stFileUploaderDropzoneInstructions"] small { color: #64748B !important; }

/* ── Uploaded file pill ── */
[data-testid="stFileUploaderFile"] {
    background: #1E3A5F !important;
    border: 1px solid #2563EB !important;
    border-radius: 6px !important;
}

/* ── Buttons ── */
[data-testid="stSidebar"] .stButton > button {
    background: #1E3A5F !important;
    border: 1px solid #3B82F6 !important;
    color: #E2E8F0 !important;
    border-radius: 6px;
    font-weight: 600;
    width: 100%;
    transition: all 0.2s;
}
[data-testid="stSidebar"] .stButton > button:hover {
    background: #2563EB !important;
    border-color: #60A5FA !important;
    color: #FFFFFF !important;
}
[data-testid="stSidebar"] .stButton > button[kind="primary"] {
    background: #2563EB !important;
    border-color: #60A5FA !important;
}
[data-testid="stSidebar"] .stButton > button[kind="primary"]:hover {
    background: #1D4ED8 !important;
}

/* ── Download buttons (export) ── */
[data-testid="stDownloadButton"] > button {
    background: #065F46 !important;
    border: 1px solid #10B981 !important;
    color: #D1FAE5 !important;
    border-radius: 6px;
    font-weight: 600;
    width: 100%;
}
[data-testid="stDownloadButton"] > button:hover {
    background: #047857 !important;
    border-color: #34D399 !important;
    color: #FFFFFF !important;
}

/* ── Metrics ── */
div[data-testid="metric-container"] {
    background: #1E293B !important;
    border-radius: 10px !important;
    padding: 14px 18px !important;
    border-left: 4px solid #3B82F6 !important;
    box-shadow: 0 2px 8px rgba(0,0,0,0.3);
}
div[data-testid="metric-container"] label { color: #94A3B8 !important; }
div[data-testid="metric-container"] [data-testid="stMetricValue"] {
    color: #F1F5F9 !important;
    font-family: 'JetBrains Mono', 'Consolas', monospace;
}

/* ── DataFrames ── */
.stDataFrame { border: 1px solid #1E3A5F !important; border-radius: 8px; }
.stDataFrame thead th {
    background: #1E3A5F !important;
    color: #93C5FD !important;
}
.stDataFrame tbody tr:nth-child(even) { background: #0F1A2E !important; }
.stDataFrame tbody tr:hover { background: #1E3A5F !important; }

/* ── Radio buttons ── */
[data-testid="stSidebar"] .stRadio label { color: #CBD5E1 !important; }
[data-testid="stSidebar"] .stRadio [data-baseweb="radio"] div { border-color: #3B82F6 !important; }

/* ── Select / number input ── */
[data-testid="stSidebar"] .stSelectbox select,
[data-testid="stSidebar"] .stNumberInput input {
    background: #1E293B !important;
    border: 1px solid #334155 !important;
    color: #E2E8F0 !important;
}

/* ── Expander ── */
[data-testid="stExpander"] {
    background: #1E293B !important;
    border: 1px solid #334155 !important;
    border-radius: 8px;
}

/* ── Info / warning / error boxes ── */
.stAlert { border-radius: 8px !important; }

/* ── Headings ── */
h1, h2, h3 { color: #F1F5F9; }
h1 { border-bottom: 2px solid #1E3A5F; padding-bottom: 8px; }

/* ── Section dividers ── */
hr { border-color: #1E3A5F !important; }

/* ── Captions ── */
[data-testid="stSidebar"] .stCaption { color: #64748B !important; }

/* ── Export section label ── */
.export-label {
    color: #10B981;
    font-weight: 700;
    font-size: 0.9rem;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    margin-bottom: 6px;
}
</style>
""", unsafe_allow_html=True)

# ─── session defaults ────────────────────────────────────────────────────────
for _k, _v in {
    "analyzed":       False,
    "df":             None,
    "mapp_df":        None,
    "bs":             {},
    "pnl":            {},
    "kpis":           {},
    "flags":          [],
    "mapping":        {},
    "mapping_name":   "",
    "tb_name":        "",
    "feedback_log":   [],
    "bs_ref_order":   [],
}.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v

# ─── load BS reference order once ───────────────────────────────────────────
if not st.session_state["bs_ref_order"]:
    st.session_state["bs_ref_order"] = load_bs_group_order()

# ════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════════════════════════════════════
SECTIONS = ["🗺️ Mapp", "⚖️ Balance Sheet", "📉 P&L"]

with st.sidebar:
    # ── Header ──────────────────────────────────────────────────────────────
    st.markdown("## 📊 Financial Analyzer")
    st.markdown("---")

    # ── 1. Navigation ────────────────────────────────────────────────────────
    st.markdown("### Nawigacja")
    section = st.radio("nav", SECTIONS, label_visibility="collapsed")
    st.markdown("---")

    # ── 2. Upload Danych ─────────────────────────────────────────────────────
    st.markdown("### 📂 Upload Danych")

    tb_file = st.file_uploader(
        "Trial Balance (XLSX)", type=["xlsx"], key="tb_upload",
        help="Plik ZOiS z trial balance"
    )

    # Mapping choice
    mapping_mode = st.radio(
        "Mapping",
        ["Użyj domyślnego mappingu", "Wgraj własny mapping"],
        index=0,
        key="mapping_mode",
    )
    map_file = None
    if mapping_mode == "Wgraj własny mapping":
        map_file = st.file_uploader(
            "Mapping (XLSX)", type=["xlsx"], key="map_upload",
            help="Plik z arkuszem 'Mapp' lub kolumną 'BS Mapp'"
        )

    # Status info
    tb_name  = st.session_state["tb_name"]
    map_name = st.session_state["mapping_name"]
    st.caption(f"📄 Source File: **{tb_name or '—'}**")
    st.caption(f"🗂️ Mapping: **{map_name or '—'}**")
    if mapping_mode == "Użyj domyślnego mappingu":
        st.caption("ℹ️ Używany wbudowany domyślny mapping.")
    elif map_file is None and st.session_state["mapping"]:
        st.caption("ℹ️ Używany ostatnio wgrany mapping.")
    elif map_file is None and not st.session_state["mapping"] and mapping_mode == "Wgraj własny mapping":
        st.caption("⚠️ Brak mappingu – wgraj plik.")

    # ── 3. Okres Raportowy ──────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### 📅 Okres Raportowy")
    col_m, col_y = st.columns(2)
    with col_m:
        month_name = st.selectbox(
            "mies", MONTHS,
            index=datetime.now().month - 1,
            label_visibility="collapsed"
        )
    with col_y:
        year = st.number_input(
            "rok", min_value=2000, max_value=2100,
            value=datetime.now().year, step=1,
            label_visibility="collapsed"
        )
    period_str = f"{month_name} {year}"

    # ── 4. Akcje ─────────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### ⚡ Akcje")
    mappuj_clicked = st.button("▶ Mappuj", use_container_width=True, type="primary")

    if st.button("🗑️ Wyczyść dane", use_container_width=True):
        for k in ["analyzed", "df", "mapp_df", "bs", "pnl", "kpis", "flags", "tb_name"]:
            if k == "analyzed":
                st.session_state[k] = False
            elif k in ("bs", "pnl", "kpis"):
                st.session_state[k] = {}
            elif k == "flags":
                st.session_state[k] = []
            else:
                st.session_state[k] = None
        st.session_state["tb_name"] = ""
        st.rerun()

    # ── 5. Status ────────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### 📊 Status")
    kpis  = st.session_state["kpis"]
    flags = st.session_state["flags"]
    n_err  = sum(1 for f in flags if f.get("type") == "error")
    n_warn = sum(1 for f in flags if f.get("type") == "warning")
    st.caption(f"Total Accounts:    **{kpis.get('total_accounts', '—')}**")
    st.caption(f"Mapped Accounts:   **{kpis.get('mapped_accounts', '—')}**")
    st.caption(f"Unmapped Accounts: **{kpis.get('unmapped_accounts', '—')}**")
    st.caption(f"Excluded (9xx):    **{kpis.get('excluded_accounts', '—')}**")
    st.caption(f"🔴 Red Flag Errors:   **{n_err}**")
    st.caption(f"🟡 Red Flag Warnings: **{n_warn}**")

    # ── 6. Eksport ───────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown('<div class="export-label">💾 Eksport</div>', unsafe_allow_html=True)

    if st.session_state.get("analyzed"):
        _raw  = st.session_state["df"]
        _mapp = st.session_state["mapp_df"]
        _bs   = st.session_state["bs"]
        _pnl  = st.session_state["pnl"]
        _flgs = st.session_state["flags"]
        _kpis = st.session_state["kpis"]
        _map  = st.session_state["mapping"]
        _fn   = st.session_state["tb_name"]

        xlsx_bytes = build_excel_export(
            _raw, _map, _mapp, _bs, _pnl, _flgs, _kpis, _fn, period_str
        )
        st.download_button(
            "⬇ Excel (.xlsx)",
            data=xlsx_bytes,
            file_name=f"analiza_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        json_str = build_json_export(
            _mapp, _bs, _pnl, _flgs, _kpis, _fn, period_str
        )
        st.download_button(
            "⬇ JSON (.json)",
            data=json_str.encode("utf-8"),
            file_name=f"analiza_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
            mime="application/json",
            use_container_width=True,
        )
    else:
        st.caption("Wykonaj Mappuj, aby odblokować eksport.")


# ════════════════════════════════════════════════════════════════════════════
# MAPPING PIPELINE
# ════════════════════════════════════════════════════════════════════════════
if mappuj_clicked:
    errors: list[str] = []

    # ── Load mapping ─────────────────────────────────────────────────────────
    if mapping_mode == "Użyj domyślnego mappingu":
        result = load_default_mapping()
        if "error" in result and result["error"]:
            errors.append(f"Błąd domyślnego mappingu: {result['error']}")
        else:
            st.session_state["mapping"]      = result["mapping"]
            st.session_state["mapping_name"] = "Domyślny (wbudowany)"
    else:
        # Custom mapping
        if map_file is not None:
            res = parse_mapping_file(map_file)
            if "error" in res and res["error"]:
                errors.append(f"Błąd mappingu: {res['error']}")
            else:
                st.session_state["mapping"]      = res["mapping"]
                st.session_state["mapping_name"] = map_file.name
        elif not st.session_state["mapping"]:
            errors.append("Brak pliku mappingu. Wgraj plik Mapping XLSX.")

    # ── Load trial balance ───────────────────────────────────────────────────
    if tb_file is None:
        errors.append("Brak pliku trial balance. Wgraj plik XLSX.")
    else:
        tb_result = parse_trial_balance(tb_file)
        if "error" in tb_result and tb_result["error"]:
            errors.append(f"Błąd parsowania: {tb_result['error']}")
        else:
            st.session_state["df"]      = tb_result["df"]
            st.session_state["tb_name"] = tb_file.name

    if errors:
        for e in errors:
            st.error(f"❌ {e}")
        st.stop()

    df      = st.session_state["df"]
    mapping = st.session_state["mapping"]

    with st.spinner("Mapowanie kont…"):
        mapp_df = run_mapping(df, mapping)
        st.session_state["mapp_df"] = mapp_df

    with st.spinner("Balance Sheet…"):
        ref_order = st.session_state["bs_ref_order"]
        bs = build_balance_sheet(mapp_df, ref_order)
        st.session_state["bs"] = bs

    with st.spinner("P&L…"):
        pnl = compute_pnl(df)
        st.session_state["pnl"] = pnl

    with st.spinner("KPI i Red Flags…"):
        kpis  = compute_kpis(mapp_df)
        st.session_state["kpis"] = kpis
        flags = build_red_flags(
            df, mapp_df, bs, kpis,
            tb_result.get("warnings", []) if "tb_result" in dir() else []
        )
        st.session_state["flags"] = flags

    st.session_state["analyzed"] = True
    st.success(
        f"✅ Mappowanie zakończone — "
        f"{kpis['total_accounts']:,} kont | "
        f"Zmapowane: {kpis['mapped_accounts']:,} | "
        f"Niezmapowane: {kpis['unmapped_accounts']:,}"
    )


# ════════════════════════════════════════════════════════════════════════════
# VIEWS
# ════════════════════════════════════════════════════════════════════════════

def _no_data():
    st.info("📂 Wgraj plik Trial Balance, wybierz mapping i kliknij **▶ Mappuj**.")


# ──────────────────────────────────────────────────────────────────────────
# MAPP
# ──────────────────────────────────────────────────────────────────────────
if section == "🗺️ Mapp":
    st.markdown(f"## 🗺️ Tabela Mappingu — {period_str}")

    if not st.session_state.get("analyzed"):
        _no_data()
    else:
        mapp_df = st.session_state["mapp_df"]
        kpis    = st.session_state["kpis"]
        flags   = st.session_state["flags"]

        # KPI row
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Accounts",  f"{kpis.get('total_accounts', 0):,}")
        c2.metric("Mapped",          f"{kpis.get('mapped_accounts', 0):,}")
        c3.metric("Unmapped",        f"{kpis.get('unmapped_accounts', 0):,}")
        c4.metric("Excluded (9xx)",  f"{kpis.get('excluded_accounts', 0):,}")

        st.markdown("---")

        # Mapping status donut
        col1, col2 = st.columns([1, 2])
        with col1:
            st.plotly_chart(mapping_donut(
                kpis.get("mapped_accounts", 0),
                kpis.get("unmapped_accounts", 0),
                kpis.get("excluded_accounts", 0),
            ), use_container_width=True)
        with col2:
            st.markdown("### 🚩 Red Flags")
            render_flags(flags)

        st.markdown("---")

        # Filters
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            side_opts   = sorted(mapp_df["side"].unique().tolist())
            side_filter = st.multiselect("Typ (side)", side_opts, default=side_opts)
        with col_f2:
            status_opts   = sorted(mapp_df["mapping_status"].unique().tolist())
            status_filter = st.multiselect("Status", status_opts, default=status_opts)
        with col_f3:
            search = st.text_input("Szukaj konta / nazwy", "")

        filtered = mapp_df[
            mapp_df["side"].isin(side_filter) &
            mapp_df["mapping_status"].isin(status_filter)
        ]
        if search:
            mask = (
                filtered["account_number"].str.contains(search, case=False, na=False) |
                filtered["account_name"].str.contains(search, case=False, na=False)
            )
            filtered = filtered[mask]

        cols = ["account_number", "account_name", "group", "side",
                "debit", "credit", "saldo_dt", "saldo_ct",
                "persaldo", "persaldo_group", "mapping_status"]
        cols = [c for c in cols if c in filtered.columns]

        st.markdown(f"**{len(filtered):,}** z **{len(mapp_df):,}** kont")
        st.dataframe(
            filtered[cols].style.format({
                c: "{:,.2f}" for c in
                ["debit","credit","saldo_dt","saldo_ct","persaldo","persaldo_group"]
                if c in filtered.columns
            }),
            use_container_width=True,
            height=460,
        )

        st.markdown("### Sumy wg grupy")
        grp_sum = (
            filtered[filtered["mapping_status"] == "mapped"]
            .groupby(["side", "group"])
            .agg(kont=("account_number","count"), persaldo=("persaldo","sum"))
            .reset_index()
            .sort_values(["side", "persaldo"], ascending=[True, False])
        )
        st.dataframe(
            grp_sum.style.format({"persaldo": "{:,.2f}"}),
            use_container_width=True,
        )

        with st.expander("📖 Zasady liczenia persaldo"):
            st.markdown("""
            | Typ | Formuła | Znaczenie dodatniego |
            |---|---|---|
            | **A** (Aktywa) | `Saldo Dt − Saldo Ct` | Saldo debetowe (normalne dla aktywów) |
            | **P** (Pasywa) | `Saldo Ct − Saldo Dt` | Saldo kredytowe (normalne dla pasywów) |
            | **X** (Wykluczone) | `Saldo Dt − Saldo Ct` | Tylko informacyjnie |

            **Persaldo per grupę** = suma persaldo wszystkich kont w danej grupie raportowej.
            """)


# ──────────────────────────────────────────────────────────────────────────
# BALANCE SHEET
# ──────────────────────────────────────────────────────────────────────────
elif section == "⚖️ Balance Sheet":
    st.markdown(f"## ⚖️ Balance Sheet — {period_str}")

    if not st.session_state.get("analyzed"):
        _no_data()
    else:
        bs   = st.session_state["bs"]
        kpis = st.session_state["kpis"]

        ta   = bs.get("total_assets", 0)
        tl   = bs.get("total_liabilities", 0)
        diff = bs.get("difference", 0)

        c1, c2, c3 = st.columns(3)
        c1.metric("Total Assets",      fmt(ta))
        c2.metric("Total Liabilities", fmt(tl))
        c3.metric("Różnica", fmt(diff),
                  delta="✓ Zbilansowany" if abs(diff) < 1 else "⚠ Niezbalansowany")

        st.markdown("---")

        col_a, col_p = st.columns(2)

        with col_a:
            st.markdown("### AKTYWA (A)")
            ag = bs.get("assets_by_group")
            if ag is not None and not ag.empty:
                # Show in BS order (no re-sorting)
                st.dataframe(
                    ag.style.format({"amount": "{:,.2f}"}),
                    use_container_width=True,
                    height=380,
                )
            st.markdown(f"**Razem Aktywa: {fmt(ta)}**")

        with col_p:
            st.markdown("### PASYWA (P)")
            lg = bs.get("liabilities_by_group")
            if lg is not None and not lg.empty:
                st.dataframe(
                    lg.style.format({"amount": "{:,.2f}"}),
                    use_container_width=True,
                    height=380,
                )
            st.markdown(f"**Razem Pasywa: {fmt(tl)}**")

        st.markdown("---")
        if abs(diff) < 1:
            st.success(f"✅ Bilans zbilansowany. Różnica: {fmt(diff)}")
        else:
            st.error(f"❌ Bilans NIE jest zbilansowany. Różnica: {fmt(diff)}")

        st.markdown("### Struktura wizualna")
        c_charts1, c_charts2 = st.columns(2)
        with c_charts1:
            ag = bs.get("assets_by_group", pd.DataFrame(columns=["group","amount"]))
            st.plotly_chart(assets_pie(ag), use_container_width=True)
        with c_charts2:
            lg = bs.get("liabilities_by_group", pd.DataFrame(columns=["group","amount"]))
            st.plotly_chart(liabilities_pie(lg), use_container_width=True)

        st.plotly_chart(balance_bar(ta, tl), use_container_width=True)

        with st.expander("ℹ️ O kolejności pozycji BS"):
            st.markdown("""
            Pozycje Balance Sheet są prezentowane w kolejności zdefiniowanej
            przez plik referencyjny `data/bs_order.xlsx` (arkusz `BS`, kolumna B)
            oraz kolejność grup w pliku mappingu. Pozycje spoza listy referencyjnej
            są dołączane na końcu.
            """)


# ──────────────────────────────────────────────────────────────────────────
# P&L
# ──────────────────────────────────────────────────────────────────────────
elif section == "📉 P&L":
    st.markdown(f"## 📉 Rachunek Wyników (P&L) — {period_str}")

    if not st.session_state.get("analyzed"):
        _no_data()
    else:
        pnl    = st.session_state["pnl"]
        pnl_df = pnl.get("pnl_df", pd.DataFrame())
        nr     = pnl.get("net_result", 0)

        st.metric(
            "Wynik Netto",
            fmt(nr),
            delta="Zysk ✓" if nr > 0 else "Strata ✗",
        )

        st.markdown("---")
        st.markdown("""
        **Reguły kwalifikacji kont P&L:**

        | Kryterium | Wartość |
        |---|---|
        | Konta syntetyczne (3 znaki) zaczynające się od **4** | ✓ (poza 409 i 490) |
        | Konta syntetyczne (3 znaki) zaczynające się od **7** | ✓ (poza 409 i 490) |
        | Konto **409** | ✗ Wykluczone |
        | Konto **490** | ✗ Wykluczone |
        | Konto **870** | ✓ Dodatkowe |
        | Konto **590** | ✓ Dodatkowe |
        | Persaldo | `Saldo Ct − Saldo Dt` (+ = przychód) |
        """)

        if pnl_df.empty:
            st.warning("Brak kont P&L po zastosowaniu reguł.")
        else:
            income  = pnl_df[pnl_df["persaldo_pnl"] > 0]["persaldo_pnl"].sum()
            expense = pnl_df[pnl_df["persaldo_pnl"] < 0]["persaldo_pnl"].sum()

            c1, c2, c3 = st.columns(3)
            c1.metric("Przychody (+)",  fmt(income))
            c2.metric("Koszty (−)",     fmt(expense))
            c3.metric("Wynik Netto",    fmt(nr))

            st.plotly_chart(pnl_waterfall(pnl_df, nr), use_container_width=True)

            st.markdown("### Konta P&L")
            cols = ["account_number", "account_name", "saldo_dt", "saldo_ct",
                    "persaldo_pnl", "pnl_type"]
            cols = [c for c in cols if c in pnl_df.columns]
            st.dataframe(
                pnl_df[cols].style.format({
                    c: "{:,.2f}" for c in ["saldo_dt","saldo_ct","persaldo_pnl"]
                    if c in pnl_df.columns
                }),
                use_container_width=True,
                height=440,
            )
