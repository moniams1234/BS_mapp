"""
Financial Analyzer
Run: python -m streamlit run app.py
"""
from __future__ import annotations
from datetime import datetime

import pandas as pd
import streamlit as st

from modules.xlsx_parser    import (parse_trial_balance, parse_mapping_file,
                                    load_default_mapping)
from modules.mapping_engine import run_mapping, compute_kpis
from modules.balance_sheet  import build_balance_sheet, load_bs_group_order
from modules.pnl            import compute_pnl
from modules.anomaly_detection import build_red_flags
from modules.charts         import (balance_bar, assets_pie, liabilities_pie,
                                    pnl_waterfall, mapping_donut)
from modules.export_utils   import build_excel_export, build_json_export
from modules.utils          import fmt, render_flags, MONTHS

# ─── Page config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Financial Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS ─────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* Base */
html, body, [data-testid="stAppViewContainer"] {
    background-color: #080D1A;
    color: #E2E8F0;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background: #0D1526 !important;
    border-right: 1px solid #1E3A5F;
    padding-top: 0;
}
[data-testid="stSidebar"] * { color: #CBD5E1 !important; }
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 { color: #93C5FD !important; }
[data-testid="stSidebar"] .stCaption { color: #64748B !important; }

/* File uploader — dark, visible */
[data-testid="stFileUploader"] {
    background: #111827 !important;
    border: 1.5px dashed #2563EB !important;
    border-radius: 8px !important;
    padding: 10px 12px !important;
}
[data-testid="stFileUploader"]:hover {
    border-color: #60A5FA !important;
    background: #1E2D4A !important;
}
[data-testid="stFileUploader"] * { color: #CBD5E1 !important; }
[data-testid="stFileUploaderDropzoneInstructions"] { color: #94A3B8 !important; }
[data-testid="stFileUploaderDropzoneInstructions"] small { color: #64748B !important; }
[data-testid="stFileUploaderFile"] {
    background: #1E3A5F !important;
    border: 1px solid #2563EB !important;
    border-radius: 6px !important;
}

/* Action buttons */
[data-testid="stSidebar"] .stButton > button {
    background: #1E3A5F !important;
    border: 1px solid #3B82F6 !important;
    color: #E2E8F0 !important;
    border-radius: 6px;
    font-weight: 600;
    width: 100%;
    transition: background 0.2s, border-color 0.2s;
}
[data-testid="stSidebar"] .stButton > button:hover {
    background: #2563EB !important;
    border-color: #60A5FA !important;
    color: #fff !important;
}
[data-testid="stSidebar"] .stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #1D4ED8, #2563EB) !important;
    border-color: #60A5FA !important;
    font-size: 0.95rem;
    padding: 10px;
}
[data-testid="stSidebar"] .stButton > button[kind="primary"]:hover {
    background: linear-gradient(135deg, #1E40AF, #1D4ED8) !important;
}

/* Download buttons — green, always active-looking */
[data-testid="stDownloadButton"] > button {
    background: #064E3B !important;
    border: 1.5px solid #10B981 !important;
    color: #D1FAE5 !important;
    border-radius: 6px;
    font-weight: 600;
    width: 100%;
    transition: background 0.2s;
}
[data-testid="stDownloadButton"] > button:hover {
    background: #065F46 !important;
    border-color: #34D399 !important;
    color: #fff !important;
}
[data-testid="stDownloadButton"] > button:disabled {
    opacity: 0.4 !important;
}

/* Radio */
[data-testid="stSidebar"] .stRadio label { color: #CBD5E1 !important; }

/* Metrics */
div[data-testid="metric-container"] {
    background: #111827 !important;
    border-radius: 10px !important;
    padding: 14px 18px !important;
    border-left: 4px solid #3B82F6 !important;
    box-shadow: 0 2px 8px rgba(0,0,0,0.4);
}
div[data-testid="metric-container"] label { color: #94A3B8 !important; }
div[data-testid="metric-container"] [data-testid="stMetricValue"] {
    color: #F1F5F9 !important;
    font-family: 'JetBrains Mono','Consolas',monospace;
}

/* DataFrames */
.stDataFrame { border: 1px solid #1E3A5F !important; border-radius: 8px; }
.stDataFrame thead th { background: #1E3A5F !important; color: #93C5FD !important; }
.stDataFrame tbody tr:nth-child(even) { background: #0D1526 !important; }
.stDataFrame tbody tr:hover { background: #1E3A5F !important; }

/* Headings */
h1, h2, h3 { color: #F1F5F9; }
h1 { border-bottom: 2px solid #1E3A5F; padding-bottom: 8px; }
hr { border-color: #1E3A5F !important; }

/* Expander */
[data-testid="stExpander"] {
    background: #111827 !important;
    border: 1px solid #1E3A5F !important;
    border-radius: 8px;
}

/* Status badge */
.status-badge {
    display: inline-block;
    padding: 3px 10px;
    border-radius: 12px;
    font-size: 0.78rem;
    font-weight: 700;
    background: #1E3A5F;
    color: #93C5FD;
    margin-bottom: 4px;
}

/* Export label */
.export-label {
    color: #10B981;
    font-weight: 700;
    font-size: 0.85rem;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    margin-bottom: 4px;
}

/* Mapping mode container */
.mapping-box {
    background: #111827;
    border: 1px solid #1E3A5F;
    border-radius: 8px;
    padding: 12px;
    margin-bottom: 8px;
}
</style>
""", unsafe_allow_html=True)

# ─── Session state defaults ──────────────────────────────────────────────────
_DEFAULTS = {
    "analyzed":     False,
    "df":           None,
    "mapp_df":      None,
    "bs":           {},
    "pnl":          {},
    "kpis":         {},
    "flags":        [],
    "mapping":      {},
    "mapping_name": "",
    "tb_name":      "",
    "bs_ref_order": [],
}
for k, v in _DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v

# Load BS reference order once at startup
if not st.session_state["bs_ref_order"]:
    st.session_state["bs_ref_order"] = load_bs_group_order()

# ════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════════════════════════════════════
SECTIONS = ["🗺️ Mapp", "⚖️ Balance Sheet", "📉 P&L"]

with st.sidebar:
    st.markdown("## 📊 Financial Analyzer")
    st.markdown("---")

    # Navigation
    st.markdown("### Widok")
    section = st.radio("section", SECTIONS, label_visibility="collapsed")
    st.markdown("---")

    # Mapping selection
    st.markdown("### 🗂️ Mapping")
    mapping_mode = st.radio(
        "mapping_mode",
        ["✅ Użyj domyślnego mappingu", "📁 Wgraj własny mapping"],
        index=0,
        label_visibility="collapsed",
    )
    use_default = mapping_mode.startswith("✅")

    map_file = None
    if not use_default:
        map_file = st.file_uploader(
            "Plik Mapping (XLSX)",
            type=["xlsx"],
            key="map_upload",
            help="XLSX z arkuszem 'Mapp' lub kolumną 'BS Mapp'",
        )
    else:
        st.caption("📦 Domyślny mapping: `data/default_mapping.xlsx`")

    st.markdown("---")

    # Trial Balance upload
    st.markdown("### 📄 Trial Balance")
    tb_file = st.file_uploader(
        "Plik Trial Balance (XLSX)",
        type=["xlsx"],
        key="tb_upload",
        help="ZOiS — plik z próbnym bilansem (saldo otwarcia, obroty, saldo)",
    )

    # Period
    st.markdown("---")
    st.markdown("### 📅 Okres")
    col_m, col_y = st.columns(2)
    with col_m:
        month_name = st.selectbox(
            "mies", MONTHS,
            index=datetime.now().month - 1,
            label_visibility="collapsed",
        )
    with col_y:
        year = st.number_input(
            "rok", min_value=2000, max_value=2100,
            value=datetime.now().year, step=1,
            label_visibility="collapsed",
        )
    period_str = f"{month_name} {year}"

    # Run button
    st.markdown("---")
    mappuj_clicked = st.button("▶ Uruchom Mappowanie", use_container_width=True, type="primary")
    if st.button("🗑️ Wyczyść", use_container_width=True):
        for k in list(_DEFAULTS.keys()):
            st.session_state[k] = _DEFAULTS[k]
        st.rerun()

    # Status summary
    st.markdown("---")
    kpis  = st.session_state["kpis"]
    flags = st.session_state["flags"]
    n_err  = sum(1 for f in flags if f.get("type") == "error")
    n_warn = sum(1 for f in flags if f.get("type") == "warning")
    st.markdown("### 📊 Status")
    st.caption(f"Source: **{st.session_state['tb_name'] or '—'}**")
    st.caption(f"Mapping: **{st.session_state['mapping_name'] or '—'}**")
    st.caption(f"Kont łącznie:   **{kpis.get('total_accounts', '—')}**")
    st.caption(f"Zmapowane:      **{kpis.get('mapped_accounts', '—')}**")
    st.caption(f"Niezmapowane:   **{kpis.get('unmapped_accounts', '—')}**")
    st.caption(f"Wykluczone 9xx: **{kpis.get('excluded_accounts', '—')}**")
    if n_err:
        st.caption(f"🔴 Błędy: **{n_err}**")
    if n_warn:
        st.caption(f"🟡 Ostrzeżenia: **{n_warn}**")

    # Export
    if st.session_state.get("analyzed"):
        st.markdown("---")
        st.markdown('<div class="export-label">💾 Eksport</div>', unsafe_allow_html=True)

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
        json_str = build_json_export(_mapp, _bs, _pnl, _flgs, _kpis, _fn, period_str)
        st.download_button(
            "⬇ JSON (.json)",
            data=json_str.encode("utf-8"),
            file_name=f"analiza_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
            mime="application/json",
            use_container_width=True,
        )

# ════════════════════════════════════════════════════════════════════════════
# PIPELINE
# ════════════════════════════════════════════════════════════════════════════
if mappuj_clicked:
    errors: list[str] = []

    # 1. Load mapping
    if use_default:
        result = load_default_mapping()
        if result.get("error"):
            errors.append(f"Błąd domyślnego mappingu: {result['error']}")
        else:
            st.session_state["mapping"]      = result["mapping"]
            st.session_state["mapping_name"] = "Domyślny (data/default_mapping.xlsx)"
    else:
        if map_file is None and not st.session_state["mapping"]:
            errors.append("Brak pliku mappingu. Wgraj plik XLSX lub wybierz domyślny mapping.")
        elif map_file is not None:
            result = parse_mapping_file(map_file)
            if result.get("error"):
                errors.append(f"Błąd parsowania mappingu: {result['error']}")
            else:
                st.session_state["mapping"]      = result["mapping"]
                st.session_state["mapping_name"] = map_file.name

    # 2. Load trial balance
    if tb_file is None:
        errors.append("Brak pliku Trial Balance. Wgraj plik XLSX.")
    else:
        tb_result = parse_trial_balance(tb_file)
        if tb_result.get("error"):
            errors.append(f"Błąd parsowania Trial Balance: {tb_result['error']}")
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

    with st.spinner("Budowanie Balance Sheet…"):
        bs = build_balance_sheet(mapp_df, st.session_state["bs_ref_order"])
        st.session_state["bs"] = bs

    with st.spinner("Obliczanie P&L…"):
        pnl = compute_pnl(df)
        st.session_state["pnl"] = pnl

    with st.spinner("KPI i Red Flags…"):
        kpis = compute_kpis(mapp_df)
        st.session_state["kpis"] = kpis
        flags = build_red_flags(
            df, mapp_df, bs, kpis,
            tb_result.get("warnings", []),
        )
        st.session_state["flags"] = flags

    st.session_state["analyzed"] = True
    st.success(
        f"✅ Gotowe — {kpis['total_accounts']:,} kont | "
        f"zmapowane: {kpis['mapped_accounts']:,} | "
        f"niezmapowane: {kpis['unmapped_accounts']:,} | "
        f"wykluczone: {kpis['excluded_accounts']:,}"
    )

# ════════════════════════════════════════════════════════════════════════════
# ONBOARDING (when no data)
# ════════════════════════════════════════════════════════════════════════════
def _no_data():
    st.markdown("## 📊 Financial Analyzer")
    st.info(
        "**Jak zacząć:**\n\n"
        "1. Wybierz mapping w panelu bocznym (domyślny lub własny XLSX)\n"
        "2. Wgraj plik **Trial Balance** (XLSX / ZOiS)\n"
        "3. Kliknij **▶ Uruchom Mappowanie**\n"
        "4. Przejdź do zakładek **Mapp · Balance Sheet · P&L**"
    )

# ════════════════════════════════════════════════════════════════════════════
# MAPP VIEW
# ════════════════════════════════════════════════════════════════════════════
if section == "🗺️ Mapp":
    st.markdown(f"## 🗺️ Mapp — {period_str}")

    if not st.session_state.get("analyzed"):
        _no_data()
    else:
        mapp_df = st.session_state["mapp_df"]
        kpis    = st.session_state["kpis"]
        flags   = st.session_state["flags"]

        # KPI row
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Kont łącznie",    f"{kpis.get('total_accounts', 0):,}")
        c2.metric("Zmapowane",       f"{kpis.get('mapped_accounts', 0):,}")
        c3.metric("Niezmapowane",    f"{kpis.get('unmapped_accounts', 0):,}")
        c4.metric("Wykluczone 9xx",  f"{kpis.get('excluded_accounts', 0):,}")

        st.markdown("---")

        # Charts + flags
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
        f1, f2, f3 = st.columns(3)
        with f1:
            side_opts   = sorted(mapp_df["side"].unique())
            side_filter = st.multiselect("Strona (side)", side_opts, default=side_opts)
        with f2:
            status_opts   = sorted(mapp_df["mapping_status"].unique())
            status_filter = st.multiselect("Status", status_opts, default=status_opts)
        with f3:
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

        show_cols = [c for c in
                     ["account_number","account_name","group","side",
                      "saldo_dt","saldo_ct","persaldo","persaldo_group","mapping_status"]
                     if c in filtered.columns]

        st.markdown(f"**{len(filtered):,}** z **{len(mapp_df):,}** kont")
        st.dataframe(
            filtered[show_cols].style.format({
                c: "{:,.2f}" for c in
                ["saldo_dt","saldo_ct","persaldo","persaldo_group"]
                if c in filtered.columns
            }),
            use_container_width=True,
            height=460,
        )

        # Group sums
        st.markdown("### Sumy wg grupy")
        grp = (
            filtered[filtered["mapping_status"] == "mapped"]
            .groupby(["side","group"])
            .agg(kont=("account_number","count"), persaldo=("persaldo","sum"))
            .reset_index()
            .sort_values(["side","persaldo"], ascending=[True, False])
        )
        st.dataframe(
            grp.style.format({"persaldo": "{:,.2f}"}),
            use_container_width=True,
        )

        with st.expander("ℹ️ Zasady persaldo"):
            st.markdown("""
| Strona | Formuła | Interpretacja |
|---|---|---|
| **A** (Aktywa) | `Saldo Dt − Saldo Ct` | Saldo debetowe |
| **P** (Pasywa/Kapitał) | `Saldo Ct − Saldo Dt` | Saldo kredytowe |
| **X** (Wykluczone z BS) | `Saldo Dt − Saldo Ct` | Tylko info |
| **excluded** (9xx) | — | Wykluczone automatycznie |
""")

# ════════════════════════════════════════════════════════════════════════════
# BALANCE SHEET VIEW
# ════════════════════════════════════════════════════════════════════════════
elif section == "⚖️ Balance Sheet":
    st.markdown(f"## ⚖️ Balance Sheet — {period_str}")

    if not st.session_state.get("analyzed"):
        _no_data()
    else:
        bs   = st.session_state["bs"]
        ta   = bs.get("total_assets", 0.0)
        tl   = bs.get("total_liabilities", 0.0)
        diff = bs.get("difference", 0.0)

        # Summary metrics
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Assets",      fmt(ta))
        c2.metric("Total Liabilities", fmt(tl))
        c3.metric(
            "Różnica (A−P)",
            fmt(diff),
            delta="✓ Zbilansowany" if abs(diff) < 1 else "⚠ Niezbalansowany",
        )

        if abs(diff) < 1:
            st.success(f"✅ Bilans zbilansowany. Różnica: {fmt(diff)}")
        else:
            st.error(f"❌ Bilans NIE jest zbilansowany. Różnica: {fmt(diff)}")

        st.markdown("---")

        # Two-column layout
        col_a, col_p = st.columns(2)

        ag = bs.get("assets_by_group", pd.DataFrame(columns=["group","amount"]))
        lg = bs.get("liabilities_by_group", pd.DataFrame(columns=["group","amount"]))

        with col_a:
            st.markdown("### AKTYWA (A)")
            if not ag.empty:
                st.dataframe(
                    ag.style.format({"amount": "{:,.2f}"}),
                    use_container_width=True,
                    height=460,
                )
            st.markdown(f"**Razem Aktywa: {fmt(ta)}**")

        with col_p:
            st.markdown("### PASYWA / KAPITAŁ (P)")
            if not lg.empty:
                st.dataframe(
                    lg.style.format({"amount": "{:,.2f}"}),
                    use_container_width=True,
                    height=460,
                )
            st.markdown(f"**Razem Pasywa: {fmt(tl)}**")

        st.markdown("---")
        st.markdown("### Wykres struktury")
        ch1, ch2 = st.columns(2)
        with ch1:
            st.plotly_chart(assets_pie(ag), use_container_width=True)
        with ch2:
            st.plotly_chart(liabilities_pie(lg), use_container_width=True)
        st.plotly_chart(balance_bar(ta, tl), use_container_width=True)

        with st.expander("ℹ️ O kolejności pozycji"):
            st.markdown("""
Pozycje Balance Sheet są prezentowane w kolejności z pliku referencyjnego
`data/bs_order.xlsx` (arkusz `BS`, kolumna B) — kolejność Hyperion/CCL.
Pozycje spoza listy referencyjnej dołączane są na końcu.
**Brak sortowania po kwocie.**
""")

# ════════════════════════════════════════════════════════════════════════════
# P&L VIEW
# ════════════════════════════════════════════════════════════════════════════
elif section == "📉 P&L":
    st.markdown(f"## 📉 P&L — Rachunek Wyników — {period_str}")

    if not st.session_state.get("analyzed"):
        _no_data()
    else:
        pnl    = st.session_state["pnl"]
        pnl_df = pnl.get("pnl_df", pd.DataFrame())
        nr     = pnl.get("net_result", 0.0)

        # Net result
        c1, c2, c3 = st.columns(3)
        income  = pnl_df[pnl_df["persaldo_pnl"] > 0]["persaldo_pnl"].sum() if not pnl_df.empty else 0.0
        expense = pnl_df[pnl_df["persaldo_pnl"] < 0]["persaldo_pnl"].sum() if not pnl_df.empty else 0.0

        c1.metric("Przychody (+)", fmt(income))
        c2.metric("Koszty (−)",    fmt(expense))
        c3.metric(
            "Wynik Netto",
            fmt(nr),
            delta="Zysk ✓" if nr > 0 else "Strata ✗",
        )

        st.markdown("---")

        # Rules
        with st.expander("ℹ️ Reguły kwalifikacji kont P&L"):
            st.markdown("""
| Kryterium | Wartość |
|---|---|
| Konta 3-znakowe zaczynające się od **4** | ✓ (bez 409, 490) |
| Konta 3-znakowe zaczynające się od **7** | ✓ (bez 409, 490) |
| Konto **409** | ✗ wykluczone |
| Konto **490** | ✗ wykluczone |
| Konto **870** | ✓ dodatkowe |
| Konto **590** | ✓ dodatkowe |
| Persaldo | `Saldo Ct − Saldo Dt`  (+  = przychód, − = koszt) |
""")

        if pnl_df.empty:
            st.warning("Brak kont P&L po zastosowaniu reguł.")
        else:
            st.plotly_chart(pnl_waterfall(pnl_df, nr), use_container_width=True)

            st.markdown("### Konta P&L")
            cols = [c for c in
                    ["account_number","account_name","saldo_dt","saldo_ct",
                     "persaldo_pnl","pnl_type"]
                    if c in pnl_df.columns]
            st.dataframe(
                pnl_df[cols].style.format({
                    c: "{:,.2f}" for c in
                    ["saldo_dt","saldo_ct","persaldo_pnl"]
                    if c in pnl_df.columns
                }),
                use_container_width=True,
                height=500,
            )
