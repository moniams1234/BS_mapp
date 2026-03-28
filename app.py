"""
app.py – Financial Analyzer
Main Streamlit application.

Run: python -m streamlit run app.py
"""
from __future__ import annotations

import io
import json
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st

from modules.xlsx_parser import parse_xlsx
from modules.mapping_engine import run_mapping, compute_balance_sheet
from modules.anomaly_detection import build_red_flags
from modules.balance_sheet import build_balance_sheet_table
from modules.charts import (
    balance_sheet_waterfall,
    assets_breakdown_pie,
    liabilities_breakdown_pie,
    mapp_group_bar,
    top_accounts_bar,
    mapping_status_donut,
)
from modules.export_utils import build_excel_export, build_json_export
from modules.ai_analysis import (
    offline_cfo_summary,
    offline_board_memo,
    offline_nl_query,
    try_llm_query,
)
from modules.utils import (
    fmt_currency,
    render_flag_section,
    save_feedback,
    empty_state_message,
    check_session_data,
    get_session,
)

# ─────────────────────────── page config ─────────────────────────

st.set_page_config(
    page_title="Financial Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────── CSS ─────────────────────────────────

st.markdown("""
<style>
  [data-testid="stSidebar"] {background: #0F172A;}
  [data-testid="stSidebar"] * {color: #E2E8F0 !important;}
  .kpi-card {
    background: #1E293B;
    border-radius: 10px;
    padding: 16px 20px;
    text-align: center;
    border-left: 4px solid #2563EB;
  }
  .kpi-label {font-size: 0.8rem; color: #94A3B8; text-transform: uppercase; letter-spacing: 0.05em;}
  .kpi-value {font-size: 1.8rem; font-weight: 700; color: #F1F5F9; font-family: monospace;}
  .section-header {
    background: linear-gradient(90deg, #1E3A5F, #0F172A);
    padding: 12px 20px;
    border-radius: 8px;
    margin-bottom: 16px;
  }
  .flag-error {background: #450A0A; border-left: 4px solid #DC2626; padding: 8px 12px; border-radius: 4px; margin: 4px 0;}
  .flag-warning {background: #451A03; border-left: 4px solid #D97706; padding: 8px 12px; border-radius: 4px; margin: 4px 0;}
  .flag-success {background: #052E16; border-left: 4px solid #059669; padding: 8px 12px; border-radius: 4px; margin: 4px 0;}
  div[data-testid="metric-container"] {background: #1E293B; border-radius: 8px; padding: 12px; border-left: 3px solid #2563EB;}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────── sidebar ─────────────────────────────

SECTIONS = [
    "📈 XML Analysis",
    "💬 CFO Chat",
    "📝 Board Memo LLM",
    "🔍 Anomaly Detection",
    "🗣️ NL Query",
    "🗺️ Mapp",
    "⚖️ Balance Sheet",
    "📦 Batch Processing",
    "💡 User Feedback",
]

with st.sidebar:
    st.markdown("## 📊 Financial Analyzer")
    st.markdown("---")

    uploaded_file = st.file_uploader(
        "Upload Trial Balance (.xlsx)",
        type=["xlsx"],
        key="main_upload",
    )

    st.markdown("---")
    section = st.radio("Navigation", SECTIONS, label_visibility="collapsed")
    st.markdown("---")

    # Analyze button
    analyze_clicked = st.button("▶ Analyze", use_container_width=True, type="primary")

    # Optional mapping file
    with st.expander("⚙️ Advanced"):
        mapping_file_upload = st.file_uploader(
            "Custom Mapping JSON (optional)",
            type=["json"],
            key="mapping_upload",
        )
        st.caption("Leave empty to use built-in mapping.")

    st.markdown("---")
    st.caption("Financial Analyzer v1.0")


# ─────────────────────────── analyze pipeline ────────────────────

def run_analysis(file_obj, mapping_path: str | None = None):
    """Full analysis pipeline. Stores results in st.session_state."""
    with st.spinner("Parsing XLSX…"):
        parse_result = parse_xlsx(file_obj)

    if "error" in parse_result:
        st.error(f"❌ Parse error: {parse_result['error']}")
        return

    df = parse_result["df"]
    warnings = parse_result.get("warnings", [])
    sheet_used = parse_result.get("sheet_used", "?")

    with st.spinner("Running mapping engine…"):
        mapp_df = run_mapping(df, mapping_path)

    with st.spinner("Building balance sheet…"):
        bs_raw = compute_balance_sheet(mapp_df)
        bs = build_balance_sheet_table(mapp_df)

    with st.spinner("Computing red flags…"):
        flags = build_red_flags(df, mapp_df, bs_raw, warnings)

    # KPIs
    total_assets = bs.get("total_assets", 0)
    total_liab = bs.get("total_liabilities", 0)
    total_equity = bs.get("total_equity", 0)
    net_position = total_assets - total_liab  # assets minus debt liabilities

    # Store in session_state
    st.session_state["analyzed"] = True
    st.session_state["df"] = df
    st.session_state["mapp_df"] = mapp_df
    st.session_state["bs"] = bs
    st.session_state["bs_raw"] = bs_raw
    st.session_state["flags"] = flags
    st.session_state["sheet_used"] = sheet_used
    st.session_state["filename"] = getattr(file_obj, "name", "uploaded_file.xlsx")
    st.session_state["kpi"] = {
        "total_assets": total_assets,
        "total_liabilities": total_liab,
        "total_equity": total_equity,
        "net_position": net_position,
        "n_accounts": len(df),
        "n_mapped": len(mapp_df[mapp_df["mapping_status"] == "mapped"]) if not mapp_df.empty else 0,
        "n_heuristic": len(mapp_df[mapp_df["mapping_status"] == "heuristic"]) if not mapp_df.empty else 0,
        "difference": bs.get("difference", 0),
    }

    st.success(f"✅ Analysis complete. Sheet used: **{sheet_used}** | {len(df)} accounts processed.")


# ─────────────────────────── trigger analysis ────────────────────

if analyze_clicked:
    if uploaded_file is None:
        st.sidebar.error("Please upload an XLSX file first.")
    else:
        mapping_path = None
        if mapping_file_upload is not None:
            tmp = Path("/tmp/custom_mapping.json")
            tmp.write_bytes(mapping_file_upload.read())
            mapping_path = str(tmp)
        run_analysis(uploaded_file, mapping_path)


# ─────────────────────────── view routing ────────────────────────

section_key = section.split(" ", 1)[1].strip()  # strip emoji

# ══════════════════════════════════════════════════════════════════
# 1. XML Analysis (main dashboard)
# ══════════════════════════════════════════════════════════════════

if section_key == "XML Analysis":
    st.markdown("## 📈 Financial Dashboard")

    if not check_session_data():
        empty_state_message("Upload a trial balance XLSX and click ▶ Analyze to begin.")
        st.markdown("""
        ### How to use
        1. Upload your **XLSX trial balance** (ZOiS / zestawienie obrotów i sald) in the sidebar.
        2. Click **▶ Analyze** to process the file.
        3. Navigate sections using the sidebar.
        4. Export to Excel or JSON using the buttons below.

        ### Supported formats
        - Polish ZOiS: Numer, Nazwa, BO Dt, BO Ct, Obroty Dt, Obroty Ct, Saldo Dt, Saldo Ct, Persaldo
        - English Trial Balance: Account, Account Name, Debit, Credit, Opening Balance, Closing Balance
        - Mixed / custom: Column names are detected heuristically
        """)
    else:
        kpi = get_session("kpi", {})
        flags = get_session("flags", [])
        mapp_df = get_session("mapp_df")
        bs = get_session("bs", {})
        filename = get_session("filename", "")

        # KPI row
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Assets", fmt_currency(kpi.get("total_assets", 0)))
        with col2:
            st.metric("Total Liabilities", fmt_currency(kpi.get("total_liabilities", 0)))
        with col3:
            st.metric("Total Equity", fmt_currency(kpi.get("total_equity", 0)))
        with col4:
            diff = kpi.get("difference", 0)
            st.metric("Balance Diff", fmt_currency(diff),
                      delta="Balanced ✓" if abs(diff) < 1 else f"⚠ {fmt_currency(abs(diff))}")

        col5, col6, col7 = st.columns(3)
        with col5:
            st.metric("Net Position (Assets − Liab)", fmt_currency(kpi.get("net_position", 0)))
        with col6:
            st.metric("Total Accounts", f"{kpi.get('n_accounts', 0):,}")
        with col7:
            total = kpi.get("n_accounts", 1) or 1
            mapped = kpi.get("n_mapped", 0) + kpi.get("n_heuristic", 0)
            st.metric("Mapping Coverage", f"{mapped/total*100:.1f}%")

        st.markdown("---")

        # Charts row 1
        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(balance_sheet_waterfall(get_session("bs_raw", {})), use_container_width=True)
        with c2:
            st.plotly_chart(mapping_status_donut(mapp_df), use_container_width=True)

        # Charts row 2
        c3, c4 = st.columns(2)
        with c3:
            st.plotly_chart(assets_breakdown_pie(bs.get("assets_detail")), use_container_width=True)
        with c4:
            st.plotly_chart(liabilities_breakdown_pie(bs.get("liabilities_detail")), use_container_width=True)

        # Top accounts
        st.plotly_chart(top_accounts_bar(mapp_df), use_container_width=True)

        # Red flags
        st.markdown("### 🚩 Red Flags")
        render_flag_section(flags)

        # Export
        st.markdown("---")
        st.markdown("### 💾 Export")
        col_e1, col_e2 = st.columns(2)

        with col_e1:
            xlsx_bytes = build_excel_export(
                get_session("df"), mapp_df, bs, flags, filename
            )
            st.download_button(
                "⬇ Download Excel",
                data=xlsx_bytes,
                file_name=f"financial_analysis_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with col_e2:
            json_str = build_json_export(get_session("df"), mapp_df, bs, flags)
            st.download_button(
                "⬇ Download JSON",
                data=json_str.encode("utf-8"),
                file_name=f"financial_analysis_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
                mime="application/json",
                use_container_width=True,
            )

        # Raw data expander
        with st.expander("📋 Raw Trial Balance (first 100 rows)"):
            df = get_session("df")
            if df is not None:
                st.dataframe(df.head(100), use_container_width=True)


# ══════════════════════════════════════════════════════════════════
# 2. CFO Chat
# ══════════════════════════════════════════════════════════════════

elif section_key == "CFO Chat":
    st.markdown("## 💬 CFO Chat")

    if not check_session_data():
        empty_state_message()
    else:
        bs = get_session("bs", {})
        flags = get_session("flags", [])
        filename = get_session("filename", "")

        summary = offline_cfo_summary(bs, flags, filename)

        # LLM enhancement if available
        llm_resp = try_llm_query(
            f"You are a CFO. Analyze this balance sheet data and provide insights:\n{summary}"
        )

        if llm_resp and not llm_resp.startswith("[LLM error"):
            st.markdown("### 🤖 AI-Enhanced Analysis")
            st.markdown(llm_resp)
        else:
            st.markdown("### 📊 Offline CFO Analysis")
            st.markdown(summary)
            if llm_resp and llm_resp.startswith("[LLM error"):
                st.caption(f"LLM unavailable: {llm_resp}")

        st.markdown("---")
        st.markdown("### Chat")
        if "cfo_chat_history" not in st.session_state:
            st.session_state["cfo_chat_history"] = []

        for msg in st.session_state["cfo_chat_history"]:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

        if prompt := st.chat_input("Ask a financial question…"):
            st.session_state["cfo_chat_history"].append({"role": "user", "content": prompt})
            with st.chat_message("user"):
                st.markdown(prompt)

            mapp_df = get_session("mapp_df")
            answer = offline_nl_query(prompt, bs, mapp_df)
            llm_ans = try_llm_query(f"Financial context: {summary}\n\nQuestion: {prompt}")
            final_answer = llm_ans if (llm_ans and not llm_ans.startswith("[LLM error")) else answer

            st.session_state["cfo_chat_history"].append({"role": "assistant", "content": final_answer})
            with st.chat_message("assistant"):
                st.markdown(final_answer)


# ══════════════════════════════════════════════════════════════════
# 3. Board Memo LLM
# ══════════════════════════════════════════════════════════════════

elif section_key == "Board Memo LLM":
    st.markdown("## 📝 Board Memo")

    if not check_session_data():
        empty_state_message()
    else:
        bs = get_session("bs", {})
        flags = get_session("flags", [])
        filename = get_session("filename", "")

        memo = offline_board_memo(bs, flags, filename)

        llm_prompt = f"""You are a senior financial analyst preparing a board memo.
Based on this data, write a professional board memo:\n{memo}"""
        llm_resp = try_llm_query(llm_prompt)

        col1, col2 = st.columns([3, 1])
        with col1:
            if llm_resp and not llm_resp.startswith("[LLM error"):
                st.markdown("*🤖 AI-Enhanced Memo*")
                st.markdown(llm_resp)
            else:
                st.markdown(memo)
        with col2:
            st.download_button(
                "⬇ Download Memo (.md)",
                data=(llm_resp or memo).encode("utf-8"),
                file_name=f"board_memo_{datetime.now().strftime('%Y%m%d')}.md",
                mime="text/markdown",
                use_container_width=True,
            )


# ══════════════════════════════════════════════════════════════════
# 4. Anomaly Detection
# ══════════════════════════════════════════════════════════════════

elif section_key == "Anomaly Detection":
    st.markdown("## 🔍 Anomaly Detection & Red Flags")

    if not check_session_data():
        empty_state_message()
    else:
        flags = get_session("flags", [])
        mapp_df = get_session("mapp_df")

        # Summary counters
        n_errors = sum(1 for f in flags if f["type"] == "error")
        n_warnings = sum(1 for f in flags if f["type"] == "warning")
        n_ok = sum(1 for f in flags if f["type"] == "success")

        c1, c2, c3 = st.columns(3)
        c1.metric("🔴 Errors", n_errors)
        c2.metric("🟡 Warnings", n_warnings)
        c3.metric("🟢 OK", n_ok)

        st.markdown("---")
        render_flag_section(flags)

        st.markdown("---")
        st.markdown("### Mapping Coverage Detail")
        if mapp_df is not None and not mapp_df.empty:
            status_summary = mapp_df.groupby("mapping_status").agg(
                count=("account_number", "count"),
                total_persaldo=("persaldo", "sum"),
            ).reset_index()
            st.dataframe(status_summary, use_container_width=True)

            st.markdown("### Heuristically Mapped Accounts")
            heuristic = mapp_df[mapp_df["mapping_status"] == "heuristic"]
            if not heuristic.empty:
                st.dataframe(
                    heuristic[["account_number", "account_name", "group", "side", "persaldo"]],
                    use_container_width=True,
                )
            else:
                st.success("All accounts have explicit mapping entries.")

        # Numerical anomalies
        st.markdown("### Numerical Anomalies")
        if mapp_df is not None and not mapp_df.empty:
            anomalies = mapp_df[
                (mapp_df["debit"] > 0) & (mapp_df["credit"] > 0) &
                ((mapp_df["debit"] / mapp_df["credit"].replace(0, np.nan)).abs() > 10)
            ]
            if not anomalies.empty:
                st.warning(f"{len(anomalies)} accounts with highly skewed Debit/Credit ratio:")
                st.dataframe(anomalies[["account_number", "account_name", "debit", "credit", "persaldo"]])
            else:
                st.success("No extreme debit/credit skew detected.")


# ══════════════════════════════════════════════════════════════════
# 5. NL Query
# ══════════════════════════════════════════════════════════════════

elif section_key == "NL Query":
    st.markdown("## 🗣️ Natural Language Query")

    if not check_session_data():
        empty_state_message()
    else:
        st.markdown("""
        Ask questions about the trial balance data in plain English or Polish.
        Examples:
        - *What are the total assets?*
        - *Show me the top 5 accounts by balance.*
        - *How many accounts were not mapped?*
        - *Is the balance sheet balanced?*
        """)

        query = st.text_input("Your question:", placeholder="e.g. What is the total equity?")

        if st.button("Ask", type="primary") and query:
            bs = get_session("bs", {})
            mapp_df = get_session("mapp_df")
            summary = offline_cfo_summary(bs, get_session("flags", []), get_session("filename", ""))

            offline_ans = offline_nl_query(query, bs, mapp_df)
            llm_ans = try_llm_query(
                f"Financial data context:\n{summary}\n\nAnswer this question concisely: {query}"
            )

            if llm_ans and not llm_ans.startswith("[LLM error"):
                st.markdown("**🤖 AI Answer:**")
                st.markdown(llm_ans)
            else:
                st.markdown("**Answer:**")
                st.markdown(offline_ans)

        st.markdown("---")
        st.markdown("### Quick Stats")
        bs = get_session("bs", {})
        kpi = get_session("kpi", {})
        if kpi:
            data = {
                "Metric": ["Total Assets", "Total Liabilities", "Total Equity", "Balance Diff", "Accounts"],
                "Value": [
                    fmt_currency(kpi.get("total_assets", 0)),
                    fmt_currency(kpi.get("total_liabilities", 0)),
                    fmt_currency(kpi.get("total_equity", 0)),
                    fmt_currency(kpi.get("difference", 0)),
                    str(kpi.get("n_accounts", 0)),
                ],
            }
            st.table(pd.DataFrame(data))


# ══════════════════════════════════════════════════════════════════
# 6. Mapp
# ══════════════════════════════════════════════════════════════════

elif section_key == "Mapp":
    st.markdown("## 🗺️ Mapp – Account Mapping Table")

    if not check_session_data():
        empty_state_message()
    else:
        mapp_df = get_session("mapp_df")

        if mapp_df is None or mapp_df.empty:
            st.warning("Mapping table is empty.")
        else:
            # Filters
            col1, col2, col3 = st.columns(3)
            with col1:
                side_filter = st.multiselect("Side", ["A", "P", "R"], default=["A", "P", "R"])
            with col2:
                group_filter = st.multiselect(
                    "Group", sorted(mapp_df["group"].unique()), default=[]
                )
            with col3:
                status_filter = st.multiselect(
                    "Mapping Status",
                    mapp_df["mapping_status"].unique().tolist(),
                    default=mapp_df["mapping_status"].unique().tolist(),
                )

            filtered = mapp_df[mapp_df["side"].isin(side_filter)]
            if group_filter:
                filtered = filtered[filtered["group"].isin(group_filter)]
            filtered = filtered[filtered["mapping_status"].isin(status_filter)]

            # Display columns
            display_cols = [
                "side", "account_number", "account_name", "group",
                "debit", "credit", "saldo_dt", "saldo_ct",
                "persaldo", "persaldo_group", "mapping_status",
            ]
            display_cols = [c for c in display_cols if c in filtered.columns]

            st.markdown(f"**Showing {len(filtered):,} of {len(mapp_df):,} accounts**")
            st.dataframe(
                filtered[display_cols].style.format({
                    "debit": "{:,.2f}",
                    "credit": "{:,.2f}",
                    "saldo_dt": "{:,.2f}",
                    "saldo_ct": "{:,.2f}",
                    "persaldo": "{:,.2f}",
                    "persaldo_group": "{:,.2f}",
                }),
                use_container_width=True,
                height=500,
            )

            # Group summary
            st.markdown("### Group Totals")
            group_summary = (
                filtered.groupby(["side", "group"])
                .agg(
                    accounts=("account_number", "count"),
                    total_debit=("debit", "sum"),
                    total_credit=("credit", "sum"),
                    total_persaldo=("persaldo", "sum"),
                )
                .reset_index()
                .sort_values(["side", "total_persaldo"], ascending=[True, False])
            )
            st.dataframe(
                group_summary.style.format({
                    "total_debit": "{:,.2f}",
                    "total_credit": "{:,.2f}",
                    "total_persaldo": "{:,.2f}",
                }),
                use_container_width=True,
            )

            # Chart
            st.plotly_chart(mapp_group_bar(filtered), use_container_width=True)

            # Persaldo explanation
            with st.expander("📖 Persaldo calculation explained"):
                st.markdown("""
                **Persaldo sign convention:**

                | Side | Formula | Positive means |
                |---|---|---|
                | **A** (Assets) | `Saldo Dt − Saldo Ct` | Debit balance (normal for assets) |
                | **P** (Liabilities & Equity) | `Saldo Ct − Saldo Dt` | Credit balance (normal for liabilities) |
                | **R** (P&L) | `Saldo Ct − Saldo Dt` | Net income (credit side) |

                **Persaldo per group** = sum of persaldo for all accounts within the same reporting group.
                """)


# ══════════════════════════════════════════════════════════════════
# 7. Balance Sheet
# ══════════════════════════════════════════════════════════════════

elif section_key == "Balance Sheet":
    st.markdown("## ⚖️ Balance Sheet")

    if not check_session_data():
        empty_state_message()
    else:
        bs = get_session("bs", {})
        bs_raw = get_session("bs_raw", {})

        if not bs:
            st.warning("No balance sheet data available.")
        else:
            total_assets = bs.get("total_assets", 0)
            total_liab = bs.get("total_liabilities", 0)
            total_equity = bs.get("total_equity", 0)
            total_le = bs.get("total_liab_equity", 0)
            diff = bs.get("difference", 0)

            # Summary KPIs
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total Assets", fmt_currency(total_assets))
            c2.metric("Total Liabilities", fmt_currency(total_liab))
            c3.metric("Total Equity", fmt_currency(total_equity))
            c4.metric(
                "Balance Difference",
                fmt_currency(diff),
                delta="✓ Balanced" if abs(diff) < 1 else "⚠ Unbalanced",
            )

            st.markdown("---")

            col_a, col_p = st.columns(2)

            with col_a:
                st.markdown("### ASSETS")
                assets_df = bs.get("assets_detail")
                if assets_df is not None and not assets_df.empty:
                    st.dataframe(
                        assets_df.style.format({"Amount": "{:,.2f}"}),
                        use_container_width=True,
                    )
                    st.markdown(f"**Total Assets: {fmt_currency(total_assets)}**")

            with col_p:
                st.markdown("### LIABILITIES")
                liab_df = bs.get("liabilities_detail")
                if liab_df is not None and not liab_df.empty:
                    st.dataframe(
                        liab_df.style.format({"Amount": "{:,.2f}"}),
                        use_container_width=True,
                    )
                    st.markdown(f"**Total Liabilities: {fmt_currency(total_liab)}**")

                st.markdown("### EQUITY")
                eq_df = bs.get("equity_detail")
                if eq_df is not None and not eq_df.empty:
                    st.dataframe(
                        eq_df.style.format({"Amount": "{:,.2f}"}),
                        use_container_width=True,
                    )
                    st.markdown(f"**Total Equity: {fmt_currency(total_equity)}**")

                st.markdown(f"**Total Liabilities + Equity: {fmt_currency(total_le)}**")
                if abs(diff) < 1:
                    st.success("Balance sheet is balanced ✓")
                else:
                    st.error(f"Balance sheet difference: {fmt_currency(diff)}")

            st.markdown("---")

            # P&L section
            st.markdown("### P&L Accounts (R-side)")
            pl_df = bs.get("pl_detail")
            if pl_df is not None and not pl_df.empty:
                st.dataframe(pl_df.style.format({"Amount": "{:,.2f}"}), use_container_width=True)
                st.markdown(f"**Net P&L: {fmt_currency(bs.get('total_pl', 0))}**")
            else:
                st.info("No P&L accounts found in current mapping.")

            st.markdown("---")

            # Charts
            c_left, c_right = st.columns(2)
            with c_left:
                st.plotly_chart(assets_breakdown_pie(bs.get("assets_detail")), use_container_width=True)
            with c_right:
                st.plotly_chart(liabilities_breakdown_pie(bs.get("liabilities_detail")), use_container_width=True)

            st.plotly_chart(balance_sheet_waterfall(bs_raw), use_container_width=True)

            # Export
            st.markdown("---")
            xlsx_bytes = build_excel_export(
                get_session("df"),
                get_session("mapp_df"),
                bs,
                get_session("flags", []),
                get_session("filename", ""),
            )
            st.download_button(
                "⬇ Export Balance Sheet (Excel)",
                data=xlsx_bytes,
                file_name=f"balance_sheet_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


# ══════════════════════════════════════════════════════════════════
# 8. Batch Processing
# ══════════════════════════════════════════════════════════════════

elif section_key == "Batch Processing":
    st.markdown("## 📦 Batch Processing")
    st.markdown("Upload multiple XLSX files and process them together.")

    batch_files = st.file_uploader(
        "Upload XLSX files (multiple)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="batch_upload",
    )

    if st.button("▶ Process Batch", type="primary") and batch_files:
        results = []
        progress = st.progress(0)
        status_box = st.empty()

        for i, f in enumerate(batch_files):
            status_box.info(f"Processing {f.name}… ({i+1}/{len(batch_files)})")
            try:
                parse_result = parse_xlsx(f)
                if "error" in parse_result:
                    results.append({
                        "filename": f.name,
                        "status": "error",
                        "error": parse_result["error"],
                        "accounts": 0,
                        "total_assets": None,
                        "total_liabilities": None,
                        "difference": None,
                    })
                    continue

                df = parse_result["df"]
                mapp_df = run_mapping(df)
                bs = build_balance_sheet_table(mapp_df)

                results.append({
                    "filename": f.name,
                    "status": "ok",
                    "error": "",
                    "accounts": len(df),
                    "total_assets": round(bs.get("total_assets", 0), 2),
                    "total_liabilities": round(bs.get("total_liabilities", 0), 2),
                    "total_equity": round(bs.get("total_equity", 0), 2),
                    "difference": round(bs.get("difference", 0), 2),
                })
            except Exception as e:
                results.append({
                    "filename": f.name,
                    "status": "error",
                    "error": str(e),
                    "accounts": 0,
                    "total_assets": None,
                    "total_liabilities": None,
                    "difference": None,
                })

            progress.progress((i + 1) / len(batch_files))

        status_box.success(f"✅ Batch complete: {len(batch_files)} files processed.")
        results_df = pd.DataFrame(results)
        st.dataframe(results_df, use_container_width=True)

        # Download batch summary
        csv = results_df.to_csv(index=False)
        st.download_button(
            "⬇ Download Batch Summary (CSV)",
            data=csv.encode("utf-8"),
            file_name=f"batch_summary_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
        )

        st.session_state["batch_results"] = results

    elif "batch_results" in st.session_state:
        st.markdown("### Previous Batch Results")
        st.dataframe(pd.DataFrame(st.session_state["batch_results"]), use_container_width=True)

    if not batch_files:
        st.info("Upload one or more XLSX files above and click ▶ Process Batch.")


# ══════════════════════════════════════════════════════════════════
# 9. User Feedback
# ══════════════════════════════════════════════════════════════════

elif section_key == "User Feedback":
    st.markdown("## 💡 User Feedback")
    st.markdown("Help us improve by sharing your experience.")

    with st.form("feedback_form"):
        rating = st.slider("Rating", 1, 5, 4, help="1 = Very poor, 5 = Excellent")
        stars = "⭐" * rating
        st.markdown(f"**Your rating: {stars} ({rating}/5)**")

        category = st.selectbox(
            "Feedback category",
            ["General", "Data Parsing", "Mapping Quality", "UI/UX", "Export", "Performance", "Other"],
        )
        comment = st.text_area("Comments / Suggestions", placeholder="What worked well? What could be improved?")
        submit = st.form_submit_button("Submit Feedback", type="primary")

        if submit:
            if comment.strip() or rating:
                save_feedback(rating, f"[{category}] {comment}")
                st.success("✅ Thank you for your feedback!")
            else:
                st.error("Please provide a rating or comment.")

    # Display past feedback
    feedback_log = st.session_state.get("feedback_log", [])
    if feedback_log:
        st.markdown("---")
        st.markdown("### Session Feedback History")
        for entry in reversed(feedback_log[-10:]):
            ts = entry.get("timestamp", "")[:19]
            r = entry.get("rating", 0)
            c = entry.get("comment", "")
            st.markdown(f"**{ts}** – {'⭐' * r} – {c}")
    else:
        st.caption("No feedback submitted yet in this session.")
