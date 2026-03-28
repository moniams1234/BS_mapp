# Financial Analyzer – Streamlit Application

A production-grade financial dashboard that reads trial balance XLSX files (ZOiS / zestawienie obrotów i sald), maps accounts to a reporting structure, produces a balance sheet view, and exports to Excel and JSON.

---

## Features

| Feature | Description |
|---|---|
| **XLSX Parsing** | Heuristic detection of sheet, header row, and column names (PL + EN) |
| **Account Mapping** | Prefix-based mapping to A/P/R groups with fallback heuristics |
| **Balance Sheet** | Structured assets / liabilities / equity view with totals |
| **Mapp Table** | Full account-level table with side, group, debit, credit, persaldo |
| **Red Flags** | Automated detection of data quality and balance issues |
| **Dashboard** | KPI cards, Plotly charts, mapping coverage |
| **CFO Chat** | Offline financial Q&A; optional LLM integration |
| **Board Memo** | Auto-generated executive memo |
| **NL Query** | Plain-language queries against analyzed data |
| **Anomaly Detection** | Outlier accounts and mapping gaps |
| **Batch Processing** | Multi-file processing with CSV summary |
| **Export** | Excel (7 sheets) and JSON export |
| **User Feedback** | In-app feedback form saved to JSON |

---

## Project Structure

```
financial_analyzer/
├── app.py                      # Main Streamlit app
├── requirements.txt
├── README.md
├── modules/
│   ├── __init__.py
│   ├── xlsx_parser.py          # Heuristic XLSX parser
│   ├── mapping_engine.py       # Account mapping + balance sheet
│   ├── anomaly_detection.py    # Red flag engine
│   ├── balance_sheet.py        # BS table builder
│   ├── export_utils.py         # Excel + JSON export
│   ├── charts.py               # Plotly chart builders
│   ├── ai_analysis.py          # Offline + optional LLM analysis
│   └── utils.py                # Shared utilities
└── sample_data/
    └── sample_mapping.json     # Built-in account mapping table
```

---

## Installation

```bash
python -m pip install -r requirements.txt
```

Requirements: Python 3.9+

---

## Running

```bash
python -m streamlit run app.py
```

Open http://localhost:8501 in your browser.

---

## Input Format – XLSX Trial Balance

The parser accepts any XLSX file with the following characteristics:

### Supported column names (Polish / English)

| Internal name | Accepted aliases |
|---|---|
| `account_number` | Numer, Konto, Account, Account Number |
| `account_name` | Nazwa, Name, Account Name, Description |
| `account_name2` | Nazwa 2, Label |
| `bo_dt` | BO Dt, Opening Debit |
| `bo_ct` | BO Ct, Opening Credit |
| `obroty_dt` | Obroty Dt, Debit, Wn, Dt |
| `obroty_ct` | Obroty Ct, Credit, Ma, Ct |
| `saldo_dt` | Saldo Dt, Closing Debit, Balance Debit |
| `saldo_ct` | Saldo Ct, Closing Credit, Balance Credit |
| `persaldo` | Persaldo, Saldo, Balance, Net Balance |
| `bs_mapp` | BS Mapp, Mapping, Group |

### Tolerances
- Header row may start at any of the first 20 rows.
- Multiple sheets: the parser scores each sheet name and picks the most likely trial balance.
- Numbers may be stored as text (e.g., "1 234,56" → 1234.56).
- Empty rows are removed automatically.

---

## Mapping Logic

The built-in mapping table is at `sample_data/sample_mapping.json`. Each entry specifies:
- `prefix`: account number prefix (e.g., "131")
- `side`: A (asset), P (liability/equity), R (P&L)
- `group`: reporting group label

**Matching order:**
1. Exact match on full account number
2. Longest prefix match (e.g., "131" matches "131-03", "131-04", etc.)
3. Numeric-prefix extraction (handles accounts like "131-03" → tries "131")
4. Heuristic fallback by numeric range (0–99, 100–199, 200–299, …)

To customize mapping:
- Edit `sample_data/sample_mapping.json` directly, or
- Upload a custom JSON file via **Advanced** in the sidebar.

---

## Persaldo Calculation

Persaldo is the **net balance** of an account, with sign convention by account side:

| Side | Formula | Meaning |
|---|---|---|
| **A** (Assets) | `Saldo Dt − Saldo Ct` | Positive = debit balance (normal for assets) |
| **P** (Liabilities & Equity) | `Saldo Ct − Saldo Dt` | Positive = credit balance (normal for liabilities) |
| **R** (P&L) | `Saldo Ct − Saldo Dt` | Positive = net income |

**Persaldo per group** = sum of persaldo for all accounts within the same reporting group.

If `Persaldo` is present in the source file, it is used directly. If all values are zero, it is recomputed from Saldo Dt / Saldo Ct.

---

## Export

### Excel (`.xlsx`)
Contains sheets:
- **Raw_Trial_Balance** – original parsed data
- **Mapp** – full account mapping table
- **Mapping** – same as Mapp (alias)
- **Balance_Sheet** – structured BS with section headers
- **Red_Flags** – all flags with type, category, message
- **Summary** – high-level KPIs and counts
- **Metadata** – generation timestamp and settings

### JSON
Single file with:
- `generated_at`
- `summary` (KPIs)
- `red_flags`
- `balance_sheet` (group summaries)
- `mapp` (full account list)

---

## Example Workflow

1. Export ZOiS from your ERP as XLSX.
2. Open Financial Analyzer (`streamlit run app.py`).
3. Upload the XLSX in the sidebar.
4. Click **▶ Analyze**.
5. Review the dashboard in **XML Analysis**.
6. Check **Anomaly Detection** for data quality issues.
7. Review account mapping in **Mapp**.
8. Inspect **Balance Sheet** for totals and balance check.
9. Export to Excel for stakeholder distribution.

---

## Business Assumptions

- Account numbers follow Polish GAAP chart of accounts (class 0–9).
- Accounts 0xx–3xx are balance sheet; 4xx–7xx are P&L; 8xx–9xx are liabilities/equity.
- Persaldo for assets is debit-positive; for liabilities it is credit-positive.
- The built-in mapping covers the most common Polish account prefixes. Unusual or company-specific sub-accounts fall back to heuristic classification.

---

## Known Limitations

- LLM features (CFO Chat AI, Board Memo AI) require an API key (`ANTHROPIC_API_KEY` or `OPENAI_API_KEY`) set as an environment variable.
- Batch processing does not persist results between sessions.
- Custom mapping JSON must follow the same schema as `sample_data/sample_mapping.json`.
- Accounts with non-numeric prefixes (e.g., pure text codes) fall back to heuristic classification.
- Inter-company eliminations are not performed automatically.
- Multi-currency trial balances are not currently supported.
