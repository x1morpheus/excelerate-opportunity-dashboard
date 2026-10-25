# Excelerate Opportunity Intelligence — Streamlit Dashboard

A professional, filterable Streamlit dashboard for the Excelerate opportunity dataset. The dashboard follows the EDA/Power BI plan: a concise overview, fee and approval analysis, data-quality checks, and a searchable/exportable record table.

## Files
- `app.py` — dashboard source code
- `requirements.txt` — Python dependencies
- `excelerate_cleaned.xlsx` — default workbook used by the app

## Run locally
1. Install Python 3.10+.
2. Put these files in the same folder.
3. Open a terminal in that folder and run:

   ```bash
   python -m venv .venv
   # Windows:
   .venv\Scripts\activate
   # macOS/Linux:
   source .venv/bin/activate
   pip install -r requirements.txt
   streamlit run app.py
   ```

4. The dashboard opens in your browser. You can upload another `.xlsx` workbook from the sidebar.

## Dashboard design plan
1. **Overview** — KPI cards, category ranking, monthly creation activity, location/work-mode distribution.
2. **Fees & approval** — fee status counts and mix by category, auto-approval rates using only known approval settings.
3. **Data quality** — missingness, duplicate ID checks, extreme fee/duration checks and deadline-before-creation checks.
4. **Explore records** — search, inspect and export the filtered records.

## Data interpretation guardrails
- `opportunity_id` is used as the distinct opportunity key.
- A zero fee is not the same as a missing fee.
- Raw fee totals and averages are deliberately omitted because USD, INR and EUR coexist and the EDA identified an extreme fee value that requires validation.
- Auto-approval rate excludes unknown values from its denominator.
- The date trend describes record creation, not necessarily underlying demand.
- High-missingness fields are shown in the data-quality view rather than promoted as headline metrics.

## Source snapshot note
The supplied `excelerate_cleaned.xlsx` workbook has 5,733 rows and 32 columns. The executed EDA summary elsewhere in the project described a 5,730-row, 19-column snapshot. The app computes results from whichever workbook is loaded; confirm that the workbook is the intended final version before publishing results.
