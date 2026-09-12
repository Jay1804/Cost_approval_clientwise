# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A Streamlit app that generates client-wise "Cost Approval Tracker" Excel reports directly from the
`checkpoint_live` production MySQL database (AWS RDS). It replaces an older workflow that manually
exported CSVs through a web-based Query Browser via Selenium automation — that flow is gone; this app
queries the database directly.

## Commands

```bash
pip install -r requirements.txt

# Run the app (pick a free port; 8501 may already be in use by another instance)
python -m streamlit run cost_app1.py --server.port 8502
```

There is no test suite, linter, or build step configured in this repo.

## Architecture

Three files, each with a single responsibility:

- **`sql_queries.py`** — Raw SQL templates only, no execution logic.
  - `CLIENT_LIST_QUERY`: lightweight distinct client id/name list used to populate the client picker.
    It mirrors the same `WHERE` filters as the main query so the dropdown only shows clients that
    actually have cost-approval insufficiencies.
  - `get_case_data_query(client_ids)`: the main "CS Hub -> Client wise all cases..." report query.
  - `get_flexi_field_query(client_ids)`: maps each client's `CASE_FLEX_FIELDn` columns to their
    human-readable field names (per-client custom fields).
  - Both `get_*_query` functions accept an optional list of client ids and append an
    `AND client_id IN (...)` clause. IDs are cast to `int()` before string interpolation (not bound
    params) — safe because only integers can pass, but don't relax that validation.
  - `EXCLUDED_CLIENT_IDS`: hardcoded internal/test client ids excluded from every query.

- **`db_utils.py`** — DB connection layer. Reads `DB_HOST`/`DB_PORT`/`DB_USER`/`DB_PASSWORD`/`DB_NAME`
  from `.env` via `python-dotenv`, builds a SQLAlchemy engine (`mysql+pymysql://...`), and exposes
  `fetch_client_list`, `fetch_case_data`, `fetch_flexi_fields`.
  - **Important gotcha**: queries must be wrapped in `sqlalchemy.text()` before `pd.read_sql`.
    PyMySQL's cursor tries to `%`-format the raw SQL string when no `text()` wrapper is used, and the
    literal `%` in this project's `LIKE '%dummy%'` / `LIKE '%cost%'` clauses gets misread as a printf
    format specifier (e.g. `%d`), raising `TypeError: %d format: a real number is required, not dict`.
    Any new raw-SQL query added here must go through `text()` the same way.

- **`cost_app1.py`** — Streamlit UI + all business logic, in three stages:
  1. **Cost extraction** (`extract_cost_fast` + `PATTERNS`/`APPROVAL_PATTERN`/`TOTAL_PATTERN`): a
     regex cascade that pulls a monetary amount out of the free-text `insuff_remarks` field. Rows
     with no detectable amount are dropped — this is the primary filter that turns "all insufficiency
     remarks" into "cost-approval cases".
  2. **Ageing** (`calculate_networkdays` + `get_bucket_from_ageing`): business-day age of each case
     since `Insuff Raised Date`, using the hardcoded `HOLIDAYS` list (update this yearly), bucketed
     into `0-5` / `06-10` / ... / `30+`.
  3. **Per-client tracker export** (`build_client_trackers`): groups the processed rows by `client_id`,
     renames that client's `CASE_FLEX_FIELDn` columns using the flexi-field mapping, and writes a
     two-sheet workbook per client (`Data` + `Bucket Summary` pivot). `beautify_excel()` then applies
     the visual styling — title band, navy header, zebra striping, ageing-bucket colour scale, bold
     totals, autofilter/freeze panes — to every sheet in the workbook.

  UI flow: load client list -> pick "All Clients" or specific ones (filtering happens in SQL, not
  in-memory, for efficiency) -> generate -> download a single `.xlsx` (one client) or a `.zip` of all
  trackers (multiple clients). Engine and client list are cached in `st.session_state`. There's also
  an optional "save a copy to a local folder" checkbox that writes the generated files to disk
  (defaults to `C:\Cost_Approval`) in addition to the download button.

  Email sending and Selenium-based extraction from an earlier version of this app were intentionally
  removed — there's currently no automated email delivery; trackers are downloaded manually.

  **Gotcha**: `process_case_data()` and `create_pivot_table()` key into the DataFrame by the exact
  column aliases the SQL queries produce (`insuff_remarks`, `'Insuff Raised Date'`, `Check_unique_name`,
  `Case_Check_id`, `Bucket`). If you rename a column alias in `sql_queries.py`, update these functions
  too — the mismatch fails silently (`KeyError` at runtime, or a dropped/empty pivot), not at
  query-build time.
