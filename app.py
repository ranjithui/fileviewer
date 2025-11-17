"""
Google Sheet Blacklist TEMPLATE + Streamlit App

-----------------------------
How to use this file
-----------------------------
1) Google Sheet template (create a sheet named `Blacklist`):
   - Column A header: Company Name
   - Column B header: Email

   Sample rows (starting row 2):
   | Company Name | Email             |
   | Google       | john@gmail.com    |
   | Meta         | mark@meta.com     |
   | IBM          | hr@ibm.com        |

2) Share the Google Sheet so it's viewable:
   - Click "Share" -> "Change to anyone with the link" -> Viewer
   - Copy the sheet URL (it looks like: https://docs.google.com/spreadsheets/d/XXX/edit#gid=0)

3) How to pass the Google Sheet link to the app:
   - Set the `BLACKLIST_SHEET_LINK` variable below to your Google Sheet URL.
   - The app will convert the link to the CSV export form automatically.

4) Run the app:
   - Install requirements: `pip install streamlit pandas openpyxl`
   - Run: `streamlit run app.py`

-----------------------------
What this Streamlit app does
-----------------------------
- You upload your main Excel file (XLSX).
- The app fetches the Blacklist (Company Name + Email) from the Google Sheet link.
- It removes rows from the uploaded sheet when EITHER Company OR Email is present in blacklist.
- The cleaned rows are sorted by Company Name (ascending).
- You can delete rows live from the UI.
- You can download the final cleaned XLSX.

"""

import streamlit as st
import pandas as pd
from io import BytesIO
import urllib.parse
import re

st.set_page_config(page_title="Data Cleaner + Live Editor", layout="wide")

# ----------------------
# CONFIG: set your Google Sheet link here
# ----------------------
# Example form:
# https://docs.google.com/spreadsheets/d/1AbcDEfgHIjKlmnopQRsTUvwXyz0123456789/edit#gid=0
BLACKLIST_SHEET_LINK = "https://docs.google.com/spreadsheets/d/11s1ImpN4piWLByRKwolTyGGILVPfR5Us4_66ceXAlLQ/export?format=csv&gid=0
"

# Unique keys (column names expected in BOTH the uploaded file and the blacklist)
KEY_COMPANY = "Company Name"  # recommended header in your uploaded sheet
KEY_EMAIL = "Email"

# Helper: convert various Google Sheet link types to CSV export URL
def sheet_link_to_csv(url: str) -> str:
    if not url:
        return ""
    # If user pasted already an export link, just return
    if "/export?" in url:
        return url
    # Try to extract /d/{id}/
    m = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
    if m:
        sheet_id = m.group(1)
        # Try to extract gid if present
        gid_match = re.search(r"[#&]gid=(\d+)", url)
        gid = gid_match.group(1) if gid_match else "0"
        return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
    # fallback: url encode and return (may fail)
    return url


def normalize_series(s: pd.Series) -> pd.Series:
    # lower, strip, and replace NaN with empty string
    return s.fillna("").astype(str).str.lower().str.strip()


st.title("📋 Data Cleaner + Live Editor")
st.markdown("Upload your main Excel file (XLSX). The app will fetch the blacklist from your Google Sheet and remove rows where either Company Name or Email appears in the blacklist.")

# ----------------------
# File uploader
# ----------------------
uploaded_file = st.file_uploader("Upload your main Excel file (XLSX)", type=["xlsx"]) 

# Preview / instructions column
with st.expander("Blacklist settings (click to view)", expanded=False):
    st.write("Provide a Google Sheet link to your blacklist. The sheet should have headers: 'Company Name' and 'Email'.")
    st.write("Set the link in the app variable BLACKLIST_SHEET_LINK in the code or enter below.")
    user_link = st.text_input("Or paste Blacklist Google Sheet link (optional)", value="" if BLACKLIST_SHEET_LINK == "REPLACE_WITH_YOUR_GOOGLE_SHEET_LINK" else BLACKLIST_SHEET_LINK)
    if user_link:
        BLACKLIST_SHEET_LINK = user_link

if not uploaded_file:
    st.info("Waiting for you to upload the main Excel file. The app expects columns: 'Company Name' and 'Email' (case-insensitive).")
    st.stop()

# Read uploaded file
try:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
except Exception as e:
    st.error(f"Unable to read uploaded Excel file: {e}")
    st.stop()

# Normalize uploaded columns: try to find matching column names case-insensitively
col_map = {c.lower(): c for c in df.columns}
found_company_col = None
found_email_col = None
for k in col_map:
    if k == KEY_COMPANY.lower():
        found_company_col = col_map[k]
    if k == KEY_EMAIL.lower():
        found_email_col = col_map[k]

if not found_company_col or not found_email_col:
    st.error(f"Uploaded file must contain columns named '{KEY_COMPANY}' and '{KEY_EMAIL}' (case-insensitive). Found columns: {list(df.columns)}")
    st.stop()

# Load blacklist from Google Sheets
csv_url = sheet_link_to_csv(BLACKLIST_SHEET_LINK)
if not csv_url:
    st.error("No valid blacklist Google Sheet link provided. Edit BLACKLIST_SHEET_LINK in the app or paste it in the settings panel.")
    st.stop()

try:
    blacklist = pd.read_csv(csv_url)
except Exception as e:
    st.error(f"Unable to fetch blacklist from Google Sheets. Check sharing settings and link. Error: {e}")
    st.stop()

# Normalize blacklist column names as well
bl_col_map = {c.lower(): c for c in blacklist.columns}
bl_company_col = None
bl_email_col = None
for k in bl_col_map:
    if k == KEY_COMPANY.lower():
        bl_company_col = bl_col_map[k]
    if k == KEY_EMAIL.lower():
        bl_email_col = bl_col_map[k]

if not bl_company_col and not bl_email_col:
    st.error(f"Blacklist sheet must contain at least one of the headers: '{KEY_COMPANY}' or '{KEY_EMAIL}'. Found columns: {list(blacklist.columns)}")
    st.stop()

# Prepare comparison sets (lowercased, stripped)
# If one of the blacklist columns is missing, create an empty series
if bl_company_col:
    blacklist_companies = set(normalize_series(blacklist[bl_company_col]))
else:
    blacklist_companies = set()

if bl_email_col:
    blacklist_emails = set(normalize_series(blacklist[bl_email_col]))
else:
    blacklist_emails = set()

# Normalize df keys
df_key_company = normalize_series(df[found_company_col])
df_key_email = normalize_series(df[found_email_col])

# Determine blacklisted rows: OR condition (company OR email)
is_company_blacklisted = df_key_company.isin(blacklist_companies)
is_email_blacklisted = df_key_email.isin(blacklist_emails)

blacklisted_mask = is_company_blacklisted | is_email_blacklisted

# Split clean and blocked
df_blocked = df[blacklisted_mask].copy()
df_clean = df[~blacklisted_mask].copy()

# Remove duplicates in clean list based on both keys
# Use normalized values to dedupe but keep original casing in result
if not df_clean.empty:
    dedupe_keys = [found_company_col, found_email_col]
    df_clean = df_clean.drop_duplicates(subset=dedupe_keys)

# Sort clean list by company name (ascending)
if not df_clean.empty:
    # Use the normalized series for sorting to ensure proper order
    df_clean['_sort_company'] = normalize_series(df_clean[found_company_col])
    df_clean = df_clean.sort_values('_sort_company').drop(columns=['_sort_company'])

# Initialize session state for editable dataframe
if 'edited_df' not in st.session_state:
    st.session_state.edited_df = df_clean.copy()

st.success(f"Loaded uploaded file: {uploaded_file.name} — {len(df)} rows. Blocked: {len(df_blocked)}. Clean: {len(df_clean)}.")

# Layout: two columns - blocked and clean/editor
col1, col2 = st.columns([1, 2])

with col1:
    st.subheader("🚫 Blocked (matching blacklist)")
    if df_blocked.empty:
        st.write("No blocked rows found.")
    else:
        st.dataframe(df_blocked.reset_index(drop=True))

with col2:
    st.subheader("✅ Clean list — Edit & Delete")
    edited_df = st.session_state.edited_df

    # Provide a small toolbar
    toolbar_cols = st.columns([1, 1, 1, 4])
    if toolbar_cols[0].button("Reset edits"):
        st.session_state.edited_df = df_clean.copy()
        st.experimental_rerun()
    if toolbar_cols[1].button("Remove all blocked" ):
        # already removed blocked; this is a convenience to ensure blocked rows not in edited
        st.session_state.edited_df = st.session_state.edited_df[~(normalize_series(st.session_state.edited_df[found_company_col]).isin(blacklist_companies) | normalize_series(st.session_state.edited_df[found_email_col]).isin(blacklist_emails))]
        st.experimental_rerun()
    if toolbar_cols[2].button("Show sample 10"):
        st.dataframe(st.session_state.edited_df.head(10))

    st.write("You can delete individual rows below. After edits, click Download to get the final XLSX.")

    # Display rows with delete buttons (index-based)
    if edited_df.empty:
        st.write("No rows to display.")
    else:
        # show a paginated/simple viewer to avoid too many buttons
        for idx in edited_df.index:
            row = edited_df.loc[idx]
            cols = st.columns([8, 1])
            # render the row as a single-line key summary plus expand
            with cols[0]:
                summary = f"{row.get(found_company_col, '')} — {row.get(found_email_col, '')}"
                with st.expander(summary, expanded=False):
                    st.write(row)
            if cols[1].button("🗑 Delete", key=f"del_{idx}"):
                edited_df = edited_df.drop(idx)
                st.session_state.edited_df = edited_df
                st.experimental_rerun()

    # Download final edited dataframe
    st.subheader("⬇️ Download Final Cleaned File")
    if not edited_df.empty:
        out = BytesIO()
        # write to excel
        try:
            edited_df.to_excel(out, index=False, engine='openpyxl')
            data = out.getvalue()
            st.download_button(
                label="Download Cleaned Excel (.xlsx)",
                data=data,
                file_name="final_cleaned_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Failed to create Excel file: {e}")
    else:
        st.write("No data to download.")

st.caption("Notes: The app compares Company Name and Email in lowercase & trimmed form. Blacklist entries can have empty values for either column.")
