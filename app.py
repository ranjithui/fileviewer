"""
Streamlit Data Cleaner – Dual Upload Version (Cloud-friendly)
------------------------------------------------------------
- Upload your main Excel file
- Upload your blacklist Excel file
- Removes rows where Company Name OR Email matches blacklist
- Live edit / delete rows
- Download cleaned Excel
"""

import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Data Cleaner + Live Editor", layout="wide")

# ---------------------- Uploaders ----------------------
st.title("📋 Data Cleaner + Live Editor (Dual Upload)")
st.markdown(
    "Upload your main Excel file and your blacklist Excel file. "
    "The app will remove rows where Company Name or Email exists in the blacklist. "
    "You can edit/delete rows live and download the final cleaned file."
)

uploaded_file = st.file_uploader("Upload your main Excel file (XLSX)", type=["xlsx"])
uploaded_blacklist = st.file_uploader("Upload your Blacklist Excel file (XLSX)", type=["xlsx"])

if not uploaded_file or not uploaded_blacklist:
    st.info("Please upload both the main file and the blacklist to start processing.")
    st.stop()  # Prevent processing until both files are uploaded

# ---------------------- Load main file ----------------------
try:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
except Exception as e:
    st.error(f"Unable to read main Excel file: {e}")
    st.stop()

# ---------------------- Load blacklist ----------------------
try:
    blacklist = pd.read_excel(uploaded_blacklist, engine="openpyxl")
except Exception as e:
    st.error(f"Unable to read blacklist Excel file: {e}")
    st.stop()

# ---------------------- Keys ----------------------
KEY_COMPANY = "Company Name"
KEY_EMAIL = "Email"

def normalize_series(s: pd.Series) -> pd.Series:
    return s.fillna("").astype(str).str.lower().str.strip()

def get_col(df, key):
    for c in df.columns:
        if c.lower() == key.lower():
            return c
    return None

# ---------------------- Check columns ----------------------
for df_check, name in [(df, "Main file"), (blacklist, "Blacklist")]:
    if not get_col(df_check, KEY_COMPANY):
        st.error(f"{name} missing column: {KEY_COMPANY}")
        st.stop()
    if not get_col(df_check, KEY_EMAIL):
        st.error(f"{name} missing column: {KEY_EMAIL}")
        st.stop()

main_company_col = get_col(df, KEY_COMPANY)
main_email_col = get_col(df, KEY_EMAIL)
bl_company_col = get_col(blacklist, KEY_COMPANY)
bl_email_col = get_col(blacklist, KEY_EMAIL)

# ---------------------- Prepare blacklist sets ----------------------
blacklist_companies = set(normalize_series(blacklist[bl_company_col]))
blacklist_emails = set(normalize_series(blacklist[bl_email_col]))

# ---------------------- Mark blocked rows ----------------------
df_key_company = normalize_series(df[main_company_col])
df_key_email = normalize_series(df[main_email_col])

is_company_blacklisted = df_key_company.isin(blacklist_companies)
is_email_blacklisted = df_key_email.isin(blacklist_emails)
blacklisted_mask = is_company_blacklisted | is_email_blacklisted

df_blocked = df[blacklisted_mask].copy()
df_clean = df[~blacklisted_mask].copy()

# Remove duplicates
df_clean = df_clean.drop_duplicates(subset=[main_company_col, main_email_col])

# Sort by company
df_clean['_sort_company'] = normalize_series(df_clean[main_company_col])
df_clean = df_clean.sort_values('_sort_company').drop(columns=['_sort_company'])

# Initialize session state
if 'edited_df' not in st.session_state:
    st.session_state.edited_df = df_clean.copy()

st.success(f"Loaded main file: {len(df)} rows. Blocked: {len(df_blocked)}. Clean: {len(df_clean)}.")

# ---------------------- Layout ----------------------
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

    # Toolbar
    toolbar_cols = st.columns([1, 1, 1, 4])
    if toolbar_cols[0].button("Reset edits"):
        st.session_state.edited_df = df_clean.copy()
        st.experimental_rerun()
    if toolbar_cols[1].button("Remove all blocked"):
        edited_df = edited_df[~(normalize_series(edited_df[main_company_col]).isin(blacklist_companies) |
                                normalize_series(edited_df[main_email_col]).isin(blacklist_emails))]
        st.session_state.edited_df = edited_df
        st.experimental_rerun()
    if toolbar_cols[2].button("Show sample 10"):
        st.dataframe(edited_df.head(10))

    st.write("You can delete individual rows below. After edits, click Download to get the final XLSX.")

    if edited_df.empty:
        st.write("No rows to display.")
    else:
        for idx in edited_df.index:
            row = edited_df.loc[idx]
            cols = st.columns([8, 1])
            with cols[0]:
                summary = f"{row.get(main_company_col, '')} — {row.get(main_email_col, '')}"
                with st.expander(summary, expanded=False):
                    st.write(row)
            if cols[1].button("🗑 Delete", key=f"del_{idx}"):
                edited_df = edited_df.drop(idx)
                st.session_state.edited_df = edited_df
                st.experimental_rerun()

    # Download cleaned Excel
    st.subheader("⬇️ Download Final Cleaned File")
    if not edited_df.empty:
        out = BytesIO()
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
