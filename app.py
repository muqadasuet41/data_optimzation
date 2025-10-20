# app.py
# Streamlit Skill Profiling Merger — works in Google Colab local tunnel and Streamlit Cloud
# Usage: Upload multiple employee Excel files (from Cycle1 and Cycle2). Export final_master.xlsx.

import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

st.set_page_config(page_title="Skill Profiling Merger", layout="wide")

st.title("Skill Profiling — Auto Master File Generator")
st.markdown(
    """
Upload all employee Excel files (Cycle 1 and Cycle 2).
This app will:
- Parse each employee file (supports row-wise or column-wise skill formats)
- Detect cycle where possible (filename contains 'cycle1'/'cycle2' or dates)
- Merge into Cycle1 / Cycle2 sheets and a Combined master sheet
- Provide final_master.xlsx download
"""
)

def detect_cycle_from_filename(fn: str):
    fn_low = fn.lower()
    if "cycle1" in fn_low or "c1" in fn_low:
        return 1
    if "cycle2" in fn_low or "c2" in fn_low:
        return 2
    # try detect year-month or date in filename
    date_match = re.search(r'(20\d{2}[-_]\d{1,2}[-_]\d{1,2})', fn)
    if date_match:
        try:
            dt = pd.to_datetime(date_match.group(1).replace('_','-'))
            # heuristic: older -> cycle 1, newer -> cycle 2
            # We can't be fully certain — return None to let merging logic handle.
            return None
        except:
            return None
    return None

def parse_excel_bytes(bytes_io, filename):
    """
    Return DataFrame with columns: ['Employee','Skill','Level','Cycle']
    Accepts many layouts:
    - Long format: columns include 'Skill' and 'Level' (or similar)
    - Wide format: first column maybe 'Skill' or row contains skills as columns (employee name as filename)
    """
    try:
        # try reading first sheet
        xls = pd.read_excel(bytes_io, sheet_name=0, header=None)
    except Exception as e:
        # fallback: read with default pandas detection
        bytes_io.seek(0)
        df = pd.read_excel(bytes_io, sheet_name=0)
        xls = df

    # We'll try multiple heuristics
    # 1) If it has header row with 'skill' and 'level'
    bytes_io.seek(0)
    df = pd.read_excel(bytes_io, sheet_name=0)
    raw = df.copy()
    df_cols = [str(c).strip().lower() for c in raw.columns]

    employee_name = re.sub(r'\.xlsx?$','', filename, flags=re.I)

    # Heuristic A: explicit 'skill' and 'level' columns
    skill_col = None
    level_col = None
    for c in raw.columns:
        cname = str(c).strip().lower()
        if cname in ['skill', 'skill name', 'skills', 'skill_title', 'competency', 'competency name']:
            skill_col = c
        if cname in ['level','skill level', 'rating','proficiency','score']:
            level_col = c

    rows = []
    detected_cycle = detect_cycle_from_filename(filename)
    if skill_col is not None and level_col is not None:
        for _, r in raw.iterrows():
            sk = r[skill_col]
            lv = r[level_col]
            if pd.isna(sk):
                continue
            try:
                lv_val = int(lv)
            except:
                # attempt numeric casting
                try:
                    lv_val = int(float(lv))
                except:
                    lv_val = lv
            rows.append({'Employee': employee_name, 'Skill': str(sk).strip(), 'Level': lv_val, 'Cycle': detected_cycle})
        return pd.DataFrame(rows)

    # Heuristic B: long format but header names different (first column skill, second level)
    if raw.shape[1] >= 2:
        # check whether first column looks like skill names and second numeric
        first_col = raw.columns[0]
        second_col = raw.columns[1]
        # Count numeric-like in second column
        numeric_count = pd.to_numeric(raw[second_col], errors='coerce').notna().sum()
        if numeric_count >= max(1, int(0.3 * len(raw))):
            for _, r in raw.iterrows():
                sk = r[first_col]
                lv = r[second_col]
                if pd.isna(sk):
                    continue
                try:
                    lv_val = int(lv)
                except:
                    try:
                        lv_val = int(float(lv))
                    except:
                        lv_val = lv
                rows.append({'Employee': employee_name, 'Skill': str(sk).strip(), 'Level': lv_val, 'Cycle': detected_cycle})
            return pd.DataFrame(rows)

    # Heuristic C: wide format where columns are skills and single row contains levels
    # If columns are skill names and there are numeric entries under them, treat as one employee row (or multiple rows)
    # We will iterate rows and find numeric-like values in many columns
    candidates = []
    for idx, r in raw.iterrows():
        numeric_cells = 0
        skill_values = {}
        for col in raw.columns:
            val = r[col]
            if pd.isna(val):
                continue
            # treat as numeric skill level when small integer 0-5
            try:
                n = int(val)
                if 0 <= n <= 10:
                    numeric_cells += 1
                    skill_values[str(col).strip()] = n
            except:
                # treat strings like '3' or '2'?
                try:
                    n = int(float(val))
                    numeric_cells += 1
                    skill_values[str(col).strip()] = n
                except:
                    pass
        if numeric_cells >= 1:
            # Make rows out of skill_values
            for sk, lv in skill_values.items():
                rows.append({'Employee': employee_name, 'Skill': sk, 'Level': lv, 'Cycle': detected_cycle})
    if rows:
        return pd.DataFrame(rows)

    # Last fallback: If nothing detected well, try flattening all string cells as possible skill names with level NaN
    for _, r in raw.iterrows():
        for col in raw.columns:
            val = r[col]
            if pd.isna(val):
                continue
            s = str(val).strip()
            if len(s) > 0 and len(s) < 80:
                # keep as skill with unknown level
                rows.append({'Employee': employee_name, 'Skill': s, 'Level': None, 'Cycle': detected_cycle})
    if rows:
        return pd.DataFrame(rows)

    # if still nothing, return empty df
    return pd.DataFrame(columns=['Employee','Skill','Level','Cycle'])


def merge_cycles(dfs):
    """
    dfs: list of dataframes each with Employee, Skill, Level, Cycle
    Logic:
    - Partition into cycle1 and cycle2 if cycle info available.
    - For combined master: for same Employee+Skill use the value from Cycle2 if present; otherwise Cycle1; if multiple entries use max level as a fallback.
    - Also produce Cycle1 sheet and Cycle2 sheet.
    """
    combined = pd.concat(dfs, ignore_index=True, sort=False)
    # normalize skill strings
    combined['Skill'] = combined['Skill'].astype(str).str.strip()
    combined['Employee'] = combined['Employee'].astype(str).str.strip()

    # Partition by Cycle info when available
    df_c1 = combined[combined['Cycle']==1].copy()
    df_c2 = combined[combined['Cycle']==2].copy()
    # entries with no cycle info
    df_unknown = combined[combined['Cycle'].isna()].copy()

    # If no explicit cycle tags, try to infer by upload order cannot be done reliably here.
    # We'll treat unknown as part of cycle2 (latest) if both cycles present overall.
    if df_c1.empty and df_unknown.empty and not df_c2.empty:
        df_c1 = pd.DataFrame(columns=combined.columns)
    if df_c1.empty and df_c2.empty and not df_unknown.empty:
        # all files unknown: treat all as cycle2 (latest)
        df_c2 = df_unknown.copy()
        df_unknown = pd.DataFrame(columns=combined.columns)

    # if both c1 and c2 exist, keep unknown as c2 (assuming they are updates)
    if not df_c1.empty and not df_c2.empty and not df_unknown.empty:
        df_c2 = pd.concat([df_c2, df_unknown], ignore_index=True)
        df_unknown = pd.DataFrame(columns=combined.columns)

    # If only unknown exists, assign to cycle2
    if df_c1.empty and df_c2.empty and not df_unknown.empty:
        df_c2 = df_unknown.copy()
        df_unknown = pd.DataFrame(columns=combined.columns)

    # Create helper pivot per cycle: for each employee+skill, pick max level (if multiple rows)
    def normalize_cycle_df(dfc):
        if dfc.empty:
            return dfc
        # coerce numeric
        dfc['Level_num'] = pd.to_numeric(dfc['Level'], errors='coerce')
        dfc = dfc.sort_values(by=['Employee','Skill','Level_num'], ascending=[True,True,False])
        # keep first (highest) level per employee+skill
        dfc = dfc.groupby(['Employee','Skill'], as_index=False).first()[['Employee','Skill','Level_num']]
        dfc.rename(columns={'Level_num':'Level'}, inplace=True)
        return dfc

    n1 = normalize_cycle_df(df_c1)
    n2 = normalize_cycle_df(df_c2)

    # Combined: start from n1, then overlay n2 values, and also include skills only in n2
    combined_master = pd.concat([n1, n2], ignore_index=True)
    # keep the highest-level per Employee+Skill; if duplicates, prefer n2 value by marking origin
    # Use groupby and take max Level
    combined_master['Level'] = pd.to_numeric(combined_master['Level'], errors='coerce')
    combined_master = combined_master.groupby(['Employee','Skill'], as_index=False)['Level'].max()
    combined_master = combined_master.sort_values(['Employee','Skill']).reset_index(drop=True)

    return n1.sort_values(['Employee','Skill']), n2.sort_values(['Employee','Skill']), combined_master

# === Streamlit upload UI ===
st.sidebar.header("Upload & Options")
uploaded_files = st.sidebar.file_uploader("Upload employee Excel files (.xlsx .xls) — multiple", accept_multiple_files=True, type=['xlsx','xls'])

auto_filename = st.sidebar.text_input("Master output filename", value="final_master.xlsx")

show_preview = st.sidebar.checkbox("Show parsed preview", value=True)

if uploaded_files:
    st.sidebar.markdown(f"Uploaded files: {len(uploaded_files)}")
    parsed_dfs = []
    parse_errors = []
    for uf in uploaded_files:
        try:
            bytes_data = uf.read()
            parsed = parse_excel_bytes(io.BytesIO(bytes_data), uf.name)
            if parsed.empty:
                parse_errors.append(uf.name)
            else:
                parsed_dfs.append(parsed)
                if show_preview:
                    with st.expander(f"Preview parsed from {uf.name}"):
                        st.dataframe(parsed.head(200))
        except Exception as e:
            parse_errors.append(f"{uf.name} : {str(e)}")

    if parse_errors:
        st.sidebar.warning("Files that couldn't be parsed cleanly (open them to check format):\n" + "\n".join(parse_errors))

    if parsed_dfs:
        st.success("Parsing complete. Now merging...")
        c1_df, c2_df, master_df = merge_cycles(parsed_dfs)

        st.header("Master Summary")
        st.subheader("Combined Master (Employee - Skill - Level)")
        st.dataframe(master_df)

        # Provide per-employee pivot option
        st.subheader("Pivot: Employees as rows, Skills as columns")
        pivot = master_df.pivot_table(index='Employee', columns='Skill', values='Level', aggfunc='first').fillna('')
        st.dataframe(pivot)

        # Prepare excel to download
        towrite = io.BytesIO()
        with pd.ExcelWriter(towrite, engine='openpyxl') as writer:
            c1_df.to_excel(writer, sheet_name='Cycle1', index=False)
            c2_df.to_excel(writer, sheet_name='Cycle2', index=False)
            master_df.to_excel(writer, sheet_name='Master_Combined', index=False)
            pivot.to_excel(writer, sheet_name='Master_Pivot', index=True)
        towrite.seek(0)

        st.download_button(label="Download final master .xlsx", data=towrite, file_name=auto_filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # Also allow saving to server (useful in Colab)
        if st.button("Save final_master.xlsx on server (app working dir)"):
            with open(auto_filename, "wb") as f:
                f.write(towrite.read())
            st.success(f"Saved as {auto_filename} in current working directory.")

else:
    st.info("Upload employee Excel files (multiple). If testing in Google Colab, use google.colab.files.upload to upload files to the notebook workspace and then use the Tunnel commands to run the Streamlit app.")
