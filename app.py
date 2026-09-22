import streamlit as st
import pandas as pd
import os

# 1. Page Configuration
st.set_page_config(page_title="Master CV Vault", page_icon="📄", layout="wide")
st.title("📄 Comprehensive Interactive CV Vault")
st.markdown("Live candidate database. Automatically synced with your master Excel file.")

def standardize_columns(df):
    """Cleans and consolidates messy column names to prevent duplicates."""
    df.columns = df.columns.astype(str).str.strip().str.replace(r'\s+', ' ', regex=True).str.title()
    
    rename_dict = {}
    for col in df.columns:
        if any(x in col for x in ['Education', 'Institution', 'Inatitution', 'Institition', 'University', 'Department']):
            rename_dict[col] = 'Education'
        elif any(x in col for x in ['Passing', 'Passig']):
            rename_dict[col] = 'Year Of Passing'
        elif any(x in col for x in ['Experience', 'Experiece', 'Experieece']):
            rename_dict[col] = 'Experience'
        elif 'Join' in col:
            rename_dict[col] = 'Date Of Joining'
        elif 'Salar' in col or 'Salay' in col:
            rename_dict[col] = 'Salary'
        elif any(x in col for x in ['Sl No', 'Sl.', 'Sl No.']):
            rename_dict[col] = 'Sl No.'
        elif 'Unnamed' in col:
            rename_dict[col] = 'DROP_ME'
            
    df = df.rename(columns=rename_dict)
    
    if 'DROP_ME' in df.columns:
        df = df.drop(columns=['DROP_ME'])
        
    df = df.loc[:, ~df.columns.duplicated()]
    return df

# 2. Hardcode your exact file name
file_path = "Copy of Candidate contact list_22 June 2025.xlsx"

# Check if the file exists in the folder
if os.path.exists(file_path):
    try:
        with st.spinner("Syncing latest data from database..."):
            # 3. Load, Clean, and Merge Data Automatically
            xls = pd.ExcelFile(file_path)
            all_sheets = []
            
            for sheet in xls.sheet_names:
                temp_df = pd.read_excel(file_path, sheet_name=sheet)
                temp_df = standardize_columns(temp_df)
                temp_df['Category (Sheet)'] = sheet
                all_sheets.append(temp_df)
            
            # Combine all sheets safely
            df = pd.concat(all_sheets, ignore_index=True)
        
            # 4. Final Data Polish
            df = df.dropna(subset=[col for col in df.columns if col != 'Category (Sheet)'], how='all')
            df = df.fillna("Not Specified")
            df = df.replace(r'^\s*$', 'Not Specified', regex=True)

        # 5. --- SIDEBAR FILTERS ---
        st.sidebar.header("🔍 Global Filters")
        
        active_filter_cols = []
        
        # Search by Name
        name_cols = [col for col in df.columns if 'name' in col.lower() and 'company' not in col.lower()]
        if name_cols:
            search_name = st.sidebar.text_input("Search Candidate by Name")
            if search_name:
                df = df[df[name_cols[0]].astype(str).str.contains(search_name, case=False, na=False)]

        # Filter by Category
        if 'Category (Sheet)' in df.columns:
            categories = [c for c in df['Category (Sheet)'].unique() if c != "Not Specified"]
            selected_cat = st.sidebar.multiselect("Filter by Category", sorted(categories))
            if selected_cat:
                df = df[df['Category (Sheet)'].isin(selected_cat)]
                active_filter_cols.append('Category (Sheet)')

        # Filter by Company
        company_cols = [col for col in df.columns if 'company' in col.lower()]
        if company_cols:
            comp_col = company_cols[0]
            companies = [c for c in df[comp_col].astype(str).unique() if c != "Not Specified" and c.strip() != ""]
            selected_company = st.sidebar.multiselect("Select Company", sorted(companies))
            if selected_company:
                df = df[df[comp_col].astype(str).isin(selected_company)]
                active_filter_cols.append(comp_col)

        # Filter by Designation/Role
        designation_cols = [col for col in df.columns if 'designation' in col.lower()]
        if designation_cols:
            desig_col = designation_cols[0]
            roles = [r for r in df[desig_col].astype(str).unique() if r != "Not Specified" and r.strip() != ""]
            selected_roles = st.sidebar.multiselect("Select Role / Designation", sorted(roles))
            if selected_roles:
                df = df[df[desig_col].astype(str).isin(selected_roles)]
                active_filter_cols.append(desig_col)
                
        # Search by Education
        if 'Education' in df.columns:
            search_edu = st.sidebar.text_input("Search Education (e.g., BBA, DU, MBA)")
            if search_edu:
                df = df[df['Education'].astype(str).str.contains(search_edu, case=False, na=False)]
                active_filter_cols.append('Education')

        # 6. --- DYNAMIC COLUMN REORDERING ---
        front_cols = []
        if 'Sl No.' in df.columns:
            front_cols.append('Sl No.')
        if name_cols and name_cols[0] in df.columns:
            front_cols.append(name_cols[0])
            
        remaining_cols = [col for col in df.columns if col not in front_cols and col not in active_filter_cols]
        df = df[front_cols + active_filter_cols + remaining_cols]

        # 7. --- MAIN DISPLAY ---
        st.subheader(f"Total Candidates Found: {len(df)}")
        st.dataframe(df, use_container_width=True)

        # 8. Export Feature
        csv = df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Download Current View as CSV",
            data=csv,
            file_name='filtered_master_candidates.csv',
            mime='text/csv',
        )

    except Exception as e:
        st.error(f"Error processing the database file: {e}")
else:
    st.error(f"Database file '{file_path}' not found! Please ensure it is in the same folder as app.py or uploaded to GitHub.")
