import streamlit as st
import pandas as pd

# 1. Page Configuration
st.set_page_config(page_title="Master CV Vault", page_icon="📄", layout="wide")
st.title("📄 Comprehensive Interactive CV Vault")
st.markdown("Upload your master candidate Excel file. This app merges all sheets, cleans messy data, and consolidates duplicate columns.")

def standardize_columns(df):
    """Cleans and consolidates messy column names to prevent duplicates."""
    # Standardize spaces and clean layout strings
    df.columns = df.columns.astype(str).str.strip().str.replace(r'\s+', ' ', regex=True).str.title()
    
    rename_dict = {}
    for col in df.columns:
        # Catch variations of "Education" / "Institution" / "University" / "Department"
        if any(x in col for x in ['Education', 'Institution', 'Inatitution', 'Institition', 'University', 'Department']):
            rename_dict[col] = 'Education'
        # Catch variations of "Year of Passing"
        elif any(x in col for x in ['Passing', 'Passig']):
            rename_dict[col] = 'Year Of Passing'
        # Catch variations of "Experience"
        elif any(x in col for x in ['Experience', 'Experiece', 'Experieece']):
            rename_dict[col] = 'Experience'
        # Catch variations of "Joining"
        elif 'Join' in col:
            rename_dict[col] = 'Date Of Joining'
        # Catch variations of "Salary"
        elif 'Salar' in col or 'Salay' in col:
            rename_dict[col] = 'Salary'
        # Catch serial numbers
        elif any(x in col for x in ['Sl No', 'Sl.', 'Sl No.']):
            rename_dict[col] = 'Sl No.'
        # Identify junk columns created by Excel by accident
        elif 'Unnamed' in col:
            rename_dict[col] = 'DROP_ME'
            
    # Apply the clean standardized names
    df = df.rename(columns=rename_dict)
    
    # Remove hidden junk columns
    if 'DROP_ME' in df.columns:
        df = df.drop(columns=['DROP_ME'])
        
    # Drop any duplicate columns created WITHIN this single sheet after renaming 
    df = df.loc[:, ~df.columns.duplicated()]
        
    return df

# 2. File Uploader
uploaded_file = st.file_uploader("Upload your master Excel file (.xlsx) or CSV", type=['xlsx', 'csv'])

if uploaded_file is not None:
    try:
        with st.spinner("Compiling, cleaning, and standardizing data from all sheets..."):
            # 3. Load, Clean, and Merge Data
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
                df = standardize_columns(df)
                df['Category (Sheet)'] = 'CSV Upload'
            else:
                xls = pd.ExcelFile(uploaded_file)
                all_sheets = []
                
                for sheet in xls.sheet_names:
                    temp_df = pd.read_excel(uploaded_file, sheet_name=sheet)
                    temp_df = standardize_columns(temp_df)
                    temp_df['Category (Sheet)'] = sheet
                    all_sheets.append(temp_df)
                
                # Combine all sheets safely
                df = pd.concat(all_sheets, ignore_index=True)
        
        # 4. Final Data Polish
        # Remove lines that are completely empty or contain only commas from Excel
        df = df.dropna(subset=[col for col in df.columns if col != 'Category (Sheet)'], how='all')
        
        # Force fill all empty fields with "Not Specified"
        df = df.fillna("Not Specified")
        df = df.replace(r'^\s*$', 'Not Specified', regex=True)

        st.success("All data merged and cleaned successfully!")

        # 5. --- SIDEBAR FILTERS ---
        st.sidebar.header("🔍 Global Filters")
        st.sidebar.markdown("Filter across your entire talent pool.")
        
        # Track which filters the user is actively using for dynamic column sorting
        active_filter_cols = []
        
        # A. Global Search by Name
        name_cols = [col for col in df.columns if 'name' in col.lower() and 'company' not in col.lower()]
        if name_cols:
            search_name = st.sidebar.text_input("Search Candidate by Name")
            if search_name:
                df = df[df[name_cols[0]].astype(str).str.contains(search_name, case=False, na=False)]

        # B. Filter by Category (The original Excel Sheet)
        if 'Category (Sheet)' in df.columns:
            categories = [c for c in df['Category (Sheet)'].unique() if c != "Not Specified"]
            selected_cat = st.sidebar.multiselect("Filter by Category", sorted(categories))
            if selected_cat:
                df = df[df['Category (Sheet)'].isin(selected_cat)]
                active_filter_cols.append('Category (Sheet)')

        # C. Filter by Company
        company_cols = [col for col in df.columns if 'company' in col.lower()]
        if company_cols:
            comp_col = company_cols[0]
            companies = [c for c in df[comp_col].astype(str).unique() if c != "Not Specified" and c.strip() != ""]
            selected_company = st.sidebar.multiselect("Select Company", sorted(companies))
            if selected_company:
                df = df[df[comp_col].astype(str).isin(selected_company)]
                active_filter_cols.append(comp_col)

        # D. Filter by Designation/Role
        designation_cols = [col for col in df.columns if 'designation' in col.lower()]
        if designation_cols:
            desig_col = designation_cols[0]
            roles = [r for r in df[desig_col].astype(str).unique() if r != "Not Specified" and r.strip() != ""]
            selected_roles = st.sidebar.multiselect("Select Role / Designation", sorted(roles))
            if selected_roles:
                df = df[df[desig_col].astype(str).isin(selected_roles)]
                active_filter_cols.append(desig_col)
                
        # E. Global Search by Education
        if 'Education' in df.columns:
            search_edu = st.sidebar.text_input("Search Education (e.g., BBA, DU, MBA, PGD)")
            if search_edu:
                df = df[df['Education'].astype(str).str.contains(search_edu, case=False, na=False)]
                active_filter_cols.append('Education')

        # 6. --- DYNAMIC COLUMN REORDERING ---
        # Pin Sl No. and Name to the front
        front_cols = []
        if 'Sl No.' in df.columns:
            front_cols.append('Sl No.')
        if name_cols and name_cols[0] in df.columns:
            front_cols.append(name_cols[0])
            
        # Get all the columns that aren't pinned to the front and aren't active filters
        remaining_cols = [col for col in df.columns if col not in front_cols and col not in active_filter_cols]
        
        # Reshuffle the dataframe: [Pinned] + [Active Filters] + [Everything Else]
        df = df[front_cols + active_filter_cols + remaining_cols]


        # 7. --- MAIN DISPLAY ---
        st.subheader(f"Total Candidates Found: {len(df)}")
        
        # Display the table
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
        st.error(f"Error processing file: {e}")
else:
    st.info("Waiting for file upload...")