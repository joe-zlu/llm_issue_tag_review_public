import streamlit as st
import pandas as pd
import sqlite3
import os
import time
from datetime import datetime
from io import BytesIO

# --- Configuration ---
st.set_page_config(page_title="Issue Tag Reviewer", layout="wide")
TEMP_FOLDER = "temp"

# Ensure temp folder exists
if not os.path.exists(TEMP_FOLDER):
    os.makedirs(TEMP_FOLDER)

# --- Constants ---
PREDEFINED_TAGS = [
    "SAS/AP", "Regulatory Principles", "ARTG (Australian Register of Therapeutic Goods)",
    "Transitional Pathway", "Scheduling", "Evidence", "Enforcement", "Labelling",
    "Specialist Oversight", "Clinical Guidance", "Quality Standards", "Fee Waiver",
    "PV (Pharmacovigilance)", "Child-resistant Packaging", "Advertising", "Education",
    "GMP (Good Manufacturing Practice)", "THC limit", "No THC Access", "Efficacy",
    "Restrict dosage forms", "Inhalation or Vapourisation", "Cannabinoids", "Categories",
    "HREC (Human Research Ethics Committees)", "Access", "Active Ingredient"
]

REQUIRED_COLUMNS = [
    "source", "Issue", "StakeholderTypeArray", "WorksheetLabelArray",
    "IssueTag1", "IssueTag2", "IssueTag3", "IssueTag4",
    "IssueTag5", "IssueTag6", "IssueTag7", "IssueTag8"
]

# --- Database Helper Functions ---

def get_db_path(filename):
    return os.path.join(TEMP_FOLDER, filename)

def ensure_schema_compatibility(db_path):
    """Ensures older databases have the review_date column."""
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT review_date FROM issues LIMIT 1")
    except sqlite3.OperationalError:
        cursor.execute("ALTER TABLE issues ADD COLUMN review_date TEXT")
        conn.commit()
    finally:
        conn.close()

def init_db(df, filename):
    """Saves dataframe to sqlite with specific schema adjustments."""
    db_path = get_db_path(filename)
    
    # Add ID column for record tracking
    df['record_id'] = range(1, len(df) + 1)
    
    # Add new columns for review
    df['reviewed_tags'] = None 
    df['tagging_notes'] = ""
    df['review_date'] = None
    
    # Ensure all required columns exist
    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            df[col] = None

    conn = sqlite3.connect(db_path)
    df.to_sql('issues', conn, if_exists='replace', index=False)
    conn.close()

def load_data_from_db(filename, filter_source=None, filter_label=None):
    """Loads data with optional filtering."""
    db_path = get_db_path(filename)
    if not os.path.exists(db_path):
        return pd.DataFrame()
    
    ensure_schema_compatibility(db_path)
    
    conn = sqlite3.connect(db_path)
    
    query = "SELECT * FROM issues WHERE 1=1"
    params = []
    
    if filter_source and filter_source != "All":
        query += " AND source = ?"
        params.append(filter_source)
        
    if filter_label and filter_label != "All":
        query += " AND WorksheetLabelArray = ?"
        params.append(filter_label)
        
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df

def update_record(filename, record_id, reviewed_tags, tagging_notes):
    """Updates a specific record in the DB and sets review_date."""
    db_path = get_db_path(filename)
    conn = sqlite3.connect(db_path, timeout=10)
    cursor = conn.cursor()
    
    review_timestamp = datetime.now().isoformat()
    
    cursor.execute("""
        UPDATE issues 
        SET reviewed_tags = ?, tagging_notes = ?, review_date = ?
        WHERE record_id = ?
    """, (reviewed_tags, tagging_notes, review_timestamp, int(record_id)))
    
    conn.commit()
    conn.close()

def get_distinct_values(filename, column_name):
    """Helper to get distinct values for filters."""
    db_path = get_db_path(filename)
    conn = sqlite3.connect(db_path)
    try:
        df = pd.read_sql_query(f"SELECT DISTINCT \"{column_name}\" FROM issues", conn)
        return sorted(df[column_name].dropna().astype(str).unique().tolist())
    except:
        return []
    finally:
        conn.close()

def to_excel(df):
    """Converts dataframe to excel bytes."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Reviewed Data')
    processed_data = output.getvalue()
    return processed_data

# --- App Logic ---

tab1, tab2, tab3 = st.tabs(["📂 Data Management", "📝 Tag Review", "📤 Export Data"])

# ==========================================
# TAB 1: Data Management
# ==========================================
with tab1:
    st.header("Import and Manage Datasets")
    
    # --- Section 1: New Upload ---
    st.subheader("1. Upload New Excel File")
    st.info("Required Excel columns: source, Issue, StakeholderTypeArray, WorksheetLabelArray, \
        IssueTag1, IssueTag2, IssueTag3, IssueTag4, \
        IssueTag5, IssueTag6, IssueTag7, IssueTag8")
    uploaded_file = st.file_uploader("Choose an XLSX file", type="xlsx")
    
    if uploaded_file:
        try:
            xl = pd.ExcelFile(uploaded_file)
            sheet_names = xl.sheet_names
            selected_sheet = st.selectbox("Select Worksheet", sheet_names)
            
            if selected_sheet:
                df_preview = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
                st.write("Preview (First 5 rows):")
                st.dataframe(df_preview.head())
                
                missing_cols = [c for c in REQUIRED_COLUMNS if c not in df_preview.columns]
                
                if missing_cols:
                    st.error(f"Missing columns: {', '.join(missing_cols)}")
                else:
                    if st.button("Import to Database"):
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        clean_name = uploaded_file.name.replace(".xlsx", "")
                        db_filename = f"{clean_name}_{timestamp}.db"
                        
                        with st.spinner("Creating database..."):
                            init_db(df_preview, db_filename)
                        
                        st.success(f"Database created: {db_filename}")
                        st.session_state['refresh_trigger'] = time.time() 
                        st.rerun()
        except Exception as e:
            st.error(f"Error reading file: {e}")

    st.markdown("---")

    # --- Section 2: Manage Database ---
    st.subheader("2. Load/Manage Database")
    
    # Get List of files
    db_files = [f for f in os.listdir(TEMP_FOLDER) if f.endswith('.db')]
    db_files.sort(reverse=True) 
    
    # Display Active Status
    active_db_name = st.session_state.get('active_db', None)
    if active_db_name and os.path.exists(get_db_path(active_db_name)):
        st.success(f"✅ **Active Database:** {active_db_name}")
    else:
        st.warning("⚠️ No database loaded. Please select and load a database.")
        active_db_name = None

    if not db_files:
        st.info("No databases found on server.")
    else:
        # Selection Dropdown
        selected_db_to_manage = st.selectbox("Select Database File:", db_files)
        
        # Action Buttons
        if selected_db_to_manage:
            
            # --- ACTION: LOAD ---
            if st.button("📂 Load Database"):
                st.session_state['active_db'] = selected_db_to_manage
                st.rerun()

            # --- ACTION: RENAME ---
            with st.expander("Rename Database"):
                new_name_input = st.text_input("New Name (without extension)", value=selected_db_to_manage.replace(".db", ""))
                if st.button("Rename"):
                    if new_name_input:
                        clean_new_name = new_name_input.strip()
                        if not clean_new_name.endswith(".db"):
                            clean_new_name += ".db"
                        
                        old_path = get_db_path(selected_db_to_manage)
                        new_path = get_db_path(clean_new_name)
                        
                        if os.path.exists(new_path):
                            st.error("A file with that name already exists.")
                        else:
                            try:
                                os.rename(old_path, new_path)
                                # If we renamed the active DB, update the session state
                                if active_db_name == selected_db_to_manage:
                                    st.session_state['active_db'] = clean_new_name
                                st.success(f"Renamed to {clean_new_name}")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Error renaming: {e}")
                    else:
                        st.error("Please enter a valid name.")

            # --- ACTION: DELETE ---
            with st.expander("Delete Database"):
                st.write(f"Are you sure you want to delete **{selected_db_to_manage}**?")
                if st.button("Confirm Delete", type="primary"):
                    try:
                        os.remove(get_db_path(selected_db_to_manage))
                        # If we deleted the active DB, clear session state
                        if active_db_name == selected_db_to_manage:
                            st.session_state['active_db'] = None
                        st.success("Database deleted.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error deleting: {e}")

# ==========================================
# TAB 2: Tag Review
# ==========================================
with tab2:
    if 'active_db' not in st.session_state or not st.session_state['active_db'] or not os.path.exists(get_db_path(st.session_state['active_db'])):
        st.warning("Please load a database in Data Management tab.")
    else:
        current_db = st.session_state['active_db']
        
        # --- Local Filters for Tab 2 ---
        st.subheader("Review Filters")
        
        # Helper to reset nav when filters change
        def reset_nav():
            st.session_state['nav_index'] = 0

        source_opts = ["All"] + get_distinct_values(current_db, "source")
        label_opts = ["All"] + get_distinct_values(current_db, "WorksheetLabelArray")
        
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            sel_source_t2 = st.selectbox("Filter by Source", source_opts, key="t2_source", on_change=reset_nav)
        with col_f2:
            sel_label_t2 = st.selectbox("Filter by Worksheet Label", label_opts, key="t2_label", on_change=reset_nav)
            
        st.markdown("---")
        
        # Load filtered data
        df_filtered = load_data_from_db(current_db, sel_source_t2, sel_label_t2)
        
        st.write(f"**Records Found:** {len(df_filtered)}")
        
        if len(df_filtered) == 0:
            st.info("No records match the filters.")
        else:
            # --- Navigation State Management ---
            if 'nav_index' not in st.session_state:
                st.session_state['nav_index'] = 0
            
            # Ensure index is within bounds
            if st.session_state['nav_index'] >= len(df_filtered):
                st.session_state['nav_index'] = 0
                
            current_idx = st.session_state['nav_index']
            
            # --- Navigation Functions ---
            def next_record():
                if st.session_state['nav_index'] < len(df_filtered) - 1:
                    st.session_state['nav_index'] += 1

            def prev_record():
                if st.session_state['nav_index'] > 0:
                    st.session_state['nav_index'] -= 1

            def next_unreviewed():
                found = False
                for i in range(st.session_state['nav_index'] + 1, len(df_filtered)):
                    val = df_filtered.iloc[i]['review_date']
                    if pd.isna(val):
                        st.session_state['nav_index'] = i
                        found = True
                        break
                
                if not found:
                    st.info("No more unreviewed records found below this point.")

            # --- Record Display ---
            row = df_filtered.iloc[current_idx]
            record_id = row['record_id']
            
            # Review Status
            is_reviewed = pd.notna(row['review_date'])
            if is_reviewed:
                st.caption(f"✅ Reviewed on: {row['review_date']}")
            else:
                st.caption("❌ Not yet reviewed")
            
            # Display Issue
            st.subheader("Issue")
            st.info(row['Issue'])
            
            # --- Tag Logic ---
            col_review, col_saved = st.columns(2)
            
            # Prepare "LLM Tags" list
            raw_tags = []
            for i in range(1, 9):
                t = row.get(f"IssueTag{i}")
                if pd.notna(t) and str(t).strip() != "":
                    raw_tags.append(str(t).strip())
            
            all_options = sorted(list(set(raw_tags + PREDEFINED_TAGS)))

            with col_review:
                st.markdown("### 1. LLM Tags for Review")
                st.caption("Review proposed tags, add new ones, then click Confirm.")
                
                llm_tags_selected = st.multiselect(
                    "Select/Edit Tags:",
                    options=all_options,
                    default=[t for t in raw_tags if t in all_options], 
                    key=f"llm_select_{record_id}"
                )
                
                if st.button("Confirm Tags >>", key=f"btn_confirm_{record_id}"):
                    csv_tags = ",".join(llm_tags_selected)
                    current_notes = row['tagging_notes'] if pd.notna(row['tagging_notes']) else ""
                    update_record(current_db, record_id, csv_tags, current_notes)
                    # Update the session state for the reviewed tags widget so it reflects the change immediately
                    st.session_state[f"db_view_{record_id}"] = llm_tags_selected
                    st.success("Tags confirmed and saved.")
                    st.rerun()

            with col_saved:
                st.markdown("### 2. Database: Reviewed Tags")
                st.caption("Current status in database.")
                
                current_reviewed_str = row['reviewed_tags']
                current_reviewed_list = []
                
                if pd.notna(current_reviewed_str) and current_reviewed_str != "":
                    current_reviewed_list = current_reviewed_str.split(",")
                
                edited_reviewed_tags = st.multiselect(
                    "Reviewed Tags (Saved):",
                    options=all_options,
                    default=[t for t in current_reviewed_list if t in all_options],
                    key=f"db_view_{record_id}"
                )
                
                if sorted(edited_reviewed_tags) != sorted(current_reviewed_list):
                    st.warning("Changes detected in reviewed tags.")
                    if st.button("Confirm Changes to Reviewed Tags", key=f"btn_update_db_{record_id}"):
                        new_csv = ",".join(edited_reviewed_tags)
                        current_notes = row['tagging_notes'] if pd.notna(row['tagging_notes']) else ""
                        update_record(current_db, record_id, new_csv, current_notes)
                        st.success("Updated.")
                        st.rerun()

            st.markdown("---")
            
            with st.expander("Enter Tagging Notes (Optional)", expanded=False):
                notes_val = st.text_area(
                    "Enter notes here:", 
                    value=row['tagging_notes'] if pd.notna(row['tagging_notes']) else "",
                    height=100,
                    key=f"notes_{record_id}"
                )
                
                db_notes = row['tagging_notes'] if pd.notna(row['tagging_notes']) else ""
                
                if notes_val != db_notes:
                    if st.button("Save Notes", key=f"save_notes_{record_id}"):
                        curr_tags = row['reviewed_tags'] if pd.notna(row['reviewed_tags']) else ""
                        update_record(current_db, record_id, curr_tags, notes_val)
                        st.success("Notes saved.")
                        st.rerun()

            st.markdown("---")
            
            # --- Bottom Navigation ---
            col_n1, col_n2, col_n3, col_n4 = st.columns([1, 1, 2, 4])
            with col_n1:
                st.button("Previous", on_click=prev_record, disabled=(current_idx == 0))
            with col_n2:
                st.button("Next", on_click=next_record, disabled=(current_idx == len(df_filtered)-1))
            with col_n3:
                st.button("Next Unreviewed", on_click=next_unreviewed)
            with col_n4:
                st.write(f"Record {current_idx + 1} of {len(df_filtered)}")
                
            st.markdown("---")
            # Create the data as a list of dictionaries

            # Full dataset from the image
            data = [
                {"Issue Tag": "SAS/AP", "Feedback Example Text": "The SAS/AP pathways were designed to provide exceptional access to unapproved products, not to facilitate routine prescribing."},
                {"Issue Tag": "Regulatory Principles", "Feedback Example Text": "There should not be any loosening of restrictions for medicinal cannabis products in the transition to the ARTG."},
                {"Issue Tag": "ARTG (Australian Register of Therapeutic Goods)", "Feedback Example Text": "This would include the ordinary course of requiring registration on the ARTG with TGA assessment for quality, safety, efficacy of the product including availability of clinical and scientific evidence, and the labelling and packaging to be regulated by the TGA, all prior to supply of the product on the Australian market."},
                {"Issue Tag": "Transitional Pathway", "Feedback Example Text": "Supports a transitional, time-limited pathway that allows sponsors of unapproved products to continue supply while generating clinical evidence to support ARTG registration."},
                {"Issue Tag": "Scheduling", "Feedback Example Text": "For example, currently CBG-only products are listed in Schedule 8 as category 5, despite having minimal psychoactive effect."},
                {"Issue Tag": "Evidence", "Feedback Example Text": "It is also suggested that high-quality, product-specific research that might provide definitive guidance for efficacy in certain conditions, use of specific formulations or doses, or lead to ARTG listing, has not eventuated as expected."},
                {"Issue Tag": "Enforcement", "Feedback Example Text": "The TGA could also regulate the labelling and packaging of unregistered medicinal cannabis products such that it would be unlawful for the sponsor or manufacturer to supply the product under SAS or AP if not labelled or packaged in accordance with the Poisons Standard, in the same way that other medicines are required to be labelled by the sponsor or manufacturer."},
                {"Issue Tag": "Labelling", "Feedback Example Text": "The TGA could also regulate the labelling and packaging of unregistered medicinal cannabis products such that it would be unlawful for the sponsor or manufacturer to supply the product under SAS or AP if not labelled or packaged in accordance with the Poisons Standard, in the same way that other medicines are required to be labelled by the sponsor or manufacturer."},
                {"Issue Tag": "Specialist Oversight", "Feedback Example Text": "Collaboration between the TGA, AHPRA, and States/Territories to ensure regulatory consistency and oversight of high-volume prescribers of unapproved medicinal cannabis products."},
                {"Issue Tag": "Clinical Guidance", "Feedback Example Text": "Nationally consistent guidance and education for clinicians and consumers of unapproved medicinal cannabis products."},
                {"Issue Tag": "Quality Standards", "Feedback Example Text": "Regulatory measures to help mitigate these risks may include full ingredient disclosure, batch testing, per-dose limits, and the provision of additional clinical guidance."},
                {"Issue Tag": "Fee Waiver", "Feedback Example Text": "Incentives for sponsors of unapproved medicinal cannabis products to pursue ARTG registration (e.g. fee waivers)."},
                {"Issue Tag": "PV (Pharmacovigilance)", "Feedback Example Text": "Robust pharmacovigilance for unapproved medicinal cannabis products, including mandatory reporting and linkage to DAEN. Increased emphasis on the requirement to report safety events linked to medicinal cannabis, and communication about this."},
                {"Issue Tag": "Child-resistant Packaging", "Feedback Example Text": "Child proof packaging on edibles"},
                {"Issue Tag": "Advertising", "Feedback Example Text": "Limiting advertising to reduce attractiveness to young people"},
                {"Issue Tag": "Education", "Feedback Example Text": "Better education for prescribers"},
                {"Issue Tag": "GMP (Good Manufacturing Practice)", "Feedback Example Text": "Requiring imported products to match domestic standards of “Good Manufacturing Practice”"},
                {"Issue Tag": "THC limit", "Feedback Example Text": "Limit THC content significantly in the current Schedule 8 entries for cannabis and THC."},
                {"Issue Tag": "No THC Access", "Feedback Example Text": "Strict regulation of high potency THC products - including no access to these until demonstrable benefit from clinical trials."},
                {"Issue Tag": "Efficacy", "Feedback Example Text": "There is far too easy access to these products through SS and AP, which have not demonstrated efficacy through the appropriate scientific processes."},
                {"Issue Tag": "Restrict dosage forms", "Feedback Example Text": "Restricting access to formulations that appeal to children and pets."},
                {"Issue Tag": "Inhalation or Vapourisation", "Feedback Example Text": "Supports TGA implementing measures to further improve the safety of medicinal cannabis products by removing access via SAS and AP to dosage forms that have evidence of harm or a lack of safety data, especially for those containing high concentrations of THC, such as liquids for vaporisation or granules. There is no clinical justifiable indication for those dosage forms when alternative forms are available with a more established safety profile."},
                {"Issue Tag": "Cannabinoids", "Feedback Example Text": "Restricting access to non-CBD/THC cannabinoids for which there is adequate safety data"},
                {"Issue Tag": "Categories", "Feedback Example Text": "Category 5 products are not those that are >98% THC by weight, rather the cannabinoid content of the products of this category is >98% THC. New categories that would better describe the level of concern related to a product would be more useful."},
                {"Issue Tag": "HREC (Human Research Ethics Committees)", "Feedback Example Text": "The process around Human Ethics Research Committee (HREC) assessment of medical practitioners for AP is not considered to be readily accessible or transparent."},
                {"Issue Tag": "Access", "Feedback Example Text": "The Department understands that the intent of the current SAS and AP schemes are to ensure “right of access” to therapeutic goods that were not marketed in Australia, primarily for economic reasons."},
                {"Issue Tag": "Active Ingredient", "Feedback Example Text": "The primary issue is the widespread confusion across different sectors—from healthcare practitioners to software developers—about how a cannabinoid's \"active\" status relates to the product's overall schedule and category. The current system creates a disconnect between the TGA's definition of an active ingredient (based on a percentage of the total product) and the scheduling criteria (based on a percentage of the total cannabinoids). For example, a product might have a low concentration of THC (e.g., less than 1%) and a low concentration of CBG. Even if neither cannabinoid meets the \"active ingredient\" threshold for the overall product, the CBG might constitute a significant percentage of the total cannabinoid content."}
            ]
            # Convert to DataFrame
            df = pd.DataFrame(data)

            # Apply text wrapping and hide index
            styled_df = df.style.hide(axis="index").set_table_styles([
                {'selector': 'td', 'props': [('white-space', 'normal'), ('word-wrap', 'break-word'), ('max-width', '400px')]},
                {'selector': 'th', 'props': [('white-space', 'normal'), ('word-wrap', 'break-word'), ('max-width', '200px')]}
            ])

            with st.expander("Masterlist of Regulatory Pathway Tags"):
                st.table(styled_df)

            
# ==========================================
# TAB 3: Export Data
# ==========================================
with tab3:
    st.header("Export Data")
    
    if 'active_db' not in st.session_state or not st.session_state['active_db'] or not os.path.exists(get_db_path(st.session_state['active_db'])):
        st.warning("Please load a database in Data Management tab.")
    else:
        current_db = st.session_state['active_db']
        
        # --- Local Filters for Tab 3 ---
        st.subheader("Export Filters")
        
        # Get options independently
        source_opts_ex = ["All"] + get_distinct_values(current_db, "source")
        label_opts_ex = ["All"] + get_distinct_values(current_db, "WorksheetLabelArray")
        
        col_f3, col_f4 = st.columns(2)
        with col_f3:
            sel_source_t3 = st.selectbox("Filter by Source", source_opts_ex, key="t3_source")
        with col_f4:
            sel_label_t3 = st.selectbox("Filter by Worksheet Label", label_opts_ex, key="t3_label")
        
        st.info(f"Exporting records from: **{current_db}**")
        st.write(f"Selected Filters - Source: **{sel_source_t3}**, Label: **{sel_label_t3}**")
        
        # Load data based on filters
        df_export = load_data_from_db(current_db, sel_source_t3, sel_label_t3)
        
        st.write(f"Records to Export: {len(df_export)}")
        
        if len(df_export) > 0:
            
        # --- Option 1: CSV Export (Simplified) ---
            st.subheader("Option 1: Export filtered records as CSV")
            st.caption("Exports: Source, Worksheet Label, Proposed Tags, Reviewed Tags, Notes")
            
            # Create Proposed Tags column by combining IssueTag1-8
            def combine_tags(row):
                tags = []
                for i in range(1, 9):
                    t = row.get(f"IssueTag{i}")
                    if pd.notna(t) and str(t).strip() != "":
                        tags.append(str(t).strip())
                return ", ".join(tags)

            df_csv = df_export.copy()
            df_csv['Proposed LLM Tags'] = df_csv.apply(combine_tags, axis=1)
            
            # Select columns
            cols_to_keep = ['source', 'WorksheetLabelArray', 'Proposed LLM Tags', 'reviewed_tags', 'tagging_notes']
            # cols_to_keep = [
            #     "source", "Issue", "StakeholderTypeArray", "WorksheetLabelArray",
            #     "IssueTag1", "IssueTag2", "IssueTag3", "IssueTag4",
            #     "IssueTag5", "IssueTag6", "IssueTag7", "IssueTag8", 'Proposed LLM Tags', 'reviewed_tags', 'tagging_notes'
            # ]
            # Ensure columns exist before selecting
            valid_cols = [c for c in cols_to_keep if c in df_csv.columns]
            df_csv_final = df_csv[valid_cols]
            
            csv_data = df_csv_final.to_csv(index=False).encode('utf-8')
            
            st.download_button(
                label="Download CSV",
                data=csv_data,
                file_name=f"reviewed_tags_subset_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv"
            )

            st.markdown("---")
            
            # --- Option 2: Excel Export (Full) ---
            st.subheader("Option 2: Export Excel (Full)")
            st.caption("Exports original file columns + Reviewed Tags + Notes")
            df_export = df_export.drop(columns=["record_id"])
            
            excel_data = to_excel(df_export)
            
            st.download_button(
                label="Download Excel (.xlsx)",
                data=excel_data,
                file_name=f"reviewed_full_data_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No records found with current filters to export.")