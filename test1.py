import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# -----------------------------
# HELPER FUNCTIONS
# -----------------------------
def get_rows_to_delete_logic(df, search_term):
    """
    Logic:
    1. Find rows containing the search term (CASE SENSITIVE).
    2. Check the row immediately below; if it contains "Total" (case insensitive), mark it too.
    """
    if not search_term:
        return []

    # 1. Find main matches (CASE SENSITIVE)
    mask = df.astype(str).apply(lambda x: x.str.contains(search_term, case=True, na=False)).any(axis=1)
    matched_indices = df[mask].index.tolist()
    
    final_deletion_set = set(matched_indices)
    
    # 2. Look for "Total" rows immediately below matches
    for idx in matched_indices:
        if idx + 1 < len(df):
            next_row = df.iloc[idx + 1]
            row_content = str(next_row.values).lower()
            if "total" in row_content:
                final_deletion_set.add(idx + 1)

    return sorted(list(final_deletion_set))

def process_excel_with_formatting(uploaded_file, indices_to_delete):
    """
    Uses OpenPyXL to delete rows while preserving styles.
    """
    uploaded_file.seek(0)
    wb = openpyxl.load_workbook(uploaded_file)
    ws = wb.active
    
    # Convert Pandas Index (0-based) to Excel Row (1-based + Header)
    excel_rows_to_delete = [i + 2 for i in indices_to_delete]
    excel_rows_to_delete.sort(reverse=True)
    
    for r in excel_rows_to_delete:
        ws.delete_rows(r)
    
    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# -----------------------------
# SESSION STATE INITIALIZATION
# -----------------------------
if 'df_original' not in st.session_state:
    st.session_state.df_original = None
    
# THE QUEUE: This is where we cache rows found in multiple searches
if 'deletion_queue' not in st.session_state:
    st.session_state.deletion_queue = set() 

# THE CURRENT SEARCH: Temporary matches from the current text box
if 'current_matches' not in st.session_state:
    st.session_state.current_matches = []

# -----------------------------
# PAGE LAYOUT
# -----------------------------
st.set_page_config(layout="wide")
st.title("Excel Duplicate Delete")

# STEP 1: UPLOAD EXCEL
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

if uploaded_file:
    # Only load dataframe once
    if st.session_state.df_original is None:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        st.session_state.df_original = df.copy()

    col1, col2, col3 = st.columns([3, 2, 3])

    # -----------------------------
    # STEP 2: SEARCH & ACCUMULATE
    # -----------------------------
    with col1:
        st.subheader("1. Find & Add to Queue")
        
        # SEARCH INPUT
        search_text = st.text_input("Search text (Case Sensitive)", placeholder="Type exact text...")

        # LOGIC: Run search immediately
        if search_text:
            found_indices = get_rows_to_delete_logic(st.session_state.df_original, search_text)
            st.session_state.current_matches = found_indices
        else:
            st.session_state.current_matches = []

        # ADD BUTTON
        match_count = len(st.session_state.current_matches)
        if match_count > 0:
            st.success(f"Found {match_count} rows.")
            if st.button(f"➕ Add {match_count} Rows to Deletion Queue"):
                # Add current matches to the permanent set
                st.session_state.deletion_queue.update(st.session_state.current_matches)
                st.session_state.current_matches = [] # Clear current selection visual
                st.rerun()

        # DATAFRAME WITH DUAL HIGHLIGHTING
        df_display = st.session_state.df_original.copy()

        def highlight_logic(row):
            # If row is already in the Queue -> RED
            if row.name in st.session_state.deletion_queue:
                return ['background-color: #ffcccc'] * len(row)
            # If row is found by current search -> YELLOW
            elif row.name in st.session_state.current_matches:
                return ['background-color: #ffffcc'] * len(row)
            else:
                return [''] * len(row)

        try:
            st.dataframe(df_display.style.apply(highlight_logic, axis=1), height=500)
        except:
            st.dataframe(df_display, height=500)

    # -----------------------------
    # STEP 3: REVIEW QUEUE
    # -----------------------------
    with col2:
        st.subheader("2. Review Deletion Queue")
        
        queue_list = sorted(list(st.session_state.deletion_queue))
        
        if queue_list:
            st.warning(f"Total Rows in Queue: {len(queue_list)}")
            
            # Show the rows currently in Queue
            delete_df = st.session_state.df_original.iloc[queue_list].copy()
            delete_df.insert(0, "Excel_Row", [i + 2 for i in queue_list])
            st.dataframe(delete_df, height=300)

            col2a, col2b = st.columns(2)
            
            with col2a:
                # RESCUE LOGIC
                st.write("**Rescue Row:**")
                row_to_rescue = st.selectbox(
                    "Select Excel Row to KEEP:", 
                    options=[i + 2 for i in queue_list],
                    key="rescue_box"
                )
                if st.button("Rescue Selected"):
                    idx_to_remove = row_to_rescue - 2
                    if idx_to_remove in st.session_state.deletion_queue:
                        st.session_state.deletion_queue.remove(idx_to_remove)
                        st.rerun()
            
            with col2b:
                # CLEAR ALL LOGIC
                st.write("**Reset:**")
                if st.button("Clear Entire Queue"):
                    st.session_state.deletion_queue = set()
                    st.rerun()

        else:
            st.info("Queue is empty. Search and add rows.")

    # -----------------------------
    # STEP 4: PREVIEW & DOWNLOAD
    # -----------------------------
    with col3:
        st.subheader("3. Preview & Download")
        
        queue_list = sorted(list(st.session_state.deletion_queue))
        
        # Preview drops whatever is in the Queue
        df_preview = st.session_state.df_original.drop(queue_list, errors='ignore')
        st.dataframe(df_preview, height=500)
        
        st.write("---")
        
        if queue_list:
            if st.button("Confirm Deletion & Download Excel"):
                with st.spinner("Processing with OpenPyXL..."):
                    processed_data = process_excel_with_formatting(uploaded_file, queue_list)
                    
                    st.success("Done!")
                    st.download_button(
                        label="Download Final Excel",
                        data=processed_data,
                        file_name="cleaned_data_preserved.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )