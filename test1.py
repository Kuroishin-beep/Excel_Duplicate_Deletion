import streamlit as st
import pandas as pd
from io import BytesIO
import streamlit.components.v1 as components

st.set_page_config(layout="wide")
st.title("Excel Duplicate Row Remover")

# ------------------------ UPLOAD FILE ------------------------
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    # Load Excel
    df = pd.read_excel(uploaded_file)

    # Normalize column names
    df.columns = (
        df.columns
        .str.strip()
        .str.lower()
        .str.replace(" ", "_")
        .str.replace("-", "_")
    )

    # Validate required fields exist
    required_cols = {"account", "narration", "description"}
    missing = required_cols - set(df.columns)

    if missing:
        st.error(f"Missing required columns: {', '.join(missing)}")
        st.stop()

    # Initialize session states
    st.session_state.setdefault("delete_rows", set())
    st.session_state.setdefault("confirmed", False)

    st.subheader("Search Rows to Delete")

    search_input = st.text_input("Enter row numbers or keywords (e.g., '15', '100,101', 'sales', '5001')")

    if search_input:
        search_input = search_input.strip()
        found_rows = set()

        # If input is row numbers
        if search_input.replace(",", "").replace(" ", "").isdigit():
            nums = [int(x) for x in search_input.split(",")]
            for n in nums:
                if n in df.index:
                    found_rows.add(n)
        else:
            mask = (
                df["account"].astype(str).str.contains(search_input, case=False, na=False) |
                df["narration"].astype(str).str.contains(search_input, case=False, na=False) |
                df["description"].astype(str).str.contains(search_input, case=False, na=False)
            )
            found_rows = set(df[mask].index)

        st.session_state.delete_rows.update(found_rows)

    # -------------- PANEL LAYOUT -----------------
    col1, col2, col3 = st.columns([3, 2, 3])

    # ------------------------ LEFT PANEL ------------------------
    with col1:
        st.subheader("Original Data (Scrollable)")

        def highlight_rows(row):
            if row.name in st.session_state.delete_rows:
                return ['background-color: #ffbbbb'] * len(row)
            return [''] * len(row)

        st.dataframe(df.style.apply(highlight_rows, axis=1), height=600)

    # ------------------------ MIDDLE PANEL ------------------------
    with col2:
        st.subheader("Rows Marked for Deletion")

        if st.session_state.delete_rows:
            delete_df = df.loc[sorted(st.session_state.delete_rows)]
            st.dataframe(delete_df, height=400)

            # Jump to row
            selected_row = st.selectbox(
                "Jump to row",
                options=["None"] + list(delete_df.index),
                key="row_jump"
            )

            if selected_row != "None":
                st.session_state.selected_row = selected_row

            # Remove row from deletion list
            row_to_remove = st.selectbox(
                "Remove from delete list",
                options=["None"] + list(delete_df.index),
                key="remove_row"
            )

            if row_to_remove != "None":
                st.session_state.delete_rows.remove(row_to_remove)
                st.success(f"Row {row_to_remove} removed.")
        else:
            st.info("No rows selected yet.")

    # ------------------------ RIGHT PANEL ------------------------
    with col3:
        st.subheader("Cleaned Preview")

        if st.button("Confirm Row Deletion"):
            st.session_state.confirmed = True
            st.success("Rows validated. Clean Excel file ready!")

        if st.session_state.confirmed:
            clean_df = df.drop(index=st.session_state.delete_rows)

            def highlight_selected(row):
                if "selected_row" in st.session_state and row.name == st.session_state.selected_row:
                    return ['background-color: yellow'] * len(row)
                return [''] * len(row)

            st.dataframe(clean_df.style.apply(highlight_selected, axis=1), height=600)

            # Auto-scroll JS
            if "selected_row" in st.session_state:
                components.html(
                    f"""
                    <script>
                    const rowIndex = {st.session_state.selected_row};
                    setTimeout(() => {{
                        const rows = window.parent.document.querySelectorAll('.stDataFrame table tbody tr');
                        if (rows[rowIndex]) {{
                            rows[rowIndex].scrollIntoView({{ behavior: "smooth", block: "center" }});
                        }}
                    }}, 500);
                    </script>
                    """,
                    height=0,
                )

            # Download Excel
            buffer = BytesIO()
            clean_df.to_excel(buffer, index=False)
            buffer.seek(0)

            st.download_button(
                "Download Clean Excel",
                data=buffer,
                file_name="cleaned_file.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("Click **Confirm Row Deletion** to generate preview.")
else:
    st.info("Upload an Excel file to continue.")
