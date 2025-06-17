import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches
import io
# from docx.oxml import OxmlElement # Not strictly needed for the current update_specific_table_cell
# from docx.text.paragraph import Paragraph # Not strictly needed

# --- Function to Replace Text in a Word Document (from previous steps) ---
def replace_text_in_document(document_stream, old_text, new_text):
    try:
        document = Document(document_stream)
        modified = False
        for paragraph in document.paragraphs:
            if old_text in paragraph.text:
                for run in paragraph.runs:
                    if old_text in run.text:
                        run.text = run.text.replace(old_text, new_text)
                        modified = True
        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if old_text in paragraph.text:
                            for run in paragraph.runs:
                                if old_text in run.text:
                                    run.text = run.text.replace(old_text, new_text)
                                    modified = True
        if not modified and old_text:
             st.warning(f"Placeholder '{old_text}' not found in the document.")
        bio = io.BytesIO()
        document.save(bio)
        bio.seek(0)
        return bio
    except Exception as e:
        st.error(f"Error processing Word document for text replacement: {e}")
        return None

def update_specific_table_cell(document_stream, table_idx, row_idx, col_idx, new_value):
    """
    Updates the content of a specific cell in a specific table within a Word document.

    Args:
        document_stream (io.BytesIO): A byte stream of the .docx file.
        table_idx (int): The 0-based index of the table in the document.
        row_idx (int): The 0-based index of the row in the table.
        col_idx (int): The 0-based index of the column in the table.
        new_value (str): The new text value for the cell.

    Returns:
        io.BytesIO: A byte stream of the modified Word document, or None if an error occurs.
    """
    try:
        document = Document(document_stream)

        if not (0 <= table_idx < len(document.tables)):
            st.error(f"Error: Table index {table_idx} is out of bounds. Document has {len(document.tables)} tables.")
            return None

        table = document.tables[table_idx]

        if not (0 <= row_idx < len(table.rows)):
            st.error(f"Error: Row index {row_idx} is out of bounds for table {table_idx}. Table has {len(table.rows)} rows.")
            return None

        if not (0 <= col_idx < len(table.columns)): # or len(table.rows[row_idx].cells)
            st.error(f"Error: Column index {col_idx} is out of bounds for table {table_idx}, row {row_idx}. Row has {len(table.rows[row_idx].cells)} cells.")
            return None

        cell_to_update = table.cell(row_idx, col_idx)

        # Clear existing content in the cell (important if cell has multiple paragraphs)
        # A simple way is to assign to cell.text, but this might not clear all paragraph elements.
        # More robustly:
        cell_to_update.text = "" # Clears existing text and adds an empty paragraph

        # Or, to be absolutely sure all paragraphs are gone before adding new one:
        # (This is more aggressive and might be needed if .text = "" doesn't suffice)
        # for p in reversed(cell_to_update.paragraphs): # Iterate reversed to safely remove
        #     p_element = p._element
        #     p_element.getparent().remove(p_element)

        # Add the new value as a new paragraph (or just set cell.text)
        cell_to_update.text = str(new_value)
        # If you want to add multiple paragraphs or runs with specific formatting,
        # you would use cell_to_update.add_paragraph() and run objects here.

        # Save the modified document to an in-memory stream
        bio = io.BytesIO()
        document.save(bio)
        bio.seek(0) # Reset stream position to the beginning
        return bio

    except Exception as e:
        st.error(f"Error updating table cell: {e}")
        return None

def update_full_table(document, table_idx, new_values: pd.DataFrame, styler):
    """
    Updates the content of a specific cell in a specific table within a Word document.

    Args:
        document_stream (io.BytesIO): A byte stream of the .docx file.
        table_idx (int): The 0-based index of the table in the document.
        row_idx (int): The 0-based index of the row in the table.
        col_idx (int): The 0-based index of the column in the table.
        new_value (str): The new text value for the cell.

    Returns:
        io.BytesIO: A byte stream of the modified Word document, or None if an error occurs.
    """
    try:
        if not (0 <= table_idx < len(document.tables)):
            st.error(f"Error: Table index {table_idx} is out of bounds. Document has {len(document.tables)} tables.")
            return None

        table = document.tables[table_idx]
        new_values_row_count = new_values.shape[0]
        new_values_col_count = new_values.shape[1]

        if not (new_values_row_count == len(table.rows)):
            st.error(f"Error: Word-Table has {len(table.rows)} rows and Excel-Table has {new_values_row_count}. Exceeding rows not replaced!")
            new_values_row_count = len(table.rows)

        if not (new_values_col_count ==  len(table.columns)): # or len(table.rows[row_idx].cells)
            st.error(f"Error: Word-Table has {len(table.columns)} columns and Excel-Table has {new_values_col_count}.")
            new_values_col_count = len(table.columns)

        for row_idx in range(new_values_row_count):
            for col_idx in range(new_values_col_count):
                cell_to_update = table.cell(row_idx, col_idx)
                paragraph = cell_to_update.paragraphs[0]
                # Clear existing content in the cell (important if cell has multiple paragraphs)
                # A simple way is to assign to cell.text, but this might not clear all paragraph elements.
                # More robustly:
                #cell_to_update.text = "" # Clears existing text and adds an empty paragraph

                # Or, to be absolutely sure all paragraphs are gone before adding new one:
                # (This is more aggressive and might be needed if .text = "" doesn't suffice)
                # for p in reversed(cell_to_update.paragraphs): # Iterate reversed to safely remove
                #     p_element = p._element
                #     p_element.getparent().remove(p_element)
                for run in paragraph.runs:
                        run.text = ""

                # Add the new value as a new paragraph (or just set cell.text)
                styl = styler[row_idx]
                if col_idx == 0:
                    styl = "{}"
                print(f"row_idx {row_idx} styler {styler[row_idx]} value {new_values.iloc[row_idx, col_idx]} style {styl.format(new_values.iloc[row_idx, col_idx])}")
                #cell_to_update.text = styl.format(new_values.iloc[row_idx, col_idx])
                paragraph.add_run(styl.format(new_values.iloc[row_idx, col_idx]))
        
        return document

    except Exception as e:
        st.error(f"Error updating table cell: {e}")
        return None


# --- 1. Set up the Streamlit Interface ---
st.set_page_config(page_title="Document Toolkit", layout="wide")
st.title("📄 Document Toolkit")

# --- Tabs for different functionalities ---
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Report Generator from Excel/CSV",
    "🔄 Word Text Replacer",
    "📝 Update Table Cell in Word",
    "📝 EYA Processor"
])

# --- TAB 1: Report Generator from Excel/CSV (from previous steps) ---
with tab1:
    st.header("Generate Word Report from Excel & CSV")
    st.write("Upload your Excel and CSV files to generate a new Word report. This will also attempt to copy a table from 'Sheet B', cells B4:F8 from the Excel file.")

    st.sidebar.header("📤 Upload for Report Generator")
    uploaded_excel_file_tab1 = st.sidebar.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"], key="excel_uploader_tab1")
    uploaded_csv_file_tab1 = st.sidebar.file_uploader("Upload CSV File (.csv)", type=["csv"], key="csv_uploader_tab1")

    if uploaded_excel_file_tab1 and uploaded_csv_file_tab1:
        st.sidebar.success("Excel and CSV files uploaded for Report Generator!")
        try:
            excel_df_main = pd.read_excel(uploaded_excel_file_tab1)
            st.subheader("📊 Preview of Excel Data (First Sheet)")
            st.dataframe(excel_df_main.head())
            csv_df = pd.read_csv(uploaded_csv_file_tab1)
            st.subheader("📊 Preview of CSV Data")
            st.dataframe(csv_df.head())

            excel_table_data_specific = None
            sheet_name_specific = "Sheet B"
            try:
                excel_df_specific_sheet = pd.read_excel(uploaded_excel_file_tab1, sheet_name=sheet_name_specific, header=None)
                excel_table_data_specific = excel_df_specific_sheet.iloc[3:8, 1:6] # Rows 4-8, Columns B-F
                st.subheader(f"📊 Preview of Specific Table from Excel ('{sheet_name_specific}', B4:F8)")
                if not excel_table_data_specific.empty:
                    st.dataframe(excel_table_data_specific)
                else:
                    st.warning(f"Range B4:F8 in '{sheet_name_specific}' is empty or out of bounds.")
            except Exception as e:
                st.warning(f"Could not read '{sheet_name_specific}' or range B4:F8 from the Excel file. Error: {e}")
                excel_table_data_specific = None

            if st.button("🚀 Generate Report from Excel/CSV", key="generate_report_button"):
                with st.spinner("Generating your report... Please wait."):
                    document = Document()
                    document.add_heading('Automated Report from Data Files', level=1)

                    # Add data from the FIRST sheet of Excel file
                    document.add_heading('Excel Data Summary (First Sheet)', level=2)
                    document.add_paragraph(f"The first sheet of the Excel file contains {excel_df_main.shape[0]} rows and {excel_df_main.shape[1]} columns.")
                    if not excel_df_main.empty:
                        document.add_paragraph("First 5 rows of Excel data (First Sheet):")
                        table_excel_main = document.add_table(rows=1, cols=excel_df_main.shape[1])
                        table_excel_main.style = 'Table Grid'
                        hdr_cells_excel_main = table_excel_main.rows[0].cells
                        for i, col_name in enumerate(excel_df_main.columns):
                            hdr_cells_excel_main[i].text = str(col_name)
                        for _, row in excel_df_main.head().iterrows():
                            row_cells = table_excel_main.add_row().cells
                            for i, cell_value in enumerate(row):
                                row_cells[i].text = str(cell_value)
                    else:
                        document.add_paragraph("First sheet of Excel file is empty or could not be read.")

                    # Add the specific table from Excel 'Sheet B', B4:F8
                    if excel_table_data_specific is not None and not excel_table_data_specific.empty:
                        document.add_heading(f"Table from Excel: '{sheet_name_specific}' (Cells B4:F8)", level=2)
                        try:
                            num_rows, num_cols = excel_table_data_specific.shape
                            table_specific = document.add_table(rows=num_rows, cols=num_cols)
                            table_specific.style = 'Table Grid'
                            for i in range(num_rows):
                                for j in range(num_cols):
                                    cell_value = excel_table_data_specific.iloc[i, j]
                                    table_specific.cell(i, j).text = str(cell_value if pd.notna(cell_value) else "")
                            document.add_paragraph(f"The extracted table has {num_rows} rows and {num_cols} columns.")
                        except Exception as e:
                            document.add_paragraph(f"Error adding specific Excel table to report: {e}")
                            st.error(f"Error occurred while adding specific table to Word doc: {e}")
                    else:
                        document.add_paragraph(f"No data extracted from '{sheet_name_specific}' range B4:F8 to add to the report.")

                    # Add data from CSV file
                    document.add_heading('CSV Data Summary', level=2)
                    document.add_paragraph(f"The CSV file contains {csv_df.shape[0]} rows and {csv_df.shape[1]} columns.")
                    if not csv_df.empty:
                        document.add_paragraph("First 5 rows of CSV data:")
                        table_csv = document.add_table(rows=1, cols=csv_df.shape[1])
                        table_csv.style = 'Table Grid'
                        hdr_cells_csv = table_csv.rows[0].cells
                        for i, col_name in enumerate(csv_df.columns):
                            hdr_cells_csv[i].text = str(col_name)
                        for _, row in csv_df.head().iterrows():
                            row_cells = table_csv.add_row().cells
                            for i, cell_value in enumerate(row):
                                row_cells[i].text = str(cell_value)
                    else:
                        document.add_paragraph("CSV file is empty or could not be read.")

                    document.add_page_break()
                    document.add_paragraph("Report generated by the Document Toolkit App. ✨")
                    bio_report = io.BytesIO()
                    document.save(bio_report)
                    bio_report.seek(0)
                    st.success("🎉 Report Generated Successfully!")
                    st.download_button(
                        label="⬇️ Download Generated Report",
                        data=bio_report,
                        file_name="generated_data_report.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="download_report_button"
                    )
        except Exception as e:
            st.error(f"An error occurred while processing Excel/CSV for report generation: {e}")
    elif uploaded_excel_file_tab1 or uploaded_csv_file_tab1:
        st.sidebar.info("☝️ Please upload both Excel and CSV files for the report generator.")
    else:
        st.info("☝️ Upload Excel and CSV files using the sidebar to generate a new report.")


# --- TAB 2: Word Text Replacer (from previous steps) ---
with tab2:
    st.header("Replace Text in Existing Word Document")
    st.write("Upload a Word document (.docx) and specify the text to be replaced globally.")

    uploaded_word_template_tab2 = st.file_uploader("Upload Word Document Template (.docx)", type=["docx"], key="word_template_uploader_tab2")
    old_text_tab2 = st.text_input("Text to find (placeholder):", placeholder="e.g., {{CLIENT_NAME}}", key="old_text_tab2")
    new_text_tab2 = st.text_input("Text to replace with:", placeholder="e.g., Acme Corp", key="new_text_tab2")

    if uploaded_word_template_tab2 and old_text_tab2:
        if st.button("🔄 Replace Text in Document", key="replace_text_button"):
            if not new_text_tab2:
                st.info("Replacing placeholder with empty text.")
            with st.spinner("Processing your Word document..."):
                document_stream = io.BytesIO(uploaded_word_template_tab2.getvalue())
                modified_document_stream = replace_text_in_document(document_stream, old_text_tab2, new_text_tab2 if new_text_tab2 else "")
                if modified_document_stream:
                    st.success("🎉 Text Replacement Successful!")
                    st.download_button(
                        label="⬇️ Download Modified Word Document",
                        data=modified_document_stream,
                        file_name=f"modified_text_{uploaded_word_template_tab2.name}",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="download_text_replaced_doc"
                    )
    elif uploaded_word_template_tab2 and not old_text_tab2:
        st.warning("Please enter the 'Text to find (placeholder)'.")
    else:
        st.info("☝️ Upload a Word document and specify the text to find and replace.")


# --- TAB 3: Update Table Cell in Word ---
with tab3:
    st.header("Update Specific Table Cell in a Word Document")
    st.write("Upload a Word document (.docx) and specify the table, row, column, and new value for a cell.")

    uploaded_word_doc_tab3 = st.file_uploader("Upload Word Document (.docx)", type=["docx"], key="word_doc_uploader_tab3")
    MAX_PREVIEW_ROWS = 12
    MAX_REVIEW_COLS = 12

    # Initialize current_cell_content
    current_cell_content = ""

    if uploaded_word_doc_tab3:
        # Display table information to help user identify indices
        try:
            doc_for_preview = Document(io.BytesIO(uploaded_word_doc_tab3.getvalue()))
            if doc_for_preview.tables:
                st.write(f"Document contains {len(doc_for_preview.tables)} table(s).")
                for i, table_item in enumerate(doc_for_preview.tables):
                    max_preview_rows = min(MAX_PREVIEW_ROWS, len(table_item.rows))
                    
                    # More robust column count for preview
                    actual_table_cols = 0
                    if table_item.rows:
                        # Get max cells from the first few rows for a more accurate preview column count
                        for r in range(min(len(table_item.rows), MAX_PREVIEW_ROWS)):
                            actual_table_cols = max(actual_table_cols, len(table_item.rows[r].cells))

                    max_preview_cols = min(MAX_REVIEW_COLS, actual_table_cols)

                    st.caption(f"Table {i} (Total Rows: {len(table_item.rows)}, Total Columns (based on cells in first row): {actual_table_cols}) - Previewing up to {max_preview_rows} rows and {max_preview_cols} columns:")
                    
                    # Preview a bit of the table
                    preview_df_data = []
                    for r_idx in range(max_preview_rows):
                        row_data = []
                        current_row_cells = len(table_item.rows[r_idx].cells) if r_idx < len(table_item.rows) else 0
                        preview_cols_for_row = min(max_preview_cols, current_row_cells)

                        for c_idx in range(preview_cols_for_row):
                            try:
                                cell_text = table_item.cell(r_idx, c_idx).text[:50]
                                row_data.append(f"'{cell_text}...'" if len(table_item.cell(r_idx, c_idx).text) > 50 else f"'{cell_text}'")
                            except IndexError:
                                row_data.append("[Error: Cell out of bounds]")
                                break
                        preview_df_data.append(row_data)

                    if preview_df_data:
                        preview_df = pd.DataFrame(preview_df_data)
                        preview_df.columns = [f"Col {k}" for k in range(preview_df.shape[1])]
                        preview_df.index = [f"Row {k}" for k in range(preview_df.shape[0])]
                        st.dataframe(preview_df, use_container_width=True)
            else:
                st.info("The uploaded document does not appear to contain any tables.")

        except Exception as e:
            st.warning(f"Could not preview tables from the document: {e}")

        st.subheader("Specify Cell to Update (0-based index):")
        col1, col2, col3 = st.columns(3)

        # Use st.session_state to manage input values and trigger updates
        # This is crucial for Streamlit to re-run and update the text_area when numbers change
        if "table_idx_tab3_state" not in st.session_state:
            st.session_state.table_idx_tab3_state = 0
        if "row_idx_tab3_state" not in st.session_state:
            st.session_state.row_idx_tab3_state = 0
        if "col_idx_tab3_state" not in st.session_state:
            st.session_state.col_idx_tab3_state = 0

        with col1:
            table_index_tab3 = st.number_input(
                "Table Index:", 
                min_value=0, 
                step=1, 
                value=st.session_state.table_idx_tab3_state, 
                key="table_idx_tab3",
                on_change=lambda: st.session_state.__setitem__("table_idx_tab3_state", st.session_state.table_idx_tab3)
            )
        with col2:
            row_index_tab3 = st.number_input(
                "Row Index:", 
                min_value=0, 
                step=1, 
                value=st.session_state.row_idx_tab3_state, 
                key="row_idx_tab3",
                on_change=lambda: st.session_state.__setitem__("row_idx_tab3_state", st.session_state.row_idx_tab3)
            )
        with col3:
            col_index_tab3 = st.number_input(
                "Column Index:", 
                min_value=0, 
                step=1, 
                value=st.session_state.col_idx_tab3_state, 
                key="col_idx_tab3",
                on_change=lambda: st.session_state.__setitem__("col_idx_tab3_state", st.session_state.col_idx_tab3)
            )
        
        # --- Retrieve and display current cell content ---
        if doc_for_preview.tables:
            try:
                selected_table = doc_for_preview.tables[table_index_tab3]
                selected_cell = selected_table.cell(row_index_tab3, col_index_tab3)
                current_cell_content = selected_cell.text
            except IndexError:
                current_cell_content = "Cell indices out of bounds for the selected table."
            except Exception as e:
                current_cell_content = f"Error retrieving cell content: {e}"
        else:
            current_cell_content = "No tables found in document."

        # Display the text_area with the current cell content as default
        new_cell_value_tab3 = st.text_area(
            "New Cell Value (current content shown below):",
            value=current_cell_content,
            key="new_val_tab3",
            height=100 # Adjust height as needed
        )

        if st.button("⚙️ Update Table Cell", key="update_cell_button"):
            # The st.text_area already has the current content as its value, so
            # an empty check directly on new_cell_value_tab3 might not be what you want
            # You might want to check if the user actually typed something *different*
            # or if the default value is still the initial "Cell indices out of bounds..." message.
            
            # A simple check if the text area is empty after user interaction
            if not new_cell_value_tab3.strip(): # .strip() to account for whitespace
                st.warning("Please enter a new value for the cell. The field cannot be empty.")
            elif new_cell_value_tab3 == current_cell_content and "Cell indices out of bounds" not in current_cell_content:
                st.info("The new cell value is the same as the current cell value. No change will be made.")
            else:
                with st.spinner("Updating table cell..."):
                    document_stream_tab3 = io.BytesIO(uploaded_word_doc_tab3.getvalue())
                    modified_doc_stream_tab3 = update_specific_table_cell(
                        document_stream_tab3,
                        table_index_tab3,
                        row_index_tab3,
                        col_index_tab3,
                        new_cell_value_tab3
                    )

                    if modified_doc_stream_tab3:
                        st.success("🎉 Table Cell Updated Successfully!")
                        st.download_button(
                            label="⬇️ Download Updated Word Document",
                            data=modified_doc_stream_tab3,
                            file_name=f"cell_updated_{uploaded_word_doc_tab3.name}",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key="download_cell_updated_doc"
                        )
                    # Errors are handled within update_specific_table_cell and shown via st.error
    else:
        st.info("☝️ Upload a Word document to begin updating table cells.")


with tab4:
    st.header("Update Yield Report from Excel")
    st.write("Upload EYA Excel sheet to fill in the information into the Word report template.")

    uploaded_EYA_excel = st.file_uploader("Upload EYA Excel (.xlsm)", type=["xlsm"], key="uploaded_EYA_excel")
    try:
        document = Document('EYA_TEMPLATE.docx')
    except Exception as e:
        st.warning(f"Opening EYA Template Word file failed: {e}")
    if uploaded_EYA_excel:
        excel_table_data_specific = None
        sheet_name_specific = "Probability scenarios 1"
        try:
            excel_df_specific_sheet = pd.read_excel(uploaded_EYA_excel, sheet_name=sheet_name_specific, header=None,
                                                    usecols="G:O",skiprows=25, nrows=12)
            excel_table_data_specific = excel_df_specific_sheet.iloc[:,:] # Rows 4-8, Columns B-K
            print(excel_table_data_specific.columns)
            print(excel_table_data_specific.index)
            for i,c_idx in enumerate(excel_table_data_specific.columns):
                if i >0 :
                    excel_table_data_specific[c_idx]=pd.to_numeric(excel_table_data_specific[c_idx], errors='coerce')
            st.subheader(f"📊 Preview of Specific Table from Excel ('{sheet_name_specific}', B25:K36)")
            if not excel_table_data_specific.empty:
                st.dataframe(excel_table_data_specific)
            else:
                st.warning(f"Range B25:K36 in '{sheet_name_specific}' is empty or out of bounds.")
        except Exception as e:
            st.warning(f"Could not read '{sheet_name_specific}' or range B25:K36 from the Excel file. Error: {e}")
            excel_table_data_specific = None

        if st.button("⚙️ Update Table Cell", key="update_cell_button_tab4"):
        # The st.text_area already has the current content as its value, so
        # an empty check directly on new_cell_value_tab3 might not be what you want
        # You might want to check if the user actually typed something *different*
        # or if the default value is still the initial "Cell indices out of bounds..." message.

            with st.spinner("Updating table cell..."):
                tab5_styler = ["{:.0f}","{:.2%}","{:.0%}","{:.2%}","{:.2%}","{:.2%}","{:.2%}","{:,.0f}","{:,.0f}","{:,.0f}","{:,.0f}","{:,.0f}"]    
                document = update_full_table(
                    document,
                    5,
                    excel_table_data_specific,
                    tab5_styler
                )
                document = update_full_table(
                    document,
                    16,
                    excel_table_data_specific,
                    tab5_styler
                )
                if document:
                    st.success("🎉 Table Cell Updated Successfully!")
                            # Save the modified document to an in-memory stream
                    bio = io.BytesIO()
                    document.save(bio)
                    bio.seek(0) # Reset stream position to the beginning
                    st.download_button(
                        label="⬇️ Download Updated Word Document",
                        data=bio,
                        file_name=f"cell_updated_.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="download_cell_updated_doc_tab4"
                    )
                        # Errors are handled within update_specific_table_cell and shown via st.error
    else:
        st.info("☝️ Upload a Excel to begin updating table cells.")

# --- Styling (remains the same) ---
st.markdown("""
    <style>
    .stButton>button {
        border-radius: 5px;
        padding: 8px 15px;
        margin-top: 10px;
    }
    .stDownloadButton>button {
        background-color: #008CBA;
        color: white;
        border-radius: 5px;
        padding: 8px 15px;
        margin-top: 5px;
    }
    .stTabs [data-baseweb="tab-list"] {
		gap: 24px;
	}
	.stTabs [data-baseweb="tab"] {
		height: 50px;
        white-space: pre-wrap;
		background-color: #F0F2F6;
		border-radius: 4px 4px 0px 0px;
		gap: 1px;
		padding-top: 10px;
		padding-bottom: 10px;
    }
    </style>
""", unsafe_allow_html=True)