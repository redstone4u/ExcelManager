import streamlit as st
import pandas as pd
import io

# Set page configuration
st.set_page_config(layout="wide")

# Title and description
st.title("Excel File Merger and Converter App tested by 정민규")

# Description
st.write(
    "This app allows you to merge and convert Excel files. "
    "You can upload multiple files, select sheets to merge, "
    "rearrange columns, and save the result."
)

# Initialize session state
uploaded_files = st.file_uploader(
    "Choose Excel files",
    type="xlsx",
    accept_multiple_files=True
)

# Step 1: Upload and select sheets
if uploaded_files:
    selected_sheets = {}
    for file in uploaded_files:
        try:
            xls = pd.ExcelFile(file)
            sheet_names = xls.sheet_names
            selected_sheets[file.name] = st.multiselect(
                f"Select sheets from {file.name}",
                sheet_names,
                default=sheet_names[0] if sheet_names else []
            )
        except Exception as e:
            st.error(f"Error reading {file.name}: {e}")

    if st.button("Merge Files"):
        all_dfs = []
        for file in uploaded_files:
            if file.name in selected_sheets:
                for sheet_name in selected_sheets[file.name]:
                    try:
                        df = pd.read_excel(file, sheet_name=sheet_name)
                        all_dfs.append(df)
                    except Exception as e:
                        st.error(f"Error reading sheet {sheet_name} from {file.name}: {e}")
        
        if all_dfs:
            merged_df = pd.concat(all_dfs, ignore_index=True)
            st.session_state.merged_df = merged_df
            st.success("Files merged successfully!")

# Step 2: Arrange columns
if 'merged_df' in st.session_state:
    st.subheader("Arrange Columns")
    merged_df = st.session_state.merged_df
    all_columns = merged_df.columns.tolist()
    
    # Allow users to select and reorder columns
    selected_columns = st.multiselect(
        "Select and reorder columns",
        all_columns,
        default=all_columns
    )
    
    # Create a new DataFrame with the selected columns in the specified order
    arranged_df = merged_df[selected_columns]
    st.session_state.arranged_df = arranged_df

# Step 3: Preview and edit data
if 'arranged_df' in st.session_state:
    st.subheader("Preview and Edit Data")
    arranged_df = st.session_state.arranged_df
    
    # Display the DataFrame in an editable table
    edited_df = st.data_editor(arranged_df)
    st.session_state.edited_df = edited_df

# Step 4: Save result
if 'edited_df' in st.session_state:
    st.subheader("Save Result")
    edited_df = st.session_state.edited_df
    
    file_name = st.text_input("Enter file name", "merged_data.xlsx")
    
    if st.button("Save to Excel"):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            edited_df.to_excel(writer, index=False, sheet_name='Sheet1')
        
        st.download_button(
            label="Download Excel file",
            data=output.getvalue(),
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success(f"File saved as {file_name}")

