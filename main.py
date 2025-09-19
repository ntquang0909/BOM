from io import BytesIO

import pandas as pd
import streamlit as st


def to_excel(df):
    """
    Helper function to convert a DataFrame to an Excel file in memory.
    """
    output = BytesIO()
    # Use the 'xlsxwriter' engine for better compatibility and to avoid deprecation warnings.
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Merged_BOM")
    processed_data = output.getvalue()
    return processed_data


# --- Streamlit App Configuration ---
st.set_page_config(page_title="Excel BOM Merger", layout="wide")

st.title("Bill of Materials (BOM) Merger Tool")
st.write(
    "Upload your Excel BOM files. The tool will merge them by summing the 'QTY.' for items with the same 'DESCRIPTION' and 'LENGTH'."
)

# --- File Uploader ---
uploaded_files = st.file_uploader(
    "Choose your Excel files", type=["xlsx", "xls"], accept_multiple_files=True
)

# --- Main Processing Logic ---
if uploaded_files:
    st.info(f"{len(uploaded_files)} file(s) uploaded successfully.")

    all_dataframes = []
    error_files = []

    # --- Step 1: Read and Clean Each Uploaded File ---
    for uploaded_file in uploaded_files:
        try:
            df = pd.read_excel(uploaded_file)

            # --- Data Cleaning ---
            # Standardize column names by removing leading/trailing whitespace.
            df.columns = df.columns.str.strip()

            # Fill any missing 'DESCRIPTION' or 'LENGTH' values with an empty string to ensure they can be grouped.
            df["DESCRIPTION"] = df["DESCRIPTION"].fillna("").str.strip()
            df["LENGTH"] = df["LENGTH"].fillna("")

            # Ensure 'PART NUMBER' is a string and strip whitespace.
            df["PART NUMBER"] = df["PART NUMBER"].astype(str).str.strip()

            # Ensure 'QTY.' is a numeric type. If a value can't be converted to a number, it becomes 0.
            df["QTY."] = pd.to_numeric(df["QTY."], errors="coerce").fillna(0)

            all_dataframes.append(df)

            # quang
            # Get the name of the file
            file_name = uploaded_file.name
            st.subheader(f"File: {file_name}")

            # Read the uploaded Excel file into a pandas DataFrame
            df = pd.read_excel(uploaded_file)

            # Display the DataFrame in the app
            st.dataframe(df)

        except Exception as e:
            error_files.append(uploaded_file.name)
            st.error(f"Error processing file '{uploaded_file.name}': {e}")

    if not all_dataframes:
        st.warning("No valid data could be processed from the uploaded files.")
    else:
        # --- Step 2: Combine All DataFrames ---
        # Concatenate all the individual dataframes into a single large one.
        combined_df = pd.concat(all_dataframes, ignore_index=True)

        # st.subheader("Combined Data (Before Merging)")
        # st.dataframe(combined_df)

        # --- Step 3: Group, Aggregate, and Sum ---
        # This is the core logic for merging based on your rule.
        # We group by both 'DESCRIPTION' and 'LENGTH'.
        # Then, we define how to aggregate the other columns.
        merged_df = (
            combined_df.groupby(["DESCRIPTION", "LENGTH"])
            .agg(
                {
                    "QTY.": "sum",  # Sum the quantities for each group.
                    "PART NUMBER": "first",  # Keep the first non-null part number for the group.
                }
            )
            .reset_index()
        )

        # --- Step 4: Final Formatting ---
        # Filter out rows where the description is blank, which can happen with empty rows in Excel.
        merged_df = merged_df[merged_df["DESCRIPTION"] != ""]

        # Reorder columns for a clean final output.
        final_columns = ["PART NUMBER", "DESCRIPTION", "LENGTH", "QTY."]
        merged_df = merged_df[final_columns]

        # Sort the results for consistency.
        merged_df = merged_df.sort_values(by=["DESCRIPTION", "LENGTH"]).reset_index(
            drop=True
        )

        st.subheader("Merged Bill of Materials")
        st.dataframe(merged_df)

        # --- Step 5: Prepare and Offer for Download ---
        excel_data = to_excel(merged_df)

        st.download_button(
            label="📥 Download Merged Excel File",
            data=excel_data,
            file_name="merged_bom.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

elif not uploaded_files:
    st.info("Please upload one or more Excel files to begin.")
