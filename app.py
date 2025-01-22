import streamlit as st
import pandas as pd
import os

# Streamlit application
st.title("Allegro Offers Management")

# File uploader
uploaded_file = st.file_uploader("Upload XLSM file", type="xlsm")

if uploaded_file:
    try:
        # Load the XLSM file
        xlsm_data = pd.ExcelFile(uploaded_file)

        # Display sheet names
        st.write("Sheets found in the file:", xlsm_data.sheet_names)

        # Select sheet
        sheet_name = st.selectbox("Select the sheet to process", xlsm_data.sheet_names)

        if sheet_name:
            # Load selected sheet, specifying the correct header row
            df = pd.read_excel(xlsm_data, sheet_name=sheet_name, header=3)

            # Display a preview of the data
            st.write("Preview of the selected sheet:", df.head())

            # Column selection for filtering and modification
            category_column = st.selectbox("Select the main category column", df.columns)
            subcategory_column = st.selectbox("Select the subcategory column", df.columns)
            description_column = st.selectbox("Select the description column", df.columns)

            # Specify category to filter
            selected_category = st.text_input("Enter the category to filter")

            # Specify sentence to remove from descriptions
            sentence_to_remove = st.text_input("Enter the sentence to remove from descriptions")

            if st.button("Apply Changes"):
                # Filter by category
                if selected_category:
                    filtered_df = df[df[category_column].str.contains(selected_category, na=False)]
                    st.write("Filtered data:", filtered_df.head())
                else:
                    filtered_df = df

                # Remove the sentence from descriptions
                if sentence_to_remove and description_column in filtered_df.columns:
                    filtered_df[description_column] = filtered_df[description_column].str.replace(sentence_to_remove, "", regex=False)

                # Display modified data
                st.write("Modified data:", filtered_df.head())

                # Allow download of modified data
                output_file = "modified_file.xlsx"
                filtered_df.to_excel(output_file, index=False, engine='openpyxl')
                with open(output_file, "rb") as f:
                    st.download_button(
                        label="Download Modified File",
                        data=f,
                        file_name="modified_file.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                # Clean up temporary file
                os.remove(output_file)

    except Exception as e:
        st.error(f"An error occurred: {e}")
