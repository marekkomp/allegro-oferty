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

            # Display the original data without any changes
            st.write("Original Data:")
            st.dataframe(df)

            # Automatically detect "Kategoria główna" and "Podkategoria" columns
            category_column = "Kategoria główna" if "Kategoria główna" in df.columns else None
            subcategory_column = "Podkategoria" if "Podkategoria" in df.columns else None

            if category_column and subcategory_column:
                # Extract unique categories and subcategories
                unique_categories = df[category_column].dropna().unique()
                unique_subcategories = df[subcategory_column].dropna().unique()

                # Create dropdowns for selection
                selected_category = st.selectbox("Select a main category to filter", unique_categories)
                selected_subcategory = st.selectbox("Select a subcategory to filter", unique_subcategories)

                # Filter data by selected category and subcategory
                filtered_df = df[(df[category_column] == selected_category) & (df[subcategory_column] == selected_subcategory)]

                # Display the number of positions in the filtered data
                st.write(f"Number of positions: {len(filtered_df)}")

                # Display filtered data
                st.write("Filtered Data by Selected Category and Subcategory:")
                st.dataframe(filtered_df)

                # Specify action: Remove or Append
                description_column = "Opis oferty" if "Opis oferty" in df.columns else None
                action = st.radio("Choose an action:", ["Remove Sentence", "Append Text"], index=0)

                if action == "Remove Sentence":
                    sentence_to_remove = st.text_input("Enter the sentence to remove from descriptions")

                    if description_column and sentence_to_remove:
                        # Remove the specified sentence
                        filtered_df[description_column] = filtered_df[description_column].str.replace(
                            sentence_to_remove, "", regex=False
                        )

                elif action == "Append Text":
                    sentence_to_find = st.text_input("Enter the sentence to search in descriptions")
                    sentence_to_append = st.text_input("Enter the text to append after the specified sentence")

                    if description_column and sentence_to_find:
                        # Filter rows containing the specified sentence
                        search_filtered_df = filtered_df[filtered_df[description_column].str.contains(sentence_to_find, na=False)]
                        st.write("Rows containing the searched sentence:")
                        st.dataframe(search_filtered_df)

                        # Append text after the specified sentence
                        if sentence_to_append:
                            search_filtered_df[description_column] = search_filtered_df[description_column].str.replace(
                                sentence_to_find, f"{sentence_to_find} {sentence_to_append}", regex=False
                            )

                        # Update filtered_df with changes from search_filtered_df
                        filtered_df.update(search_filtered_df)

                # Display modified data
                st.write("Modified Data:")
                st.dataframe(filtered_df)

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
            else:
                if not category_column:
                    st.error("The column 'Kategoria główna' was not found in the selected sheet.")
                if not subcategory_column:
                    st.error("The column 'Podkategoria' was not found in the selected sheet.")

    except Exception as e:
        st.error(f"An error occurred: {e}")
