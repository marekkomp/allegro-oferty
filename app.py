import streamlit as st
import pandas as pd
import os

# Streamlit application
st.title("Allegro Description Manager")

# File uploader
uploaded_file = st.file_uploader("Upload XLSM file", type="xlsm")

if uploaded_file:
    try:
        xlsm_data = pd.ExcelFile(uploaded_file)
        st.write("Sheets found in the file:", xlsm_data.sheet_names)
        sheet_name = st.selectbox("Select the sheet to process", xlsm_data.sheet_names)

        if sheet_name:
            df = pd.read_excel(xlsm_data, sheet_name=sheet_name, header=3)
            st.write("Original Data:")
            st.dataframe(df)

            category_column = "Kategoria główna" if "Kategoria główna" in df.columns else None
            subcategory_column = "Podkategoria" if "Podkategoria" in df.columns else None

            if category_column and subcategory_column:
                selected_category = st.selectbox("Select a main category to filter", df[category_column].dropna().unique())
                selected_subcategory = st.selectbox("Select a subcategory to filter", df[subcategory_column].dropna().unique())
                filtered_df = df[(df[category_column] == selected_category) & (df[subcategory_column] == selected_subcategory)]

                st.write(f"Number of positions: {len(filtered_df)}")
                st.dataframe(filtered_df)

                description_column = "Opis oferty" if "Opis oferty" in df.columns else None
                action = st.radio("Choose an action:", ["Remove Sentence", "Append Text", "Find Missing Sentence"], index=0)

                if action == "Remove Sentence":
                    sentence_to_remove = st.text_input("Enter the sentence to remove from descriptions")
                    if description_column and sentence_to_remove:
                        filtered_df[description_column] = filtered_df[description_column].str.replace(sentence_to_remove, "", regex=False)
                        st.write("Updated Data:")
                        st.dataframe(filtered_df)
                
                elif action == "Append Text":
                    sentence_to_find = st.text_input("Enter the sentence to search in descriptions")
                    sentence_to_append = st.text_input("Enter the text to append after the specified sentence")
                    if description_column and sentence_to_find and sentence_to_append:
                        filtered_df[description_column] = filtered_df[description_column].str.replace(
                            sentence_to_find, f"{sentence_to_find} {sentence_to_append}", regex=False
                        )
                        st.write("Updated Data:")
                        st.dataframe(filtered_df)
                
                elif action == "Find Missing Sentence":
                    sentence_to_find_missing = st.text_input("Enter the sentence to check for missing occurrences")
                    if description_column and sentence_to_find_missing:
                        missing_rows = filtered_df[~filtered_df[description_column].str.contains(sentence_to_find_missing, na=False)]
                        st.write("Rows without the specified sentence:")
                        st.dataframe(missing_rows)
                        if not missing_rows.empty:
                            missing_output_file = "missing_rows.xlsx"
                            missing_rows.to_excel(missing_output_file, index=False, engine='openpyxl')
                            with open(missing_output_file, "rb") as f:
                                st.download_button(
                                    label="Download Rows Without Sentence",
                                    data=f,
                                    file_name="missing_rows.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                            os.remove(missing_output_file)
            else:
                st.error("Required columns not found in the sheet.")
    except Exception as e:
        st.error(f"An error occurred: {e}")
