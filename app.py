import streamlit as st
import pandas as pd
from io import BytesIO

# Streamlit App
st.title("Excel Tag Transformation")

# File uploader
uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx'])
if uploaded_file:
    # Display the file details
    st.write("File uploaded:", uploaded_file.name)

    # Load Excel file
    try:
        excel_data = pd.ExcelFile(uploaded_file)
        sheet_name = st.selectbox("Select the sheet to process", excel_data.sheet_names)
        data = excel_data.parse(sheet_name)

        # Display the preview of the data
        st.write("Preview of the data:")
        st.dataframe(data.head())

        # Inputs for header row and column ranges
        header_row = st.number_input("Enter the header row number (0-indexed)", min_value=0, max_value=len(data)-1, step=1)
        start_col = st.text_input("Enter the column letter where sizes start (e.g., 'C')")
        end_col = st.text_input("Enter the column letter where sizes end (e.g., 'Z')")

        if st.button("Transform Data"):
            # Adjust header and subset data
            data.columns = data.iloc[header_row]
            data = data[header_row + 1:].reset_index(drop=True)

            # Ensure unique column names by appending suffixes to duplicates
            data.columns = pd.Series(data.columns).apply(lambda x: f"{x}_" if pd.Series(data.columns).duplicated().any() else x)

            # Validate column range
            start_idx = ord(start_col.upper()) - 65
            end_idx = ord(end_col.upper()) - 65 + 1

            if 0 <= start_idx < len(data.columns) and 0 <= end_idx <= len(data.columns):
                # Correctly extract the relevant columns for melting
                id_vars = list(data.columns[:start_idx])
                value_vars = list(data.columns[start_idx:end_idx])

                # Transform the data
                transformed_data = data.melt(
                    id_vars=id_vars,
                    value_vars=value_vars,
                    var_name='Size',
                    value_name='Quantity'
                ).dropna()

                # Ensure Size column reflects the correct header values (using original header row)
                transformed_data['Size'] = transformed_data['Size'].astype(str)

                # Rename columns properly
                transformed_data.rename(columns={id_vars[0]: "Index", id_vars[1]: "Total"}, inplace=True)

                # Display the transformed data
                st.write("Transformed Data:")
                st.dataframe(transformed_data.head())

                # Create downloadable Excel
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    transformed_data.to_excel(writer, index=False, sheet_name='Transformed')
                output.seek(0)

                st.download_button(
                    label="Download Transformed Excel",
                    data=output,
                    file_name="Transformed_Stock_Data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("Invalid column letters for size range. Please check your input.")
    except Exception as e:
        st.error(f"An error occurred: {e}")
