import streamlit as st
import pandas as pd
import io

st.title("Excel Config Cleaner")

uploaded_file = st.file_uploader("Drop your Excel file here (.xls)", type=["xls"])

if uploaded_file:
    # Step 1: Load Excel with header at row 43
    df = pd.read_excel(uploaded_file, header=43)
    df.columns = df.columns.str.strip()

    # Step 2: Modify DataFrame
    if 'Extended List Price' in df.columns:
        df['RRP'] = df['Extended List Price'] - 0
        df['Extended List Price'] = df['Extended List Price'] - 0
    if 'Quantity' in df.columns:
        df['Quantity'] = df['Quantity'] - 0
    if 'Effective Discount %' in df.columns:
        df['Customer Discount'] = df['Effective Discount %'] - 1

    # Step 3: Remove columns 4 to 42, preserving key ones
    protected = ['Extended List Price', 'Quantity', 'Effective Discount %']
    cols_to_remove = [col for i, col in enumerate(df.columns) if 4 <= i <= 41 and col not in protected]
    df.drop(columns=cols_to_remove, inplace=True)

    # Step 4: Export to Excel in memory
    output = io.BytesIO()
    df.to_excel(output, index=False)
    st.download_button("Download cleaned file", output.getvalue(), file_name="modified_output.xlsx")