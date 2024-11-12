import streamlit as st
import pandas as pd

st.title("Excel File Processor")

uploaded_file1 = st.file_uploader("Upload the first Excel file", type="xlsx")
uploaded_file2 = st.file_uploader("Upload the second Excel file", type="xlsx")

if uploaded_file1 and uploaded_file2:
    # Read the files into dataframes
    df1 = pd.read_excel(uploaded_file1)
    df2 = pd.read_excel(uploaded_file2)

    # Perform processing (replace this with your own logic)
    result = df1.merge(df2, on='YourCommonColumn')  # Example operation

    # Output to Excel file
    result.to_excel("output.xlsx", index=False)

    # Download link for output file
    st.download_button(label="Download Processed File", data=result.to_excel(index=False, engine='openpyxl'), file_name="output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
