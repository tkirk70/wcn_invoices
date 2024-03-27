import streamlit as st
import pandas as pd
from datetime import date
import openpyxl
from io import BytesIO

def main():
    st.set_page_config(layout="wide")
    st.title("Filter Invoices by pre-defined date range and product list (WCN)")
    st.subheader('File must be in the same format as _Invoice Report All Jan23-Mar24.xlsx_')

    # File upload section
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

    if uploaded_file is not None:
        st.write('Preview of DataFrame')
        try:
            # Read the Excel file
            df = pd.read_excel(uploaded_file, skiprows=4, engine='openpyxl')
            df = df.iloc[:-1]
            products = ['KH60', 'KH61', 'KL33', 'KK91', 'KK37', 'KK95', 'KK97', 'KK98', 'KL28', 'KL29', 'KL30', 'KL31', 'KL32', 'KL34', 'KK38', 'KK36']

            # Filter df by product list
            fbp_df = df[df['Product/Service'].isin(products)]
            
            today = date.today()

            st.dataframe(fbp_df)
            file_name = f'WCN_Invoices_Jan23-{today}.xlsx'
            fbp_df.to_excel(file_name, index=None)
            
            
            with open(file_name, "rb") as template_file:
                template_byte = template_file.read()

            st.download_button(label="Click to Download Filtered Invoices File",
                        data=template_byte,
                        file_name=file_name,
                        mime='application/octet-stream')

        except Exception as e:
            st.write('An Error Occurred.')
            st.error(f"Error reading the file: {e}")


if __name__ == "__main__":
    main()
