import streamlit as st
import pandas as pd
from datetime import date
import openpyxl

def main():
    st.set_page_config(layout="wide")
    st.title("Filter Invoices by pre-defined date range and product list (WCN)")
    st.subheader('File must be in same format as _Invoice Report All Jan23-Mar24.xlsx_')

    # File upload section
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

    if uploaded_file is not None:
        st.write('Preview of DataFrame')
        try:
            # Read the Excel file
            df = pd.read_excel(uploaded_file, skiprows=4, engine='openpyxl')
            df = df.iloc[:-1]
            products = ['KH60','KH61','KL33','KK91','KK37','KK95','KK97','KK98','KL28','KL29','KL30','KL31','KL32','KL34','KK38','KK36']


	    
            # filter df by product list
            fbp_df = df[df['Product/Service'].isin(products)]
            
            st.dataframe(fbp_df)
	    
            today = date.today()
            
            # Convert to new excel file
            output_text = fbp_df.to_excel(f'Invoices filtered by product list Jan23-{today}.xlsx', index=None)

            # Create a download link
            st.markdown(get_download_link(output_text), unsafe_allow_html=True)
        except Exception as e:
            st.write('An Error Occured.')
            st.error(f"Error reading the file: {e}")

def get_download_link(text):
    today = date.today()
    # Generate a download link for the text file
    href = f'<a href="data:text/plain;charset=utf-8,{text}" download="WCN_Invoices_{today}.txt">Download filtered invoices as Excel file.</a>'
    return href

if __name__ == "__main__":
    main()
