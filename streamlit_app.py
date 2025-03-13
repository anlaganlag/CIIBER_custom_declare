import streamlit as st
import pandas as pd
import os
import sys
import traceback

# Try importing the converter function
try:
    from excel_converter import convert_excel
except ImportError as e:
    st.error(f"Error importing excel_converter: {e}")
    st.stop()

def main():
    st.title("Excel Converter")
    st.write("This tool converts Excel files according to specific requirements.")
    
    # File uploaders
    st.header("Input Files")
    input_file = st.file_uploader("Upload Input Excel File (with green headers)", type=["xlsx", "xls"])
    reference_file = st.file_uploader("Upload Reference Excel File (with material codes)", type=["xlsx", "xls"])
    
    # Output file name
    st.header("Output")
    output_filename = st.text_input("Output Excel Filename", "converted_output.xlsx")
    
    # Preview section
    if input_file is not None and reference_file is not None:
        st.header("Preview Input Data")
        
        # Preview input file
        st.subheader("Input Excel File")
        input_df = pd.read_excel(input_file)
        st.dataframe(input_df.head())
        
        # Preview reference file
        st.subheader("Reference Excel File")
        reference_df = pd.read_excel(reference_file)
        st.dataframe(reference_df.head())
    
    # Convert button
    if st.button("Convert Excel Files"):
        if input_file is None or reference_file is None:
            st.error("Please upload both input and reference Excel files.")
        else:
            try:
                # Save uploaded files temporarily
                with open("temp_input.xlsx", "wb") as f:
                    f.write(input_file.getbuffer())
                
                with open("temp_reference.xlsx", "wb") as f:
                    f.write(reference_file.getbuffer())
                
                # Process the conversion
                with st.spinner("Converting..."):
                    convert_excel("temp_input.xlsx", "temp_reference.xlsx", output_filename)
                
                # Clean up temp files
                if os.path.exists("temp_input.xlsx"):
                    os.remove("temp_input.xlsx")
                if os.path.exists("temp_reference.xlsx"):
                    os.remove("temp_reference.xlsx")
                
                # Success message and download button
                st.success(f"Conversion completed successfully!")
                
                # Provide download link
                with open(output_filename, "rb") as file:
                    st.download_button(
                        label="Download Converted Excel",
                        data=file,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"An error occurred during conversion: {str(e)}")
                st.error(traceback.format_exc())

if __name__ == "__main__":
    main()