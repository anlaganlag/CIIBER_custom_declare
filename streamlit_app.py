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

# Translations for UI text in both languages
translations = {
    "en": {
        "page_title": "Excel Converter for Declaration List",
        "page_description": "This tool converts Excel files according to specific requirements for declaration purposes.",
        "upload_files": "Upload Files",
        "input_file": "Input File",
        "input_file_desc": "Excel file with your source data (green headers)",
        "upload_input": "Upload Input Excel File",
        "input_help": "This is your source data file with columns like NO., DESCRIPTION, Model NO., etc.",
        "reference_file": "Reference File",
        "reference_file_desc": "Excel file with material codes (for yellow headers)",
        "upload_reference": "Upload Reference Excel File",
        "reference_help": "This file should contain material codes and associated declaration information",
        "output_settings": "Output Settings",
        "output_filename": "Output Excel Filename",
        "output_help": "The name of the converted Excel file you'll download",
        "data_preview": "Data Preview",
        "input_preview": "Input Excel File Preview",
        "showing_rows": "Showing first 5 rows from sheet {} (total rows: {})",
        "columns": "Columns: {}",
        "reference_preview": "Reference Excel File Preview",
        "could_not_preview": "Could not preview {} file: {}",
        "error_previewing": "Error previewing files: {}",
        "convert_button": "Convert Excel Files",
        "upload_both": "Please upload both input and reference Excel files before converting.",
        "starting_conversion": "Starting conversion process...",
        "saving_temp": "Saving uploaded files temporarily...",
        "converting": "Converting files... This may take a moment.",
        "conversion_failed": "Conversion failed. Please check the console output for details.",
        "cleaning_up": "Cleaning up temporary files...",
        "success": "Conversion completed successfully!",
        "download_button": "Download Converted Excel",
        "output_not_created": "Output file '{}' was not created. Conversion may have failed.",
        "error_occurred": "An error occurred during conversion: {}",
        "view_details": "View detailed error information",
        "troubleshooting": "Troubleshooting tips:",
        "troubleshooting_tips": """
        - Make sure your input file has the expected column structure
        - Check that your reference file contains material codes
        - Try with different Excel files to see if the issue persists
        """
    },
    "zh": {
        "page_title": "报关单Excel转换工具",
        "page_description": "此工具根据特定要求转换Excel文件，用于报关目的。",
        "upload_files": "上传文件",
        "input_file": "输入文件",
        "input_file_desc": "带有源数据的Excel文件（绿色表头）",
        "upload_input": "上传输入Excel文件",
        "input_help": "这是您的源数据文件，包含NO.、DESCRIPTION、Model NO.等列",
        "reference_file": "参考文件",
        "reference_file_desc": "带有物料代码的Excel文件（用于黄色表头）",
        "upload_reference": "上传参考Excel文件",
        "reference_help": "此文件应包含物料代码和相关的申报信息",
        "output_settings": "输出设置",
        "output_filename": "输出Excel文件名",
        "output_help": "您将下载的转换后Excel文件的名称",
        "data_preview": "数据预览",
        "input_preview": "输入Excel文件预览",
        "showing_rows": "显示第{}张表的前5行（总行数：{}）",
        "columns": "列：{}",
        "reference_preview": "参考Excel文件预览",
        "could_not_preview": "无法预览{}文件：{}",
        "error_previewing": "预览文件时出错：{}",
        "convert_button": "转换Excel文件",
        "upload_both": "请在转换前上传输入和参考Excel文件。",
        "starting_conversion": "开始转换过程...",
        "saving_temp": "临时保存上传的文件...",
        "converting": "正在转换文件...这可能需要一点时间。",
        "conversion_failed": "转换失败。请查看控制台输出了解详情。",
        "cleaning_up": "清理临时文件...",
        "success": "转换成功完成！",
        "download_button": "下载转换后的Excel",
        "output_not_created": "输出文件'{}'未创建。转换可能已失败。",
        "error_occurred": "转换过程中发生错误：{}",
        "view_details": "查看详细错误信息",
        "troubleshooting": "故障排除提示：",
        "troubleshooting_tips": """
        - 确保您的输入文件具有预期的列结构
        - 检查您的参考文件是否包含物料代码
        - 尝试使用不同的Excel文件，看问题是否仍然存在
        """
    }
}

def main():
    # Set page configuration
    st.set_page_config(
        page_title="Excel Converter",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # Language selection in sidebar (default to Chinese)
    with st.sidebar:
        st.title("🌐 语言 / Language")
        lang = st.selectbox(
            "选择语言 / Select Language",
            options=["zh", "en"],
            format_func=lambda x: "中文" if x == "zh" else "English",
            index=0  # Default to Chinese (index 0)
        )
        
        # Add some information in the sidebar
        st.divider()
        if lang == "zh":
            st.info("此应用程序将Excel文件转换为符合报关要求的格式")
            st.markdown("**使用说明**")
            st.markdown("1. 上传输入Excel文件（带有绿色表头）")
            st.markdown("2. 上传参考Excel文件（用于物料代码匹配）")
            st.markdown("3. 指定输出文件名")
            st.markdown('4. 点击"转换Excel文件"按钮')
            st.markdown("5. 下载转换后的文件")
        else:
            st.info("This app converts Excel files to meet declaration requirements")
            st.markdown("**Instructions**")
            st.markdown("1. Upload the input Excel file (with green headers)")
            st.markdown("2. Upload the reference Excel file (for material code matching)")
            st.markdown("3. Specify the output filename")
            st.markdown("4. Click the 'Convert Excel Files' button")
            st.markdown("5. Download the converted file")
    
    # Get text for the selected language
    t = translations[lang]
    
    # Main page content
    st.title(t["page_title"])
    st.write(t["page_description"])
    
    # File uploaders
    st.header(t["upload_files"])
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader(t["input_file"])
        st.write(t["input_file_desc"])
        input_file = st.file_uploader(t["upload_input"], type=["xlsx", "xls"], help=t["input_help"])
    
    with col2:
        st.subheader(t["reference_file"])
        st.write(t["reference_file_desc"])
        reference_file = st.file_uploader(t["upload_reference"], type=["xlsx", "xls"], help=t["reference_help"])
    
    # Output file name
    st.header(t["output_settings"])
    output_filename = st.text_input(t["output_filename"], "merged.xlsx", help=t["output_help"])
    
    # Preview section
    if input_file is not None and reference_file is not None:
        try:
            st.header(t["data_preview"])
            
            # Preview input file - with expanded error handling
            st.subheader(t["input_preview"])
            try:
                # Try to detect sheet count and use appropriate sheet
                xl = pd.ExcelFile(input_file)
                sheet_count = len(xl.sheet_names)
                sheet_to_read = 1 if sheet_count >= 2 else 0
                
                input_df = pd.read_excel(input_file, skiprows=9, sheet_name=sheet_to_read)
                if len(input_df) > 0:
                    input_df = input_df.drop(index=0).reset_index(drop=True)
                
                st.dataframe(input_df.head())
                st.caption(t["showing_rows"].format(sheet_to_read+1, len(input_df)))
                st.text(t["columns"].format(', '.join(input_df.columns.tolist())))
            except Exception as e:
                st.warning(t["could_not_preview"].format(t["input_file"].lower(), str(e)))
            
            # Preview reference file
            st.subheader(t["reference_preview"])
            try:
                reference_df = pd.read_excel(reference_file)
                st.dataframe(reference_df.head())
                st.caption(t["showing_rows"].format(1, len(reference_df)))
                st.text(t["columns"].format(', '.join(reference_df.columns.tolist())))
            except Exception as e:
                st.warning(t["could_not_preview"].format(t["reference_file"].lower(), str(e)))
        except Exception as e:
            st.error(t["error_previewing"].format(str(e)))
    
    # Convert button
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        convert_button = st.button(t["convert_button"], type="primary", use_container_width=True)
    
    if convert_button:
        if input_file is None or reference_file is None:
            st.error(t["upload_both"])
        else:
            try:
                # Create a progress placeholder
                progress_container = st.empty()
                progress_container.info(t["starting_conversion"])
                
                # Save uploaded files temporarily
                progress_container.info(t["saving_temp"])
                with open("temp_input.xlsx", "wb") as f:
                    f.write(input_file.getbuffer())
                
                with open("temp_reference.xlsx", "wb") as f:
                    f.write(reference_file.getbuffer())
                
                # Process the conversion
                progress_container.info(t["converting"])
                result = convert_excel("temp_input.xlsx", "temp_reference.xlsx", output_filename)
                
                # Check if conversion was successful
                if result is None:
                    st.error(t["conversion_failed"])
                    st.stop()
                
                # Clean up temp files
                progress_container.info(t["cleaning_up"])
                if os.path.exists("temp_input.xlsx"):
                    os.remove("temp_input.xlsx")
                if os.path.exists("temp_reference.xlsx"):
                    os.remove("temp_reference.xlsx")
                
                progress_container.success(t["success"])
                
                # Provide download link
                if os.path.exists(output_filename):
                    with open(output_filename, "rb") as file:
                        st.download_button(
                            label=t["download_button"],
                            data=file,
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                else:
                    st.error(t["output_not_created"].format(output_filename))
            except Exception as e:
                st.error(t["error_occurred"].format(str(e)))
                with st.expander(t["view_details"]):
                    st.code(traceback.format_exc())
                
                st.info(t["troubleshooting"])
                st.markdown(t["troubleshooting_tips"])

if __name__ == "__main__":
    main()