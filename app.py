import streamlit as st
import pandas as pd
import os
import sys
import traceback
import logging
import datetime
import io
from io import StringIO

# 创建一个StringIO对象来捕获日志输出
console_log = StringIO()

# 配置根日志记录器
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

# 移除所有现有的处理程序
for handler in logging.root.handlers[:]:
    logging.root.removeHandler(handler)

# 添加StringIO处理程序用于网页显示
string_handler = logging.StreamHandler(console_log)
string_handler.setLevel(logging.INFO)
string_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
logging.getLogger().addHandler(string_handler)

# 同时输出到控制台
stream_handler = logging.StreamHandler(sys.stdout)
stream_handler.setLevel(logging.INFO)
stream_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
logging.getLogger().addHandler(stream_handler)

# 记录启动信息
logging.info("应用程序启动")

# Try importing the converter function
try:
    from excel_converter import convert_excel
    logging.info("成功导入excel_converter模块")
except ImportError as e:
    error_msg = f"导入excel_converter时出错: {e}"
    logging.error(error_msg)
    st.error(error_msg)
    st.stop()

# Translations for UI text in both languages
translations = {
    "en": {
        "page_title": "Excel Converter for Declaration List",
        "page_description": "This tool converts Excel files according to specific requirements for declaration purposes.",
        "sample_files": "Sample Files",
        "download_input_template": "Download Input Template",
        "download_reference_template": "Download Reference Template",
        "download_policy_template": "Download Policy Template",
        "input_template_help": "Download an Excel template for the input file with sample data and format",
        "reference_template_help": "Download an Excel template for the reference file with sample material codes",
        "policy_template_help": "Download an Excel template for the policy file with exchange and shipping rates",
        "upload_files": "Upload Files",
        "input_file": "Input File",
        "input_file_desc": "Excel file with export declaration data (green headers)",
        "upload_input": "Upload Input Excel File",
        "input_help": "This is your source data file with columns like NO., DESCRIPTION, Model NO., etc.",
        "reference_file": "Reference File",
        "reference_file_desc": "Excel file with declaration elements (for yellow headers)",
        "upload_reference": "Upload Reference Excel File",
        "reference_help": "This file should contain material codes and associated declaration information",
        "policy_file": "Policy File",
        "policy_file_desc": "Excel file with exchange rates and shipping information",
        "upload_policy": "Upload Policy Excel File",
        "policy_help": "This file should contain exchange rates and shipping rates",
        "policy_optional": "(Required)",
        "output_settings": "Output Settings",
        "output_filename": "Output Excel Filename",
        "output_help": "The name of the converted Excel file you'll download",
        "data_preview": "Data Preview",
        "input_preview": "Input Excel File Preview",
        "showing_rows": "Showing first 5 rows from sheet {} (total rows: {})",
        "columns": "Columns: {}",
        "reference_preview": "Reference Excel File Preview",
        "policy_preview": "Policy Excel File Preview",
        "could_not_preview": "Could not preview {} file: {}",
        "error_previewing": "Error previewing files: {}",
        "convert_button": "Convert Excel Files",
        "upload_both": "Please upload both input and reference Excel files before converting.",
        "upload_all": "Please upload input, reference and policy Excel files before converting.",
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
        "logs": "Application Logs",
        "view_logs": "View Application Logs",
        "clear_logs": "Clear Logs",
        "no_logs": "No logs available",
        "troubleshooting_tips": """
        - Make sure your input file has the expected column structure
        - Check that your reference file contains material codes
        - Try with different Excel files to see if the issue persists
        """,
        "policy_format_error": "The policy file format is incorrect. Please make sure you're uploading a valid policy file, not an output file.",
        "policy_format_guide": """
        A valid policy file should have the following structure:
        - Cell B4: Shipping fee (运费)
        - Cell B5: Exchange rate (汇率)
        - Cell B6: Price markup percentage (加价百分比)
        - Cell B7: Insurance coefficient 1 (保费系数1)
        - Cell B8: Insurance coefficient 2 (保费系数2)
        - Cell B16: Insurance amount (保险金额)
        """
    },
    "zh": {
        "page_title": "报关单Excel转换工具",
        "page_description": "此工具根据特定要求转换Excel文件，用于报关目的。",
        "sample_files": "示例文件",
        "download_input_template": "下载输入文件模板",
        "download_reference_template": "下载参考文件模板",
        "download_policy_template": "下载政策文件模板",
        "input_template_help": "下载输入文件的Excel模板，其中包含示例数据和格式",
        "reference_template_help": "下载参考文件的Excel模板，其中包含示例物料代码和商品编号",
        "policy_template_help": "下载政策文件的Excel模板，其中包含汇率和运输费率设置",
        "upload_files": "上传文件",
        "input_file": "出口发票文件",
        "input_file_desc": "出口报关文件",
        "upload_input": "上传输入Excel文件",
        "input_help": "这是您的源数据文件，包含NO.、DESCRIPTION、Model NO.等列",
        "reference_file": "申报要素文件",
        "reference_file_desc": "申报要素文件",
        "upload_reference": "上传参考Excel文件",
        "reference_help": "此文件应包含物料代码和相关的申报信息",
        "policy_file": "政策文件",
        "policy_file_desc": "包含汇率和运输信息的Excel文件",
        "upload_policy": "上传政策Excel文件",
        "policy_help": "此文件应包含汇率和运输费率",
        "policy_optional": "（必须）",
        "output_settings": "输出设置",
        "output_filename": "输出Excel文件名",
        "output_help": "您将下载的转换后Excel文件的名称",
        "data_preview": "数据预览",
        "input_preview": "输入Excel文件预览",
        "showing_rows": "显示第{}张表的前5行（总行数：{}）",
        "columns": "列：{}",
        "reference_preview": "参考Excel文件预览",
        "policy_preview": "政策Excel文件预览",
        "could_not_preview": "无法预览{}文件：{}",
        "error_previewing": "预览文件时出错：{}",
        "convert_button": "转换Excel文件",
        "upload_both": "请在转换前上传输入和参考Excel文件。",
        "upload_all": "请在转换前上传输入、参考和政策Excel文件。",
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
        "logs": "应用日志",
        "view_logs": "查看应用日志",
        "clear_logs": "清除日志",
        "no_logs": "没有可用的日志",
        "troubleshooting_tips": """
        - 确保您的输入文件具有预期的列结构
        - 检查您的参考文件是否包含物料代码
        - 尝试使用不同的Excel文件，看问题是否仍然存在
        """,
        "policy_format_error": "政策文件格式不正确。请确保您上传的是有效的政策文件，而不是输出文件。",
        "policy_format_guide": """
        有效的政策文件应具有以下结构：
        - 单元格 B4：运费
        - 单元格 B5：汇率
        - 单元格 B6：加价百分比
        - 单元格 B7：保费系数1
        - 单元格 B8：保费系数2
        - 单元格 B16：保险金额
        """
    }
}

def main():
    # 记录主函数调用
    logging.info("主函数开始执行")

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
        logging.info(f"已选择语言: {lang}")

        # Add some information in the sidebar
        st.divider()
        if lang == "zh":
            st.info("此应用程序将Excel文件转换为符合报关要求的格式")
            st.markdown("**使用说明**")
            st.markdown("1. 上传输入Excel文件（带有绿色表头）")
            st.markdown("2. 上传参考Excel文件（用于物料代码匹配）")
            st.markdown("3. 上传政策Excel文件（必须，用于汇率和运输信息）")
            st.markdown("4. 指定输出文件名")
            st.markdown('5. 点击"转换Excel文件"按钮')
            st.markdown("6. 下载转换后的文件")
        else:
            st.info("This app converts Excel files to meet declaration requirements")
            st.markdown("**Instructions**")
            st.markdown("1. Upload the input Excel file (with green headers)")
            st.markdown("2. Upload the reference Excel file (for material code matching)")
            st.markdown("3. Upload the policy Excel file (required, for exchange and shipping rates)")
            st.markdown("4. Specify the output filename")
            st.markdown("5. Click the 'Convert Excel Files' button")
            st.markdown("6. Download the converted file")

    # Get text for the selected language
    t = translations[lang]

    # Main page content
    st.title(t["page_title"])
    st.write(t["page_description"])

    # 移除所有模板下载功能和标题

    # File uploaders
    st.header(t["upload_files"])

    col1, col2 = st.columns(2)

    with col1:
        st.subheader(t["input_file"])
        st.write(t["input_file_desc"])
        input_file = st.file_uploader(t["upload_input"], type=["xlsx", "xls"], help=t["input_help"])
        if input_file is not None:
            logging.info(f"已上传输入文件: {input_file.name}")

    with col2:
        st.subheader(t["reference_file"])
        st.write(t["reference_file_desc"])
        reference_file = st.file_uploader(t["upload_reference"], type=["xlsx", "xls"], help=t["reference_help"])
        if reference_file is not None:
            logging.info(f"已上传参考文件: {reference_file.name}")

    # Policy file uploader (required)
    st.subheader(f"{t['policy_file']}")
    st.write(t["policy_file_desc"])
    policy_file = st.file_uploader(t["upload_policy"], type=["xlsx", "xls"], help=t["policy_help"])
    if policy_file is not None:
        logging.info(f"已上传政策文件: {policy_file.name}")

    # Output file name
    st.header(t["output_settings"])
    output_filename = st.text_input(t["output_filename"], "报关单.xlsx", help=t["output_help"])
    logging.info(f"输出文件名设置为: {output_filename}")

    # Preview section
    if input_file is not None and reference_file is not None:
        try:
            logging.info("开始预览数据")
            st.header(t["data_preview"])

            # Preview input file - with expanded error handling
            # Preview input file
            st.subheader(t["input_preview"])
            try:
                # Try to detect sheet count and use appropriate sheet
                xl = pd.ExcelFile(input_file)
                sheet_count = len(xl.sheet_names)
                sheet_to_read = 1 if sheet_count >= 2 else 0

                input_df = pd.read_excel(input_file, skiprows=9, sheet_name=sheet_to_read)
                if len(input_df) > 0:
                    input_df = input_df.drop(index=0).reset_index(drop=True)
                    # 将所有列转换为字符串类型
                    input_df = input_df.astype(str)

                st.dataframe(input_df.head())
                st.caption(t["showing_rows"].format(sheet_to_read+1, len(input_df)))
                st.text(t["columns"].format(', '.join(input_df.columns.tolist())))
                logging.info(f"输入文件预览成功: {sheet_count}个工作表, 已读取第{sheet_to_read+1}个, {len(input_df)}行数据")
            except Exception as e:
                error_msg = f"无法预览输入文件: {str(e)}"
                logging.error(error_msg)
                st.warning(t["could_not_preview"].format(t["input_file"].lower(), str(e)))

            # Preview reference file
            st.subheader(t["reference_preview"])
            try:
                reference_df = pd.read_excel(reference_file)
                # 将参考文件的所有列也转换为字符串类型
                reference_df = reference_df.astype(str)

                st.dataframe(reference_df.head())
                st.caption(t["showing_rows"].format(1, len(reference_df)))
                st.text(t["columns"].format(', '.join(reference_df.columns.tolist())))
                logging.info(f"参考文件预览成功: {len(reference_df)}行数据")
            except Exception as e:
                error_msg = f"无法预览参考文件: {str(e)}"
                logging.error(error_msg)
                st.warning(t["could_not_preview"].format(t["reference_file"].lower(), str(e)))

            # Preview policy file (if uploaded)
            if policy_file is not None:
                st.subheader(t["policy_preview"])
                try:
                    policy_df = pd.read_excel(policy_file)
                    policy_df = policy_df.astype(str)

                    st.dataframe(policy_df.head())
                    st.caption(t["showing_rows"].format(1, len(policy_df)))
                    st.text(t["columns"].format(', '.join(policy_df.columns.tolist())))
                    logging.info(f"政策文件预览成功: {len(policy_df)}行数据")
                except Exception as e:
                    error_msg = f"无法预览政策文件: {str(e)}"
                    logging.error(error_msg)
                    st.warning(t["could_not_preview"].format(t["policy_file"].lower(), str(e)))
        except Exception as e:
            error_msg = f"预览文件时出错: {str(e)}"
            logging.error(error_msg)
            st.error(t["error_previewing"].format(str(e)))

    # Convert button
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        convert_button = st.button(t["convert_button"], type="primary", use_container_width=True)

    if convert_button:
        logging.info("点击了转换按钮")
        if input_file is None or reference_file is None or policy_file is None:
            error_msg = "请在转换前上传输入、参考和政策Excel文件"
            logging.warning(error_msg)
            st.error(t["upload_all"])
        else:
            try:
                # Create a progress placeholder
                progress_container = st.empty()
                progress_container.info(t["starting_conversion"])
                logging.info("开始转换过程")

                # Save uploaded files temporarily
                progress_container.info(t["saving_temp"])
                logging.info("临时保存上传的文件")
                with open("temp_input.xlsx", "wb") as f:
                    f.write(input_file.getbuffer())

                with open("temp_reference.xlsx", "wb") as f:
                    f.write(reference_file.getbuffer())

                # Save policy file with a unique name to avoid conflicts
                import uuid
                policy_filename = f"temp_policy_{uuid.uuid4().hex[:8]}.xlsx"
                with open(policy_filename, "wb") as f:
                    f.write(policy_file.getbuffer())
                policy_path = policy_filename
                logging.info(f"已保存政策文件: {policy_path}")

                # Process the conversion
                progress_container.info(t["converting"])
                logging.info(f"开始调用convert_excel函数，参数：input={input_file.name}, reference={reference_file.name}, output={output_filename}, policy={policy_path}")

                try:
                    result = convert_excel("temp_input.xlsx", "temp_reference.xlsx",policy_path, output_filename )

                    # Check if conversion was successful
                    if result is None:
                        error_msg = "转换失败，convert_excel返回None"
                        logging.error(error_msg)
                        st.error(t["conversion_failed"])
                        st.stop()
                except ValueError as e:
                    # 捕获 policy 文件验证错误
                    error_msg = f"Policy 文件验证失败: {str(e)}"
                    logging.error(error_msg)
                    st.error(error_msg)

                    # 提供更具体的指导
                    st.warning(t["policy_format_error"])
                    st.info(t["policy_format_guide"])
                    st.stop()

                # Clean up temp files
                progress_container.info(t["cleaning_up"])
                logging.info("清理临时文件")
                import time

                def safe_remove(file_path, max_retries=3, delay=1):
                    for i in range(max_retries):
                        try:
                            if os.path.exists(file_path):
                                os.close(os.open(file_path, os.O_RDONLY))  # 确保文件句柄被关闭
                                os.remove(file_path)
                                logging.info(f"已删除临时文件: {file_path}")
                                return True
                        except Exception as e:
                            error_msg = f"无法删除文件 {file_path}: {str(e)}"
                            logging.warning(error_msg)
                            if i < max_retries - 1:
                                time.sleep(delay)
                                continue
                            else:
                                print(error_msg)
                                return False
                    return False

                # 尝试删除临时文件
                safe_remove("temp_input.xlsx")
                safe_remove("temp_reference.xlsx")
                safe_remove(policy_path)  # 使用实际的 policy 文件路径

                progress_container.success(t["success"])
                logging.info("转换成功完成")

                # Provide download link
                if os.path.exists(output_filename):
                    logging.info(f"输出文件已创建: {output_filename}")
                    with open(output_filename, "rb") as file:
                        st.download_button(
                            label=t["download_button"],
                            data=file,
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                else:
                    error_msg = f"输出文件 '{output_filename}' 未创建，转换可能已失败"
                    logging.error(error_msg)
                    st.error(t["output_not_created"].format(output_filename))
            except Exception as e:
                error_msg = f"转换过程中发生错误: {str(e)}"
                logging.error(error_msg)
                logging.error(traceback.format_exc())
                st.error(t["error_occurred"].format(str(e)))
                with st.expander(t["view_details"]):
                    st.code(traceback.format_exc())

                st.info(t["troubleshooting"])
                st.markdown(t["troubleshooting_tips"])

    # 添加日志查看器
    st.divider()
    st.header(t["logs"])

    log_cols = st.columns([1, 1, 3])

    with log_cols[0]:
        if st.button(t["view_logs"], use_container_width=True):
            logging.info("查看日志按钮被点击")

    with log_cols[1]:
        if st.button(t["clear_logs"], use_container_width=True):
            try:
                # 清除内存中的日志
                console_log.truncate(0)
                console_log.seek(0)

                logging.info("日志已清除")
                st.success("日志已成功清除")

            except Exception as e:
                error_msg = f"清除日志时出错: {str(e)}"
                st.error(error_msg)
                logging.error(error_msg)
                logging.error(traceback.format_exc())

    # 显示日志内容
    try:
        # 获取内存中的日志内容
        log_content = console_log.getvalue()

        if log_content:
            with st.expander("日志内容", expanded=True):
                st.code(log_content)
        else:
            st.info(t["no_logs"])
    except Exception as e:
        st.error(f"读取日志时出错: {str(e)}")
        logging.error(f"读取日志时出错: {traceback.format_exc()}")

if __name__ == "__main__":
    main()