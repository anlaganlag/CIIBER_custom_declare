# Excel Converter for Declaration List

This is a Streamlit application for converting Excel files to meet declaration requirements. The application processes input Excel files, reference Excel files, and policy files to generate output Excel files in the required format for declaration purposes.

## One-Click Installation and Startup

### Windows Users

1. Ensure Python 3.8 or higher is installed on your computer
2. Download all files of this project
3. Double-click to run `setup_windows.bat`
4. The script will automatically create a virtual environment, install dependencies, and launch the application

### Mac/Linux Users

1. Ensure Python 3.8 or higher is installed on your computer
2. Download all files of this project
3. Open a terminal and navigate to the project folder
4. Run the following command to make the script executable:
   ```
   chmod +x setup_mac.sh
   ```
5. Execute the script:
   ```
   ./setup_mac.sh
   ```
6. The script will automatically create a virtual environment, install dependencies, and launch the application

## Manual Installation

If the one-click script doesn't work, you can manually execute the following steps:

### Windows

```
python -m venv venv
venv\Scripts\activate
pip install --upgrade pip
pip install -r requirements.txt
streamlit run streamlit_app.py
```

### Mac/Linux

```
python3 -m venv venv
source venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Application Usage Instructions

1. After launching the application, a window will open in your browser
2. You can download sample file templates at the top of the page:
   - Input file template (input_template.xlsx)
   - Reference file template (reference_template.xlsx)
   - Policy file template (policy_template.xlsx)
3. Upload your files:
   - Input Excel file: Contains source data with green headers
   - Reference Excel file: Contains material codes and declaration information for yellow headers
   - Policy Excel file: (Optional) Contains exchange rates and shipping information
4. Specify the output filename
5. Click the "Convert Excel Files" button
6. After conversion is complete, you can download the generated Excel file

## File Format Requirements

### Input File
- Contains columns such as NO., DESCRIPTION, Model NO., etc.
- First 9 rows are for headers, actual data starts from row 10

### Reference File
- Must contain a MaterialCode column for matching with material codes in the input file
- Contains columns like product code, declaration elements, etc.

### Policy File
- Contains settings for exchange rates, shipping fees, insurance coefficients, etc.

## Frequently Asked Questions

1. **Can't start the application?**
   Make sure Python 3.8 or higher is installed on your system and dependencies are correctly installed.

2. **Incorrect file format?**
   Please download and refer to the sample template files to ensure your files meet the format requirements.

3. **Conversion failed?**
   Check if the uploaded files meet the requirements and look at the log information in the application for detailed error reasons.

## System Requirements

- Python 3.8 or higher
- Dependencies: streamlit, pandas, openpyxl, numpy

## Language Support

The application supports both Chinese and English interfaces, with Chinese as the default. You can switch languages in the sidebar. 