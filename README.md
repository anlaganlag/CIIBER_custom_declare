# Excel Converter for Declaration List

这是一个用于转换Excel文件以符合报关要求的Streamlit应用程序。该应用程序可以处理输入Excel文件、参考Excel文件和政策文件，并生成符合报关格式要求的输出Excel文件。

## 一键安装和启动

### Windows用户

1. 确保您的计算机已安装Python 3.8或更高版本
2. 下载本项目的所有文件
3. 双击运行`setup_windows.bat`
4. 脚本将自动创建虚拟环境、安装依赖项并启动应用程序

### Mac/Linux用户

1. 确保您的计算机已安装Python 3.8或更高版本
2. 下载本项目的所有文件
3. 打开终端，导航到项目文件夹
4. 运行以下命令使脚本可执行：
   ```
   chmod +x setup_mac.sh
   ```
5. 执行脚本：
   ```
   ./setup_mac.sh
   ```
6. 脚本将自动创建虚拟环境、安装依赖项并启动应用程序

## 手动安装

如果一键脚本无法正常工作，您也可以手动执行以下步骤：

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

## 应用程序使用说明

1. 启动应用程序后，会在浏览器中打开一个窗口
2. 您可以在页面顶部下载示例文件模板：
   - 输入文件模板 (input_template.xlsx)
   - 参考文件模板 (reference_template.xlsx)
   - 政策文件模板 (policy_template.xlsx)
3. 上传您的文件：
   - 输入Excel文件：包含源数据，带有绿色表头
   - 参考Excel文件：包含物料代码和申报信息，用于黄色表头
   - 政策Excel文件：(可选)包含汇率和运输信息
4. 指定输出文件名
5. 点击"转换Excel文件"按钮
6. 转换完成后，您可以下载生成的Excel文件

## 文件格式要求

### 输入文件
- 包含列如NO.、DESCRIPTION、Model NO.等
- 前9行用于表头，实际数据从第10行开始

### 参考文件
- 必须包含MaterialCode列，用于与输入文件中的材料代码匹配
- 包含商品编号、申报要素等列

### 政策文件
- 包含汇率、运费、保费系数等设置

## 常见问题

1. **无法启动应用程序？**
   确保您的系统已安装Python 3.8或更高版本，并且正确安装了依赖项。

2. **文件格式不正确？**
   请下载并参考示例模板文件，确保您的文件格式符合要求。

3. **转换失败？**
   检查上传的文件是否符合要求，并查看应用程序中的日志信息，了解详细错误原因。

## 系统要求

- Python 3.8或更高版本
- 依赖项：streamlit, pandas, openpyxl, numpy

## 语言支持

应用程序支持中文和英文界面，默认为中文。您可以在侧边栏中切换语言。

## Overview

This application helps in preparing export declaration lists by:
1. Reading data from an input Excel file (source data with green headers)
2. Matching material codes with a reference Excel file (to obtain yellow headers data)
3. Adding fixed values as required for declaration purposes
4. Outputting a properly formatted Excel file ready for submission

## Features

- **Preserve Original Data**: Maintains essential data from the source file (green headers)
- **Material Code Matching**: Matches products by material code to add declaration fields (yellow headers)
- **Fixed Values**: Automatically adds standardized values (currency, origin country, etc.)
- **User-Friendly Interface**: Simple web interface built with Streamlit for easy use
- **Command Line Support**: Can be used via command line for batch processing

## Usage

### Web Interface

The simplest way to use the application is through the Streamlit web interface:

```bash
streamlit run streamlit_app.py
```

Through the web interface, you can:
1. Upload your input Excel file (with green headers)
2. Upload your reference Excel file (with material codes and declaration data)
3. Specify an output filename
4. Preview the data before conversion
5. Download the converted file

### Command Line

For batch processing or automation, use the command line interface:

```bash
python excel_converter.py input.xlsx reference.xlsx output.xlsx
```

## Configuration

The application uses a configuration file (`config.py`) to determine which columns to include:

### Green Headers (Preserved Columns)
- 项号 (NO.)
- 品名 (DESCRIPTION)
- 型号 (Model NO.)
- 数量 (Qty)
- 单位 (Unit)
- 单价 (Unit Price)
- 总价 (Amount)
- 净重 (net weight)

### Yellow Headers (Matched Columns)
- 商品编号 (Material Code)
- 申报要素 (Declaration Elements)

### Fixed Values
- 币制: 美元 (Currency: USD)
- 原产国（地区）: 中国 (Country of Origin: China)
- 最终目的国（地区）: 印度 (Destination Country: India)
- 境内货源地: 深圳特区 (Domestic Source: Shenzhen Special Zone)
- 征免: 照章征税 (Taxation: According to regulations)

## Process Flow

```mermaid
flowchart LR
    %% Main flow with larger nodes and clearer labels
    InputFile["Input Excel File\n(Source Data)"] ==> |"Green Headers"| Processor
    ReferenceFile["Reference Excel File\n(Lookup Data)"] ==> |"Yellow Headers"| Processor
    FixedValues["Fixed Values\n(Constants)"] ==> Processor
    Processor["Excel Converter"] ==> OutputFile["Final Excel File\n(Declaration Ready)"]
    
    %% Define styles
    classDef greenBox fill:#d4ffda,stroke:#00a300,stroke-width:2px,color:#004400,font-weight:bold
    classDef yellowBox fill:#ffffcc,stroke:#ffd700,stroke-width:2px,color:#8b6914,font-weight:bold
    classDef blueBox fill:#ecf4ff,stroke:#0078d7,stroke-width:2px,color:#003366,font-weight:bold
    classDef grayBox fill:#f0f0f0,stroke:#444444,stroke-width:2px,color:#333333,font-weight:bold
    
    %% Column details with subgraphs
    subgraph GreenColumns["Green Header Columns"]
        direction TB
        G1["项号 (NO.)"]
        G2["品名 (DESCRIPTION)"]
        G3["型号 (Model NO.)"]
        G4["数量 (Qty)"]
        G5["单位 (Unit)"]
        G6["单价 (Unit Price)"]
        G7["总价 (Amount)"]
        G8["净重 (Net Weight)"]
    end
    
    subgraph YellowColumns["Yellow Header Columns"]
        direction TB
        Y1["商品编号 (Material Code)"]
        Y2["申报要素 (Declaration Elements)"]
    end
    
    subgraph FixedColumns["Fixed Value Columns"]
        direction TB
        F1["币制: 美元\n(Currency: USD)"]
        F2["原产国（地区）: 中国\n(Origin: China)"]
        F3["最终目的国（地区）: 印度\n(Destination: India)"]
        F4["境内货源地: 深圳特区\n(Source: Shenzhen)"]
        F5["征免: 照章征税\n(Taxation: Standard)"]
    end
    
    %% Connect subgraphs to main flow
    GreenColumns -.-> InputFile
    YellowColumns -.-> ReferenceFile
    FixedColumns -.-> FixedValues
    
    %% Apply styles
    class GreenColumns,G1,G2,G3,G4,G5,G6,G7,G8 greenBox
    class YellowColumns,Y1,Y2 yellowBox
    class FixedColumns,F1,F2,F3,F4,F5 grayBox
    class InputFile,ReferenceFile,Processor,OutputFile,FixedValues blueBox
```

## Testing

The project includes a comprehensive test suite to ensure reliability and correct functionality. The tests are organized in the `tests/` directory.

### Test Structure

- `tests/test_excel_converter.py` - Tests for the core Excel conversion functionality
- `tests/test_config.py` - Tests for the configuration settings
- `tests/test_streamlit_app.py` - Tests for the Streamlit web interface
- `tests/conftest.py` - Common fixtures and setup for all tests

### Running Tests

To run the tests, you'll need to install the test dependencies:

```bash
pip install -r requirements-test.txt
```

Then, run the tests using pytest:

```bash
pytest
```

For more detailed output, use:

```bash
pytest -v
```

To generate a test coverage report:

```bash
pytest --cov=. --cov-report=html
```

This will create an HTML coverage report in the `htmlcov/` directory.

## Requirements

- Python 3.6+
- pandas
- openpyxl
- streamlit (for web interface)

## Installation

```bash
pip install pandas openpyxl streamlit
```

## License

This project is proprietary and intended for internal use only.