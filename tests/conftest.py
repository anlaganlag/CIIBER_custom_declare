import os
import sys
import pytest
import pandas as pd
import tempfile
import shutil

# Add the parent directory to sys.path to allow importing the application modules
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

@pytest.fixture(scope="session")
def sample_data():
    """Provides sample data dictionaries for testing"""
    return {
        "green_headers": {
            "NO.": [1, 2, 3],
            "DESCRIPTION": ["Product A", "Product B", "Product C"],
            "Model NO.": ["A-100", "B-200", "C-300"],
            "Qty": [10, 20, 30],
            "Unit": ["pcs", "pcs", "set"],
            "Unit Price": [100.0, 200.0, 300.0],
            "Amount": [1000.0, 4000.0, 9000.0],
            "net weight": [5.0, 10.0, 15.0]
        },
        "yellow_headers": {
            "MaterialCode": ["MC001", "MC002", "MC003"],
            "商品编号": ["SH001", "SH002", "SH003"],
            "申报要素": ["Element A", "Element B", "Element C"]
        },
        "fixed_values": {
            "币制": "美元",
            "原产国（地区）": "中国",
            "最终目的国（地区）": "印度",
            "境内货源地": "深圳特区",
            "征免": "照章征税"
        }
    }

@pytest.fixture(scope="session")
def temp_test_directory():
    """Creates a temporary directory for test files"""
    temp_dir = tempfile.mkdtemp()
    yield temp_dir
    # Clean up after all tests
    shutil.rmtree(temp_dir)

@pytest.fixture
def excel_converter_import():
    """Makes sure the excel_converter module can be imported"""
    try:
        import excel_converter
        return excel_converter
    except ImportError as e:
        pytest.skip(f"Skipping test: Could not import excel_converter module: {e}")

@pytest.fixture
def config_import():
    """Makes sure the config module can be imported"""
    try:
        import config
        return config
    except ImportError as e:
        pytest.skip(f"Skipping test: Could not import config module: {e}")

@pytest.fixture
def streamlit_app_import():
    """Makes sure the streamlit_app module can be imported"""
    try:
        import streamlit_app
        return streamlit_app
    except ImportError as e:
        pytest.skip(f"Skipping test: Could not import streamlit_app module: {e}")

@pytest.fixture
def create_excel_file():
    """Fixture that creates an Excel file from a dataframe"""
    def _create_excel_file(dataframe, filepath):
        """Inner function to create Excel file"""
        dataframe.to_excel(filepath, index=False)
        return filepath
    return _create_excel_file 