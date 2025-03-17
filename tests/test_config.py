import os
import sys
import importlib
import pytest

# Add the parent directory to sys.path to import the module
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Import the config module
import config


class TestConfig:
    """Test suite for configuration validation"""
    
    def test_preserved_columns_defined(self):
        """Test that PRESERVED_COLUMNS is properly defined"""
        assert hasattr(config, 'PRESERVED_COLUMNS'), "PRESERVED_COLUMNS not defined in config"
        assert isinstance(config.PRESERVED_COLUMNS, list), "PRESERVED_COLUMNS should be a list"
        assert len(config.PRESERVED_COLUMNS) > 0, "PRESERVED_COLUMNS should not be empty"
        
        # Check that essential columns are included
        essential_columns = ['项号', '品名', '型号', '数量', '单位', '单价', '总价']
        for col in essential_columns:
            assert col in config.PRESERVED_COLUMNS, f"Essential column '{col}' missing from PRESERVED_COLUMNS"
    
    def test_material_code_column_defined(self):
        """Test that MATERIAL_CODE_COLUMN is properly defined"""
        assert hasattr(config, 'MATERIAL_CODE_COLUMN'), "MATERIAL_CODE_COLUMN not defined in config"
        assert isinstance(config.MATERIAL_CODE_COLUMN, str), "MATERIAL_CODE_COLUMN should be a string"
        assert config.MATERIAL_CODE_COLUMN.strip(), "MATERIAL_CODE_COLUMN should not be empty"
    
    def test_matched_columns_defined(self):
        """Test that MATCHED_COLUMNS is properly defined"""
        assert hasattr(config, 'MATCHED_COLUMNS'), "MATCHED_COLUMNS not defined in config"
        assert isinstance(config.MATCHED_COLUMNS, list), "MATCHED_COLUMNS should be a list"
        assert len(config.MATCHED_COLUMNS) > 0, "MATCHED_COLUMNS should not be empty"
        
        # Check that essential matched columns are included
        essential_matched = ['商品编号', '申报要素']
        for col in essential_matched:
            assert col in config.MATCHED_COLUMNS, f"Essential matched column '{col}' missing from MATCHED_COLUMNS"
    
    def test_fixed_columns_defined(self):
        """Test that FIXED_COLUMNS is properly defined"""
        assert hasattr(config, 'FIXED_COLUMNS'), "FIXED_COLUMNS not defined in config"
        assert isinstance(config.FIXED_COLUMNS, dict), "FIXED_COLUMNS should be a dictionary"
        assert len(config.FIXED_COLUMNS) > 0, "FIXED_COLUMNS should not be empty"
        
        # Check that essential fixed columns are included
        essential_fixed = {
            '币制': '美元',
            '原产国（地区）': '中国',
            '最终目的国（地区）': '印度',
            '境内货源地': '深圳特区',
            '征免': '照章征税'
        }
        
        for col, default_value in essential_fixed.items():
            assert col in config.FIXED_COLUMNS, f"Essential fixed column '{col}' missing from FIXED_COLUMNS"
    
    def test_config_values_consistency(self):
        """Test that there are no overlaps between different column types"""
        # No column should be in both preserved and matched
        for col in config.PRESERVED_COLUMNS:
            assert col not in config.MATCHED_COLUMNS, f"Column '{col}' appears in both PRESERVED_COLUMNS and MATCHED_COLUMNS"
        
        # No column should be in both preserved and fixed
        for col in config.PRESERVED_COLUMNS:
            assert col not in config.FIXED_COLUMNS, f"Column '{col}' appears in both PRESERVED_COLUMNS and FIXED_COLUMNS"
        
        # No column should be in both matched and fixed
        for col in config.MATCHED_COLUMNS:
            assert col not in config.FIXED_COLUMNS, f"Column '{col}' appears in both MATCHED_COLUMNS and FIXED_COLUMNS"
    
    def test_output_file_name_defined(self):
        """Test that OUTPUT_FILE_NAME is properly defined"""
        assert hasattr(config, 'OUTPUT_FILE_NAME'), "OUTPUT_FILE_NAME not defined in config"
        assert isinstance(config.OUTPUT_FILE_NAME, str), "OUTPUT_FILE_NAME should be a string"
        assert config.OUTPUT_FILE_NAME.strip(), "OUTPUT_FILE_NAME should not be empty"
        assert config.OUTPUT_FILE_NAME.endswith(('.xlsx', '.xls')), "OUTPUT_FILE_NAME should have an Excel extension"


if __name__ == "__main__":
    pytest.main(["-v", __file__]) 