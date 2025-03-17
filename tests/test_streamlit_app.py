import os
import sys
import pytest
import streamlit as st
import pandas as pd
import tempfile
from unittest.mock import patch, MagicMock

# Add the parent directory to sys.path to import the module
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# We need to mock Streamlit since it expects to run in a Streamlit context
class TestStreamlitApp:
    """Test suite for the Streamlit application functionality"""
    
    @pytest.fixture
    def mock_streamlit_session(self):
        """Mock Streamlit session for testing"""
        # Create mock objects for st.* functions that will be used in the app
        with patch('streamlit.title') as mock_title, \
             patch('streamlit.write') as mock_write, \
             patch('streamlit.header') as mock_header, \
             patch('streamlit.subheader') as mock_subheader, \
             patch('streamlit.file_uploader') as mock_file_uploader, \
             patch('streamlit.text_input') as mock_text_input, \
             patch('streamlit.button') as mock_button, \
             patch('streamlit.success') as mock_success, \
             patch('streamlit.error') as mock_error, \
             patch('streamlit.spinner') as mock_spinner, \
             patch('streamlit.dataframe') as mock_dataframe, \
             patch('streamlit.download_button') as mock_download_button:
            
            # Configure the mocks to return appropriate values
            mock_title.return_value = None
            mock_write.return_value = None
            mock_header.return_value = None
            mock_subheader.return_value = None
            mock_text_input.return_value = "output_test.xlsx"
            mock_button.return_value = True
            mock_success.return_value = None
            mock_error.return_value = None
            mock_dataframe.return_value = None
            mock_download_button.return_value = None
            
            # Create spinner context manager mock
            spinner_cm = MagicMock()
            spinner_cm.__enter__.return_value = None
            spinner_cm.__exit__.return_value = None
            mock_spinner.return_value = spinner_cm
            
            yield {
                'title': mock_title,
                'write': mock_write,
                'header': mock_header,
                'subheader': mock_subheader,
                'file_uploader': mock_file_uploader,
                'text_input': mock_text_input,
                'button': mock_button,
                'success': mock_success,
                'error': mock_error,
                'spinner': mock_spinner,
                'dataframe': mock_dataframe,
                'download_button': mock_download_button
            }
    
    @pytest.fixture
    def test_excel_files(self):
        """Create temporary Excel files for testing"""
        # Create temp directory
        temp_dir = tempfile.mkdtemp()
        
        # Create a test input Excel file
        input_df = pd.DataFrame({
            'NO.': [1, 2],
            'DESCRIPTION': ['Test Product A', 'Test Product B'],
            'Qty': [10, 20],
            'Unit Price': [100, 200],
            'Amount': [1000, 4000]
        })
        
        # Create a test reference Excel file
        reference_df = pd.DataFrame({
            'MaterialCode': ['TEST001', 'TEST002'],
            '商品编号': ['SH001', 'SH002'],
            '申报要素': ['Test Element A', 'Test Element B']
        })
        
        # Save the dataframes to Excel files
        input_path = os.path.join(temp_dir, 'input_test.xlsx')
        reference_path = os.path.join(temp_dir, 'reference_test.xlsx')
        output_path = os.path.join(temp_dir, 'output_test.xlsx')
        
        input_df.to_excel(input_path, index=False)
        reference_df.to_excel(reference_path, index=False)
        
        # Create file-like objects that mimic uploaded files
        class MockUploadedFile:
            def __init__(self, file_path):
                self.file_path = file_path
                self._buffer = open(file_path, 'rb').read()
            
            def getvalue(self):
                return self._buffer
            
            def read(self):
                return self._buffer
            
            def getbuffer(self):
                return self._buffer
        
        mock_input_file = MockUploadedFile(input_path)
        mock_reference_file = MockUploadedFile(reference_path)
        
        yield {
            'temp_dir': temp_dir,
            'input_path': input_path,
            'reference_path': reference_path,
            'output_path': output_path,
            'mock_input_file': mock_input_file,
            'mock_reference_file': mock_reference_file
        }
        
        # Clean up
        os.unlink(input_path)
        os.unlink(reference_path)
        if os.path.exists(output_path):
            os.unlink(output_path)
        os.rmdir(temp_dir)
    
    def test_streamlit_app_successful_conversion(self, mock_streamlit_session, test_excel_files):
        """Test successful conversion flow in Streamlit app"""
        # Import the app module here to ensure Streamlit is mocked first
        from streamlit_app import main
        
        # Configure file uploader mock to return our test files
        mock_streamlit_session['file_uploader'].side_effect = [
            test_excel_files['mock_input_file'],
            test_excel_files['mock_reference_file']
        ]
        
        # Mock convert_excel to return a DataFrame and actually create the output file
        with patch('streamlit_app.convert_excel') as mock_convert:
            result_df = pd.DataFrame({'test': [1, 2, 3]})
            mock_convert.return_value = result_df
            
            # Create the expected output file
            result_df.to_excel(test_excel_files['output_path'], index=False)
            
            # Call the app's main function
            with patch('streamlit_app.open', create=True) as mock_open:
                mock_open.return_value.__enter__.return_value = open(test_excel_files['output_path'], 'rb')
                main()
            
            # Verify the app flow
            mock_streamlit_session['title'].assert_called_once()  # Title should be shown
            assert mock_streamlit_session['file_uploader'].call_count == 2  # Two file uploaders
            mock_streamlit_session['button'].assert_called_once()  # Convert button clicked
            mock_convert.assert_called_once()  # Conversion function called
            mock_streamlit_session['success'].assert_called_once()  # Success message shown
            mock_streamlit_session['download_button'].assert_called_once()  # Download button shown
    
    def test_streamlit_app_missing_files(self, mock_streamlit_session):
        """Test error handling when files are missing"""
        # Import the app module here to ensure Streamlit is mocked first
        from streamlit_app import main
        
        # Configure file uploader mock to return None (no files uploaded)
        mock_streamlit_session['file_uploader'].return_value = None
        
        # Call the app's main function
        main()
        
        # Verify error is shown when button is clicked but files are missing
        mock_streamlit_session['button'].assert_called_once()  # Convert button clicked
        mock_streamlit_session['error'].assert_called_once()  # Error message shown
    
    def test_streamlit_app_conversion_error(self, mock_streamlit_session, test_excel_files):
        """Test error handling when conversion fails"""
        # Import the app module here to ensure Streamlit is mocked first
        from streamlit_app import main
        
        # Configure file uploader mock to return our test files
        mock_streamlit_session['file_uploader'].side_effect = [
            test_excel_files['mock_input_file'],
            test_excel_files['mock_reference_file']
        ]
        
        # Mock convert_excel to raise an exception
        with patch('streamlit_app.convert_excel') as mock_convert:
            mock_convert.side_effect = Exception("Test conversion error")
            
            # Call the app's main function
            main()
            
            # Verify error is shown when conversion fails
            mock_streamlit_session['button'].assert_called_once()  # Convert button clicked
            assert mock_streamlit_session['error'].call_count >= 1  # Error message shown


if __name__ == "__main__":
    pytest.main(["-v", __file__]) 