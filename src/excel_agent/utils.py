import os
from typing import List
from src.excel_agent.config import Config

def validate_data_file_path(file_paths: List[str]) -> bool: # MODIFIED: Renamed function
    """
    Validates if the given file paths exist and are supported data files (Excel or CSV).
    """
    all_valid = True
    supported_extensions = tuple(Config.SUPPORTED_FILE_EXTENSIONS)
    
    for file_path in file_paths:
        if not os.path.exists(file_path):
            print(f"Error: File not found at '{file_path}'")
            all_valid = False
        
        # Check if the file extension is in the supported list
        file_ext = os.path.splitext(file_path)[1].lower()
        if file_ext not in supported_extensions:
            print(f"Error: '{file_path}' is not a supported file type. Supported types are: {', '.join(Config.SUPPORTED_FILE_EXTENSIONS)}.")
            all_valid = False
    return all_valid

def generate_output_filename(original_filename: str, prefix: str) -> str:
    """
    Generates a new filename for the output Excel file.
    """
    base, ext = os.path.splitext(original_filename)
    return f"{base}_{prefix}{ext}"

