import os
from typing import List

def validate_excel_path(file_paths: List[str]) -> bool:
    """
    Validates if the given file paths exist and are Excel files.
    """
    all_valid = True
    for file_path in file_paths:
        if not os.path.exists(file_path):
            print(f"Error: File not found at '{file_path}'")
            all_valid = False
        if not (file_path.endswith(".xlsx") or file_path.endswith(".xls")):
            print(f"Error: '{file_path}' is not a valid Excel file (.xlsx or .xls).")
            all_valid = False
    return all_valid

def generate_output_filename(original_filename: str, prefix: str) -> str:
    """
    Generates a new filename for the output Excel file.
    """
    base, ext = os.path.splitext(original_filename)
    return f"{base}_{prefix}{ext}"

