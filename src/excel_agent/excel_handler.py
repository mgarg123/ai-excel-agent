import pandas as pd
import os
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from typing import List, Dict, Any, Union
from abc import ABC, abstractmethod

from src.excel_agent.tools import tool
from src.excel_agent.config import Config
from src.excel_agent.output.abstract_output_handler import AbstractOutputHandler

class ExcelHandler:
    """
    Handles all Excel and DataFrame operations.
    Manages the active DataFrame and provides tools for manipulation.
    """
    def __init__(self, file_path: str, output_handler: AbstractOutputHandler):
        self.file_path = file_path
        self.active_df: pd.DataFrame = None
        self.active_sheet_name: str = None
        self.output_handler = output_handler

    def _load_data_internal(self, file_path: str, sheet_name: str = None) -> pd.DataFrame:
        """
        Internal method to load data from an Excel or CSV file.
        Does not set active_df or active_sheet_name.
        """
        try:
            file_ext = os.path.splitext(file_path)[1].lower()
            if file_ext == ".csv":
                df = pd.read_csv(file_path)
                self.output_handler.print_message(f"Successfully loaded CSV file: '{file_path}'", style='success')
            elif file_ext in [".xlsx", ".xls"]:
                if sheet_name:
                    df = pd.read_excel(file_path, sheet_name=sheet_name)
                    self.output_handler.print_message(f"Successfully loaded Excel file: '{file_path}' sheet '{sheet_name}'", style='success')
                else:
                    # If no sheet name specified for Excel, load the first sheet
                    df = pd.read_excel(file_path)
                    self.output_handler.print_message(f"Successfully loaded Excel file: '{file_path}' (first sheet)", style='success')
            else:
                self.output_handler.show_error(f"Unsupported file type: {file_ext}. Only .xlsx, .xls, and .csv are supported.")
                return None
            
            if df.empty:
                self.output_handler.show_warning(f"Loaded data from '{file_path}' (sheet '{sheet_name}' if applicable) is empty.")
            return df
        except FileNotFoundError:
            self.output_handler.show_error(f"File not found: '{file_path}'")
            return None
        except Exception as e:
            self.output_handler.show_error(f"Error loading data from '{file_path}' (sheet '{sheet_name}' if applicable): {e}")
            return None

    def get_sheet_names(self) -> List[str]:
        """
        Returns a list of sheet names in the Excel file.
        For CSV, returns ['Sheet1'] as a default.
        """
        file_ext = os.path.splitext(self.file_path)[1].lower()
        if file_ext == ".csv":
            return ["Sheet1"] # CSV files don't have sheets, return a default
        elif file_ext in [".xlsx", ".xls"]:
            try:
                xls = pd.ExcelFile(self.file_path)
                return xls.sheet_names
            except FileNotFoundError:
                self.output_handler.show_error(f"File not found: '{self.file_path}'")
                return []
            except Exception as e:
                self.output_handler.show_error(f"Error reading sheet names from '{self.file_path}': {e}")
                return []
        else:
            return [] # Should be caught by validation earlier

    def get_column_headers(self, sheet_name: str = None) -> List[str]:
        """
        Returns a list of column headers for the specified sheet or active DataFrame.
        """
        if self.active_df is not None and (sheet_name is None or sheet_name == self.active_sheet_name):
            return self.active_df.columns.tolist()
        else:
            # Load data just to get headers if not active or different sheet requested
            df = self._load_data_internal(self.file_path, sheet_name=sheet_name)
            if df is not None:
                return df.columns.tolist()
            return []

    @tool(description="Loads data from a specified sheet of an Excel or CSV file and sets it as the active DataFrame for subsequent operations. This is the first tool to call for any data analysis.")
    def load_and_display_data(self, file_path: str, sheet_name: str = None) -> pd.DataFrame:
        """
        Loads data from the specified file and sheet, sets it as the active DataFrame,
        and displays its head.
        """
        df = self._load_data_internal(file_path, sheet_name)
        if df is not None:
            self.active_df = df
            self.active_sheet_name = sheet_name if sheet_name else "Sheet1" if os.path.splitext(file_path)[1].lower() == ".csv" else self.get_sheet_names()[0]
            self.output_handler.print_message(f"Active DataFrame set to '{file_path}' (Sheet: {self.active_sheet_name}). Displaying head:", style='info')
            return self.active_df.head()
        return None

    @tool(description="Filters the active DataFrame based on a query string and displays the result. Use this when the user asks to 'filter data', 'show records where', or 'select rows based on criteria'.")
    def filter_and_display_dataframe(self, query_string: str) -> pd.DataFrame:
        """
        Filters the active DataFrame using a query string.
        """
        if self.active_df is None:
            self.output_handler.show_error("No active DataFrame. Please load data first.")
            return None
        try:
            filtered_df = self.active_df.query(query_string)
            self.active_df = filtered_df.copy() # MODIFIED: Update active_df
            self.output_handler.show_success(f"DataFrame filtered by query: '{query_string}'. Displaying head of filtered data:")
            return self.active_df.head()
        except Exception as e:
            self.output_handler.show_error(f"Error filtering DataFrame: {e}. Please check your query string and column names.")
            return None

    @tool(description="Groups the active DataFrame by specified columns, aggregates a target column, and displays the result. Use this for 'summarize by', 'total by', 'average by', etc.")
    def group_and_display_dataframe(self, group_by_columns: List[str], target_column: str, aggregation_type: str) -> pd.DataFrame:
        """
        Groups the active DataFrame by specified columns and aggregates a target column.
        Aggregation types: 'sum', 'mean', 'count', 'min', 'max', 'std'.
        The aggregated column will be renamed for clarity.
        """
        if self.active_df is None:
            self.output_handler.show_error("No active DataFrame. Please load data first.")
            return None
        if not all(col in self.active_df.columns for col in group_by_columns):
            self.output_handler.show_error(f"One or more group-by columns not found in DataFrame: {group_by_columns}")
            return None
        if target_column not in self.active_df.columns:
            self.output_handler.show_error(f"Target column '{target_column}' not found in DataFrame.")
            return None
        
        valid_aggregations = ['sum', 'mean', 'count', 'min', 'max', 'std']
        if aggregation_type not in valid_aggregations:
            self.output_handler.show_error(f"Invalid aggregation type: '{aggregation_type}'. Must be one of {valid_aggregations}.")
            return None

        try:
            # Perform aggregation
            grouped_series = self.active_df.groupby(group_by_columns)[target_column].agg(aggregation_type)
            grouped_df = grouped_series.reset_index()

            # Rename the aggregated column predictably
            if aggregation_type == 'count':
                new_agg_column_name = 'CountOfRecords'
            else:
                new_agg_column_name = f"{target_column}_{aggregation_type}"
            
            grouped_df.rename(columns={grouped_df.columns[-1]: new_agg_column_name}, inplace=True)
            
            self.active_df = grouped_df.copy() # MODIFIED: Update active_df
            self.output_handler.show_success(f"DataFrame grouped by {group_by_columns} and '{target_column}' aggregated by '{aggregation_type}'. New aggregated column: '{new_agg_column_name}'. Displaying result:")
            return self.active_df # Return the full grouped DataFrame for display

        except Exception as e:
            self.output_handler.show_error(f"Error grouping DataFrame: {e}")
            return None

    @tool(description="Sorts the active DataFrame by one or more columns and displays the head of the sorted result. Use this for 'sort by', 'order by', 'highest/lowest values'.")
    def sort_and_display_dataframe(self, sort_by_columns: List[str], ascending: bool = True) -> pd.DataFrame:
        """
        Sorts the active DataFrame by specified columns.
        """
        if self.active_df is None:
            self.output_handler.show_error("No active DataFrame. Please load data first.")
            return None
        if not all(col in self.active_df.columns for col in sort_by_columns):
            self.output_handler.show_error(f"One or more sort-by columns not found in DataFrame: {sort_by_columns}")
            return None
        try:
            sorted_df = self.active_df.sort_values(by=sort_by_columns, ascending=ascending)
            self.active_df = sorted_df.copy() # MODIFIED: Update active_df
            self.output_handler.show_success(f"DataFrame sorted by {sort_by_columns} (ascending={ascending}). Displaying head:")
            return self.active_df.head()
        except Exception as e:
            self.output_handler.show_error(f"Error sorting DataFrame: {e}")
            return None

    @tool(description="Adds a new column to the active DataFrame based on a pandas formula string and displays the updated DataFrame. Use this for 'calculate new column', 'add column', 'derive new metric'.")
    def add_column_and_display_dataframe(self, new_column_name: str, formula: str) -> pd.DataFrame:
        """
        Adds a new column to the active DataFrame based on a formula string.
        Formula can use existing column names (e.g., 'col1 * col2').
        """
        if self.active_df is None:
            self.output_handler.show_error("No active DataFrame. Please load data first.")
            return None
        try:
            self.active_df[new_column_name] = self.active_df.eval(formula)
            self.output_handler.show_success(f"New column '{new_column_name}' added with formula '{formula}'. Displaying head:")
            return self.active_df.head()
        except Exception as e:
            self.output_handler.show_error(f"Error adding column '{new_column_name}' with formula '{formula}': {e}. Check formula syntax and column names.")
            return None

    @tool(description="Calculates a single scalar value (e.g., sum, average, count) for a specified column in the active DataFrame, optionally after filtering. Use this for 'what is the total', 'average', 'count of', 'min/max value'.")
    def calculate_scalar_value(self, column: str, aggregation_type: str, query_string: str = None) -> Union[int, float]:
        """
        Calculates a scalar value (sum, mean, count, min, max, std) for a column,
        optionally after filtering.
        """
        if self.active_df is None:
            self.output_handler.show_error("No active DataFrame. Please load data first.")
            return None
        if column not in self.active_df.columns:
            self.output_handler.show_error(f"Column '{column}' not found in DataFrame.")
            return None
        
        valid_aggregations = ['sum', 'mean', 'count', 'min', 'max', 'std']
        if aggregation_type not in valid_aggregations:
            self.output_handler.show_error(f"Invalid aggregation type: '{aggregation_type}'. Must be one of {valid_aggregations}.")
            return None

        df_to_aggregate = self.active_df
        if query_string:
            try:
                df_to_aggregate = self.active_df.query(query_string)
                if df_to_aggregate.empty:
                    self.output_handler.show_warning(f"Filtering by '{query_string}' resulted in an empty DataFrame for scalar calculation.")
                    return None
            except Exception as e:
                self.output_handler.show_error(f"Error applying query string '{query_string}' for scalar calculation: {e}. Calculation aborted.")
                return None

        try:
            if aggregation_type == 'sum':
                result = df_to_aggregate[column].sum()
            elif aggregation_type == 'mean':
                result = df_to_aggregate[column].mean()
            elif aggregation_type == 'count':
                result = df_to_aggregate[column].count()
            elif aggregation_type == 'min':
                result = df_to_aggregate[column].min()
            elif aggregation_type == 'max':
                result = df_to_aggregate[column].max()
            elif aggregation_type == 'std':
                result = df_to_aggregate[column].std()
            
            self.output_handler.show_success(f"Calculated {aggregation_type} of '{column}' (filtered by '{query_string}' if applicable): {result}")
            return result
        except Exception as e:
            self.output_handler.show_error(f"Error calculating {aggregation_type} for column '{column}': {e}")
            return None

    @tool(description="Saves the active DataFrame to a new Excel file. Use this when the user asks to 'save data', 'export to new file', or 'create a new Excel file'.")
    def save_dataframe_to_new_excel(self, output_file_path: str) -> str:
        """
        Saves the active DataFrame to a new Excel file.
        """
        if self.active_df is None:
            self.output_handler.show_error("No active DataFrame to save. Please load data first.")
            return None
        try:
            # Ensure the directory exists
            output_dir = os.path.dirname(output_file_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            self.active_df.to_excel(output_file_path, index=False)
            self.output_handler.show_success(f"DataFrame successfully saved to '{output_file_path}'")
            return output_file_path
        except Exception as e:
            self.output_handler.show_error(f"Error saving DataFrame to '{output_file_path}': {e}")
            return None

    @tool(description="Applies an Excel-like formula to a specified column in the active DataFrame. This is for simple, direct cell-wise operations. For complex column additions, use 'add_column_and_display_dataframe'.")
    def apply_excel_formula(self, column: str, formula: str) -> pd.DataFrame:
        """
        Applies a simple Excel-like formula to a column.
        Example: 'A1*2' or 'B1+C1'. This is a simplified tool.
        For more complex pandas expressions, use add_column_and_display_dataframe.
        """
        if self.active_df is None:
            self.output_handler.show_error("No active DataFrame. Please load data first.")
            return None
        if column not in self.active_df.columns:
            self.output_handler.show_error(f"Column '{column}' not found in DataFrame.")
            return None
        
        try:
            # This is a very basic interpretation. A real Excel formula parser would be complex.
            # For now, assume formula is a simple operation on the column itself or a scalar.
            # E.g., "value * 1.1" or "value + 100"
            # This tool is less powerful than add_column_and_display_dataframe.
            # It's more for direct cell-wise operations.
            # Example: if formula is "value * 1.1", it means active_df[column] * 1.1
            
            # Attempt to evaluate the formula using the column's series
            # This is a very naive implementation and might fail for complex formulas.
            # A robust solution would involve a proper formula parser.
            # For now, we'll try to use pandas eval with a placeholder 'value'
            temp_df = pd.DataFrame({'value': self.active_df[column].copy()})
            temp_df['result'] = temp_df.eval(formula)
            self.active_df[column] = temp_df['result']

            self.output_handler.show_success(f"Formula '{formula}' applied to column '{column}'. Displaying head:")
            return self.active_df.head()
        except Exception as e:
            self.output_handler.show_error(f"Error applying formula '{formula}' to column '{column}': {e}. This tool supports simple operations like 'value * 1.1' or 'value + 100'. For complex expressions, use 'add_column_and_display_dataframe'.")
            return None

    @tool(description="Applies basic formatting (e.g., number format, alignment) to a specified column in the active DataFrame. Note: This tool primarily affects display, not underlying data type.")
    def apply_formatting(self, column: str, format_type: str, format_value: Any = None) -> pd.DataFrame:
        """
        Applies basic formatting to a column.
        Note: This is a conceptual tool. Pandas DataFrames don't directly store display formats.
        This would typically involve converting data types or preparing for export.
        For now, it will attempt basic type conversion or just acknowledge.
        """
        if self.active_df is None:
            self.output_handler.show_error("No active DataFrame. Please load data first.")
            return None
        if column not in self.active_df.columns:
            self.output_handler.show_error(f"Column '{column}' not found in DataFrame.")
            return None
        
        try:
            if format_type == "currency" and pd.api.types.is_numeric_dtype(self.active_df[column]):
                # This changes the data to string for display, not ideal for further calculations
                self.active_df[column] = self.active_df[column].apply(lambda x: f"${x:,.2f}" if pd.notna(x) else x)
                self.output_handler.show_success(f"Column '{column}' formatted as currency. Displaying head:")
            elif format_type == "percentage" and pd.api.types.is_numeric_dtype(self.active_df[column]):
                self.active_df[column] = self.active_df[column].apply(lambda x: f"{x:.2%}" if pd.notna(x) else x)
                self.output_handler.show_success(f"Column '{column}' formatted as percentage. Displaying head:")
            elif format_type == "datetime" and pd.api.types.is_datetime64_any_dtype(self.active_df[column]):
                # Example: 'YYYY-MM-DD'
                self.active_df[column] = self.active_df[column].dt.strftime(format_value or '%Y-%m-%d')
                self.output_handler.show_success(f"Column '{column}' formatted as datetime ('{format_value or '%Y-%m-%d'}'). Displaying head:")
            else:
                self.output_handler.show_warning(f"Unsupported format type '{format_type}' or column '{column}' is not suitable for this format. No changes applied.")
                return None # No change to DF
            return self.active_df.head()
        except Exception as e:
            self.output_handler.show_error(f"Error applying formatting to column '{column}': {e}")
            return None

    @tool(description="Handles missing values in a specified column of the active DataFrame by filling them with a given value or a strategy (mean, median, mode, ffill, bfill). Use this for 'fill missing values', 'handle NaNs'.")
    def handle_missing_values(self, column: str, strategy: str, fill_value: Any = None) -> pd.DataFrame:
        """
        Handles missing values in a column using specified strategy.
        Strategies: 'fill_value', 'mean', 'median', 'mode', 'drop_row', 'drop_column'.
        For advanced imputation (ffill, bfill, interpolate), use 'impute_missing_values_advanced'.
        """
        if self.active_df is None:
            self.output_handler.show_error("No active DataFrame. Please load data first.")
            return None
        if column not in self.active_df.columns:
            self.output_handler.show_error(f"Column '{column}' not found in DataFrame.")
            return None
        
        original_nan_count = self.active_df[column].isnull().sum()
        if original_nan_count == 0:
            self.output_handler.show_warning(f"No missing values found in column '{column}'. No action taken.")
            return self.active_df.head()

        try:
            if strategy == 'fill_value':
                if fill_value is None:
                    self.output_handler.show_error("fill_value strategy requires a 'fill_value' parameter.")
                    return None
                self.active_df[column].fillna(fill_value, inplace=True)
                self.output_handler.show_success(f"Missing values in '{column}' filled with '{fill_value}'.")
            elif strategy == 'mean':
                if pd.api.types.is_numeric_dtype(self.active_df[column]):
                    fill_val = self.active_df[column].mean()
                    self.active_df[column].fillna(fill_val, inplace=True)
                    self.output_handler.show_success(f"Missing values in '{column}' filled with mean ({fill_val:.2f}).")
                else:
                    self.output_handler.show_error(f"Cannot fill with mean: Column '{column}' is not numeric.")
                    return None
            elif strategy == 'median':
                if pd.api.types.is_numeric_dtype(self.active_df[column]):
                    fill_val = self.active_df[column].median()
                    self.active_df[column].fillna(fill_val, inplace=True)
                    self.output_handler.show_success(f"Missing values in '{column}' filled with median ({fill_val:.2f}).")
                else:
                    self.output_handler.show_error(f"Cannot fill with median: Column '{column}' is not numeric.")
                    return None
            elif strategy == 'mode':
                fill_val = self.active_df[column].mode()[0] # mode can return multiple, take first
                self.active_df[column].fillna(fill_val, inplace=True)
                self.output_handler.show_success(f"Missing values in '{column}' filled with mode ('{fill_val}').")
            elif strategy == 'drop_row':
                initial_rows = len(self.active_df)
                self.active_df.dropna(subset=[column], inplace=True, errors='ignore')
                rows_dropped = initial_rows - len(self.active_df)
                self.output_handler.show_success(f"Rows with missing values in '{column}' dropped. {rows_dropped} rows removed.")
            elif strategy == 'drop_column':
                self.active_df.drop(columns=[column], inplace=True)
                self.output_handler.show_success(f"Column '{column}' with missing values dropped.")
            else:
                self.output_handler.show_error(f"Invalid strategy: '{strategy}'. Supported: 'fill_value', 'mean', 'median', 'mode', 'drop_row', 'drop_column'. For ffill/bfill/interpolate, use 'impute_missing_values_advanced'.")
                return None
            
            return self.active_df.head()
        except Exception as e:
            self.output_handler.show_error(f"Error handling missing values in column '{column}' with strategy '{strategy}': {e}")
            return None

    @tool(description="Removes duplicate rows from the active DataFrame based on all columns or a subset of columns. Use this for 'remove duplicates', 'deduplicate data'.")
    def remove_duplicates(self, subset_columns: List[str] = None) -> pd.DataFrame:
        """
        Removes duplicate rows from the active DataFrame.
        """
        if self.active_df is None:
            self.output_handler.show_error("No active DataFrame. Please load data first.")
            return None
        
        initial_rows = len(self.active_df)
        if subset_columns:
            if not all(col in self.active_df.columns for col in subset_columns):
                self.output_handler.show_error(f"One or more subset columns not found in DataFrame: {subset_columns}")
                return None
            self.active_df.drop_duplicates(subset=subset_columns, inplace=True)
            self.output_handler.show_success(f"Duplicate rows removed based on columns {subset_columns}.")
        else:
            self.active_df.drop_duplicates(inplace=True)
            self.output_handler.show_success("All duplicate rows removed.")
        
        rows_removed = initial_rows - len(self.active_df)
        self.output_handler.print_message(f"{rows_removed} duplicate rows removed.", style='info')
        return self.active_df.head()

    @tool(description="Renames a column in the active DataFrame. Use this for 'rename column', 'change column name'.")
    def rename_column(self, old_column_name: str, new_column_name: str) -> pd.DataFrame:
        """
        Renames a column in the active DataFrame.
        """
        if self.active_df is None:
            self.output_handler.show_error("No active DataFrame. Please load data first.")
            return None
        if old_column_name not in self.active_df.columns:
            self.output_handler.show_error(f"Column '{old_column_name}' not found in DataFrame.")
            return None
        try:
            self.active_df.rename(columns={old_column_name: new_column_name}, inplace=True)
            self.output_handler.show_success(f"Column '{old_column_name}' renamed to '{new_column_name}'. Displaying head:")
            return self.active_df.head()
        except Exception as e:
            self.output_handler.show_error(f"Error renaming column '{old_column_name}': {e}")
            return None

    @tool(description="Selects a subset of columns from the active DataFrame and displays the result. Use this for 'show only these columns', 'select columns'.")
    def select_columns_and_display(self, columns_to_select: List[str]) -> pd.DataFrame:
        """
        Selects a subset of columns from the active DataFrame.
        """
        if self.active_df is None:
            self.output_handler.show_error("No active DataFrame. Please load data first.")
            return None
        if not all(col in self.active_df.columns for col in columns_to_select):
            self.output_handler.show_error(f"One or more columns to select not found in DataFrame: {columns_to_select}")
            return None
        try:
            selected_df = self.active_df[columns_to_select]
            self.active_df = selected_df.copy() # MODIFIED: Update active_df
            self.output_handler.show_success(f"Selected columns: {columns_to_select}. Displaying head:")
            return self.active_df.head()
        except Exception as e:
            self.output_handler.show_error(f"Error selecting columns: {e}")
            return None

    @tool(description="Generates descriptive statistics (e.g., mean, std, min, max) for numeric columns in the active DataFrame. Use this for 'summarize data', 'get statistics'.")
    def get_descriptive_statistics(self) -> pd.DataFrame:
        """
        Generates descriptive statistics for the active DataFrame.
        """
        if self.active_df is None:
            self.output_handler.show_error("No active DataFrame. Please load data first.")
            return None
        try:
            stats_df = self.active_df.describe()
            self.output_handler.show_success("Descriptive statistics for the active DataFrame:")
            return stats_df
        except Exception as e:
            self.output_handler.show_error(f"Error generating descriptive statistics: {e}")
            return None

    @tool(description="Deletes specified rows or columns from the active DataFrame. Use this for 'remove rows', 'delete columns'.")
    def delete_rows_or_columns(self, target_type: str, identifiers: Union[List[Any], List[str]]) -> pd.DataFrame:
        """
        Deletes rows by index or columns by name.
        'target_type': 'rows' or 'columns'.
        'identifiers': list of row indices (for 'rows') or column names (for 'columns').
        """
        if self.active_df is None:
            self.output_handler.show_error("No active DataFrame. Please load data first.")
            return None
        
        try:
            if target_type == 'rows':
                initial_rows = len(self.active_df)
                self.active_df.drop(index=identifiers, inplace=True, errors='ignore')
                rows_removed = initial_rows - len(self.active_df)
                self.output_handler.show_success(f"{rows_removed} rows deleted. Displaying head:")
            elif target_type == 'columns':
                if not all(col in self.active_df.columns for col in identifiers):
                    self.output_handler.show_error(f"One or more columns to delete not found in DataFrame: {identifiers}")
                    return None
                self.active_df.drop(columns=identifiers, inplace=True)
                self.output_handler.show_success(f"Columns {identifiers} deleted. Displaying head:")
            else:
                self.output_handler.show_error("Invalid target_type. Must be 'rows' or 'columns'.")
                return None
            return self.active_df.head()
        except Exception as e:
            self.output_handler.show_error(f"Error deleting {target_type}: {e}")
            return None

    @tool(description="Creates a pivot table from the active DataFrame. Use this for 'create pivot table', 'summarize data by rows and columns'.")
    def pivot_table(self, index_column: str, columns_column: str, values_column: str, aggregation_type: str = 'sum') -> pd.DataFrame:
        """
        Creates a pivot table from the active DataFrame.
        Aggregation types: 'sum', 'mean', 'count', 'min', 'max', 'std'.
        """
        if self.active_df is None:
            self.output_handler.show_error("No active DataFrame. Please load data first.")
            return None
        if not all(col in self.active_df.columns for col in [index_column, columns_column, values_column]):
            self.output_handler.show_error(f"One or more specified columns not found in DataFrame: {index_column}, {columns_column}, {values_column}")
            return None
        
        valid_aggregations = ['sum', 'mean', 'count', 'min', 'max', 'std']
        if aggregation_type not in valid_aggregations:
            self.output_handler.show_error(f"Invalid aggregation type: '{aggregation_type}'. Must be one of {valid_aggregations}.")
            return None

        try:
            pivot_df = pd.pivot_table(self.active_df, values=values_column, 
                                      index=index_column, columns=columns_column, 
                                      aggfunc=aggregation_type)
            self.active_df = pivot_df.copy() # MODIFIED: Update active_df
            self.output_handler.show_success(f"Pivot table created with index '{index_column}', columns '{columns_column}', values '{values_column}', aggregated by '{aggregation_type}'.")
            return self.active_df
        except Exception as e:
            self.output_handler.show_error(f"Error creating pivot table: {e}")
            return None

    @tool(description="Displays the first or last N rows of the active DataFrame. Use this for 'show top N', 'show bottom N', 'preview data'.")
    def display_head_or_tail(self, num_rows: int = 5, from_end: bool = False) -> pd.DataFrame:
        """
        Displays the head or tail of the active DataFrame.
        """
        if self.active_df is None:
            self.output_handler.show_error("No active DataFrame. Please load data first.")
            return None
        try:
            if from_end:
                self.output_handler.show_success(f"Displaying last {num_rows} rows:")
                return self.active_df.tail(num_rows)
            else:
                self.output_handler.show_success(f"Displaying first {num_rows} rows:")
                return self.active_df.head(num_rows)
        except Exception as e:
            self.output_handler.show_error(f"Error displaying head/tail: {e}")
            return None

    @tool(description="Compares multiple calculated values from the active DataFrame and presents them. Use this for 'compare X to Y', 'what is the difference between'.")
    def compare_values(self, comparisons: List[Dict[str, Any]]) -> pd.DataFrame:
        """
        Compares multiple calculated values based on provided specifications.
        Each comparison dict needs 'label', 'column', 'aggregation_type', and optional 'query_string'.
        """
        if self.active_df is None:
            self.output_handler.show_error("No active DataFrame. Please load data first.")
            return None
        
        results = []
        for comp in comparisons:
            label = comp.get('label')
            column = comp.get('column')
            agg_type = comp.get('aggregation_type')
            query_string = comp.get('query_string')

            if not all([label, column, agg_type]):
                self.output_handler.show_warning(f"Skipping comparison due to missing required parameters: {comp}")
                continue

            try:
                value = self.calculate_scalar_value(column, agg_type, query_string)
                if value is not None:
                    results.append({'Comparison': label, 'Value': value})
            except Exception as e:
                self.output_handler.show_error(f"Error calculating value for comparison '{label}': {e}")
        
        if results:
            comparison_df = pd.DataFrame(results)
            self.output_handler.show_success("Comparison results:")
            return comparison_df
        else:
            self.output_handler.show_warning("No valid comparisons could be performed.")
            return None

    @tool(description="Extracts a specific part (year, month, day, quarter) from a date column and creates a new column. Use this for 'get year from date', 'extract month'.")
    def extract_date_part(self, date_column: str, part: str, new_column_name: str) -> pd.DataFrame:
        """
        Extracts a specific part (year, month, day, quarter) from a date column.
        """
        if self.active_df is None:
            self.output_handler.show_error("No active DataFrame. Please load data first.")
            return None
        if date_column not in self.active_df.columns:
            self.output_handler.show_error(f"Date column '{date_column}' not found in DataFrame.")
            return None
        
        try:
            # Ensure the column is datetime type
            self.active_df[date_column] = pd.to_datetime(self.active_df[date_column], errors='coerce')
            if self.active_df[date_column].isnull().all():
                self.output_handler.show_error(f"Column '{date_column}' could not be converted to datetime. Check its format.")
                return None

            if part == 'year':
                self.active_df[new_column_name] = self.active_df[date_column].dt.year
            elif part == 'month':
                self.active_df[new_column_name] = self.active_df[date_column].dt.month
            elif part == 'day':
                self.active_df[new_column_name] = self.active_df[date_column].dt.day
            elif part == 'quarter':
                self.active_df[new_column_name] = self.active_df[date_column].dt.quarter
            else:
                self.output_handler.show_error(f"Invalid date part '{part}'. Supported: 'year', 'month', 'day', 'quarter'.")
                return None
            
            self.output_handler.show_success(f"Extracted '{part}' from '{date_column}' into new column '{new_column_name}'. Displaying head:")
            return self.active_df.head()
        except Exception as e:
            self.output_handler.show_error(f"Error extracting date part from column '{date_column}': {e}")
            return None

    @tool(description="Adds a new column with lagged (previous period) values of an existing column, optionally grouped by other columns. Useful for time-series analysis like 'month-over-month change'.")
    def add_lagged_column(self, column: str, new_column_name: str, periods: int = 1, group_by_columns: List[str] = None) -> pd.DataFrame:
        """
        Adds a new column with lagged values of an existing column.
        Optionally groups by specified columns before applying the lag.
        """
        if self.active_df is None:
            self.output_handler.show_error("No active DataFrame. Please load data first.")
            return None
        if column not in self.active_df.columns:
            self.output_handler.show_error(f"Column '{column}' not found in DataFrame.")
            return None
        if group_by_columns and not all(col in self.active_df.columns for col in group_by_columns):
            self.output_handler.show_error(f"One or more group-by columns not found in DataFrame: {group_by_columns}")
            return None
        
        try:
            if group_by_columns:
                self.active_df[new_column_name] = self.active_df.groupby(group_by_columns)[column].shift(periods=periods)
                self.output_handler.show_success(f"Lagged column '{new_column_name}' added for '{column}', grouped by {group_by_columns}. Displaying head:")
            else:
                self.active_df[new_column_name] = self.active_df[column].shift(periods=periods)
                self.output_handler.show_success(f"Lagged column '{new_column_name}' added for '{column}'. Displaying head:")
            return self.active_df.head()
        except Exception as e:
            self.output_handler.show_error(f"Error adding lagged column '{new_column_name}': {e}")
            return None

    @tool(description="Generates various types of plots (line, bar, scatter, hist, box, pie) from the active DataFrame and saves it as an image file. Use this for 'plot X vs Y', 'show distribution of', 'create chart'.")
    def plot_dataframe(self, plot_type: str, output_filename: str, x_column: str = None, y_column: str = None, title: str = None, x_label: str = None, y_label: str = None, hue_column: str = None) -> str:
        """
        Generates a plot from the active DataFrame and saves it as an image file.
        Supported plot types: 'line', 'bar', 'scatter', 'hist', 'box', 'pie'.
        """
        if self.active_df is None:
            self.output_handler.show_error("No active DataFrame. Please load data first.")
            return None

        # Ensure output directory exists
        plots_dir = Config.PLOTS_DIR
        os.makedirs(plots_dir, exist_ok=True)
        
        # Construct the full output path relative to the current working directory
        # The agent sets CWD to the Excel file's directory, so plots_dir will be relative to that.
        full_output_path = os.path.join(plots_dir, output_filename)

        plt.figure(figsize=(10, 6))
        try:
            if plot_type == 'line':
                if x_column not in self.active_df.columns or y_column not in self.active_df.columns:
                    self.output_handler.show_error(f"For line plot, 'x_column' ('{x_column}') and 'y_column' ('{y_column}') must be present in DataFrame.")
                    return None
                sns.lineplot(x=self.active_df[x_column], y=self.active_df[y_column], hue=self.active_df[hue_column] if hue_column else None, data=self.active_df)
            elif plot_type == 'bar':
                if x_column not in self.active_df.columns or y_column not in self.active_df.columns:
                    self.output_handler.show_error(f"For bar plot, 'x_column' ('{x_column}') and 'y_column' ('{y_column}') must be present in DataFrame.")
                    return None
                sns.barplot(x=self.active_df[x_column], y=self.active_df[y_column], hue=self.active_df[hue_column] if hue_column else None, data=self.active_df)
            elif plot_type == 'scatter':
                if x_column not in self.active_df.columns or y_column not in self.active_df.columns:
                    self.output_handler.show_error(f"For scatter plot, 'x_column' ('{x_column}') and 'y_column' ('{y_column}') must be present in DataFrame.")
                    return None
                sns.scatterplot(x=self.active_df[x_column], y=self.active_df[y_column], hue=self.active_df[hue_column] if hue_column else None, data=self.active_df)
            elif plot_type == 'hist':
                if x_column not in self.active_df.columns:
                    self.output_handler.show_error(f"For histogram, 'x_column' ('{x_column}') must be present in DataFrame.")
                    return None
                sns.histplot(x=self.active_df[x_column], kde=True)
            elif plot_type == 'box':
                if x_column not in self.active_df.columns or y_column not in self.active_df.columns:
                    self.output_handler.show_error(f"For box plot, 'x_column' ('{x_column}') and 'y_column' ('{y_column}') must be present in DataFrame.")
                    return None
                sns.boxplot(x=self.active_df[x_column], y=self.active_df[y_column], data=self.active_df)
            elif plot_type == 'pie':
                if x_column not in self.active_df.columns or y_column not in self.active_df.columns:
                    self.output_handler.show_error(f"For pie chart, 'x_column' (labels: '{x_column}') and 'y_column' (values: '{y_column}') must be present in DataFrame.")
                    return None
                # Ensure numeric type for y_column for pie chart values
                if not pd.api.types.is_numeric_dtype(self.active_df[y_column]):
                    self.output_handler.show_error(f"Y-column '{y_column}' for pie chart must be numeric. Current type: {self.active_df[y_column].dtype}")
                    return None
                
                plt.pie(self.active_df[y_column], labels=self.active_df[x_column], autopct='%1.1f%%', startangle=90)
                plt.axis('equal') # Equal aspect ratio ensures that pie is drawn as a circle.
            else:
                self.output_handler.show_error(f"Unsupported plot type: '{plot_type}'.")
                return None

            plt.title(title or f"{plot_type.capitalize()} Plot")
            plt.xlabel(x_label or x_column)
            plt.ylabel(y_label or y_column)
            plt.tight_layout()
            plt.savefig(full_output_path)
            plt.close()
            self.output_handler.show_success(f"Plot saved to: {full_output_path}")
            return full_output_path # Return the path relative to CWD for agent to handle
        except Exception as e:
            self.output_handler.show_error(f"Error generating {plot_type} plot: {e}")
            plt.close()
            return None

    @tool(description="Generates a radar chart from the active DataFrame, showing average metrics across categories. Use this for 'radar chart of metrics by category'.")
    def plot_radar_chart(self, category_column: str, value_columns: List[str], output_filename: str, title: str = None) -> str:
        """
        Generates a radar chart showing average metrics across categories.
        Automatically calculates the mean of value_columns for each category.
        """
        if self.active_df is None:
            self.output_handler.show_error("No active DataFrame. Please load data first.")
            return None
        if category_column not in self.active_df.columns:
            self.output_handler.show_error(f"Category column '{category_column}' not found in DataFrame.")
            return None
        if not all(col in self.active_df.columns for col in value_columns):
            self.output_handler.show_error(f"One or more value columns not found in DataFrame: {value_columns}")
            return None

        plots_dir = Config.PLOTS_DIR
        os.makedirs(plots_dir, exist_ok=True)
        full_output_path = os.path.join(plots_dir, output_filename)

        try:
            # Calculate mean for each value column grouped by category
            df_radar = self.active_df.groupby(category_column)[value_columns].mean().reset_index()

            # Normalize data for radar chart (0 to 1 scale)
            df_normalized = df_radar.copy()
            for col in value_columns:
                min_val = df_normalized[col].min()
                max_val = df_normalized[col].max()
                if max_val - min_val > 0:
                    df_normalized[col] = (df_normalized[col] - min_val) / (max_val - min_val)
                else:
                    df_normalized[col] = 0.5 # If all values are same, put in middle

            # Plotting
            num_vars = len(value_columns)
            angles = np.linspace(0, 2 * np.pi, num_vars, endpoint=False).tolist()
            angles += angles[:1] # Complete the loop

            fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(polar=True))
            
            for i, row in df_normalized.iterrows():
                values = row[value_columns].tolist()
                values += values[:1]
                ax.plot(angles, values, label=row[category_column])
                ax.fill(angles, values, alpha=0.25)

            ax.set_theta_offset(np.pi / 2)
            ax.set_theta_direction(-1)
            ax.set_rlabel_position(0)
            plt.xticks(angles[:-1], value_columns)
            plt.yticks([0.2, 0.4, 0.6, 0.8, 1.0], ['20%', '40%', '60%', '80%', '100%'], color="grey", size=8)
            plt.ylim(0, 1)
            plt.title(title or f"Radar Chart of Average {', '.join(value_columns)} by {category_column}", size=15, color='blue', y=1.1)
            plt.legend(loc='upper right', bbox_to_anchor=(0.1, 0.1))
            plt.tight_layout()
            plt.savefig(full_output_path)
            plt.close()
            self.output_handler.show_success(f"Radar chart saved to: {full_output_path}")
            return full_output_path
        except Exception as e:
            self.output_handler.show_error(f"Error generating radar chart: {e}")
            plt.close()
            return None

    @tool(description="Converts the data type of a specified column in the active DataFrame. Use this for 'change column to numeric', 'convert to date', 'make text column'.")
    def convert_column_type(self, column: str, target_type: str) -> pd.DataFrame:
        """
        Converts a column to a specified data type ('numeric', 'datetime', 'string').
        """
        if self.active_df is None:
            self.output_handler.show_error("No active DataFrame. Please load data first.")
            return None
        if column not in self.active_df.columns:
            self.output_handler.show_error(f"Column '{column}' not found in DataFrame.")
            return None
        
        try:
            if target_type == 'numeric':
                self.active_df[column] = pd.to_numeric(self.active_df[column], errors='coerce')
            elif target_type == 'datetime':
                self.active_df[column] = pd.to_datetime(self.active_df[column], errors='coerce')
            elif target_type == 'string':
                self.active_df[column] = self.active_df[column].astype(str)
            else:
                self.output_handler.show_error(f"Invalid target_type: '{target_type}'. Supported: 'numeric', 'datetime', 'string'.")
                return None
            
            if self.active_df[column].isnull().any():
                self.output_handler.show_warning(f"Some values in column '{column}' could not be converted to '{target_type}' and were set to NaN.")
            
            self.output_handler.show_success(f"Column '{column}' converted to '{target_type}'. Displaying head:")
            return self.active_df.head()
        except Exception as e:
            self.output_handler.show_error(f"Error converting column '{column}' to '{target_type}': {e}")
            return None

    @tool(description="Splits a single text column into multiple new columns based on a delimiter. Use this for 'split address', 'separate names'.")
    def split_column_by_delimiter(self, column: str, delimiter: str, new_column_names: List[str]) -> pd.DataFrame:
        """
        Splits a text column into multiple new columns based on a delimiter.
        """
        if self.active_df is None:
            self.output_handler.show_error("No active DataFrame. Please load data first.")
            return None
        if column not in self.active_df.columns:
            self.output_handler.show_error(f"Column '{column}' not found in DataFrame.")
            return None
        
        try:
            # Ensure the column is string type
            self.active_df[column] = self.active_df[column].astype(str)
            
            split_data = self.active_df[column].str.split(delimiter, expand=True)
            
            if split_data.shape[1] != len(new_column_names):
                self.output_handler.show_warning(f"Number of new columns ({len(new_column_names)}) does not match the number of parts after splitting ({split_data.shape[1]}). Some new columns might be empty or data truncated.")

            for i, new_col_name in enumerate(new_column_names):
                if i < split_data.shape[1]:
                    self.active_df[new_col_name] = split_data[i]
                else:
                    self.active_df[new_col_name] = np.nan # Fill with NaN if no corresponding split part
            
            self.output_handler.show_success(f"Column '{column}' split by '{delimiter}' into new columns: {new_column_names}. Displaying head:")
            return self.active_df.head()
        except Exception as e:
            self.output_handler.show_error(f"Error splitting column '{column}': {e}")
            return None

    @tool(description="Extracts specific patterns (e.g., numbers, emails) from a text column using regular expressions and creates a new column. Use this for 'extract ID from text', 'find emails'.")
    def extract_pattern_from_column(self, column: str, regex_pattern: str, new_column_name: str, group_index: int = 0) -> pd.DataFrame:
        """
        Extracts a specific pattern from a text column using regular expressions.
        """
        if self.active_df is None:
            self.output_handler.show_error("No active DataFrame. Please load data first.")
            return None
        if column not in self.active_df.columns:
            self.output_handler.show_error(f"Column '{column}' not found in DataFrame.")
            return None
        
        try:
            self.active_df[column] = self.active_df[column].astype(str)
            extracted_data = self.active_df[column].str.extract(regex_pattern)
            
            if extracted_data.shape[1] > group_index:
                self.active_df[new_column_name] = extracted_data[group_index]
                self.output_handler.show_success(f"Pattern '{regex_pattern}' extracted from '{column}' into new column '{new_column_name}'. Displaying head:")
            else:
                self.output_handler.show_error(f"Regex pattern '{regex_pattern}' did not yield a group at index {group_index}. No data extracted.")
                return None
            
            return self.active_df.head()
        except Exception as e:
            self.output_handler.show_error(f"Error extracting pattern from column '{column}': {e}")
            return None

    @tool(description="Cleans a text column by applying operations like stripping whitespace, changing case, or removing digits/punctuation. Use this for 'clean text', 'standardize names'.")
    def clean_text_column(self, column: str, cleaning_operations: List[str]) -> pd.DataFrame:
        """
        Cleans a text column by applying specified operations.
        Operations: 'strip', 'lower', 'upper', 'remove_digits', 'remove_punctuation'.
        """
        if self.active_df is None:
            self.output_handler.show_error("No active DataFrame. Please load data first.")
            return None
        if column not in self.active_df.columns:
            self.output_handler.show_error(f"Column '{column}' not found in DataFrame.")
            return None
        
        try:
            cleaned_series = self.active_df[column].astype(str)
            
            for op in cleaning_operations:
                if op == 'strip':
                    cleaned_series = cleaned_series.str.strip()
                elif op == 'lower':
                    cleaned_series = cleaned_series.str.lower()
                elif op == 'upper':
                    cleaned_series = cleaned_series.str.upper()
                elif op == 'remove_digits':
                    cleaned_series = cleaned_series.str.replace(r'\d+', '', regex=True)
                elif op == 'remove_punctuation':
                    cleaned_series = cleaned_series.str.replace(r'[^\w\s]', '', regex=True)
                else:
                    self.output_handler.show_warning(f"Unsupported cleaning operation: '{op}'. Skipping.")
            
            self.active_df[column] = cleaned_series
            self.output_handler.show_success(f"Column '{column}' cleaned with operations: {cleaning_operations}. Displaying head:")
            return self.active_df.head()
        except Exception as e:
            self.output_handler.show_error(f"Error cleaning column '{column}': {e}")
            return None

    @tool(description="Performs a lookup (like VLOOKUP) to add columns from another Excel file/sheet to the active DataFrame based on a matching column. Use this for 'add data from another file', 'lookup values'.")
    def perform_lookup(self, lookup_file_path: str, lookup_sheet_name: str, on_column_main_df: str, on_column_lookup_df: str, columns_to_add: List[str], how: str = 'left') -> pd.DataFrame:
        """
        Performs a lookup operation to add columns from another Excel file/sheet.
        """
        if self.active_df is None:
            self.output_handler.show_error("No active DataFrame. Please load data first.")
            return None
        
        if on_column_main_df not in self.active_df.columns:
            self.output_handler.show_error(f"Matching column '{on_column_main_df}' not found in the active DataFrame.")
            return None

        lookup_df = self._load_data_internal(lookup_file_path, lookup_sheet_name)
        if lookup_df is None:
            self.output_handler.show_error(f"Could not load lookup data from '{lookup_file_path}' sheet '{lookup_sheet_name}'.")
            return None
        
        if on_column_lookup_df not in lookup_df.columns:
            self.output_handler.show_error(f"Matching column '{on_column_lookup_df}' not found in the lookup DataFrame.")
            return None
        if not all(col in lookup_df.columns for col in columns_to_add):
            self.output_handler.show_error(f"One or more columns to add ({columns_to_add}) not found in the lookup DataFrame.")
            return None

        try:
            merged_df = pd.merge(self.active_df, lookup_df[[on_column_lookup_df] + columns_to_add], 
                                 left_on=on_column_main_df, right_on=on_column_lookup_df, how=how)
            
            # Drop the duplicate 'on_column_lookup_df' if it's different from 'on_column_main_df'
            if on_column_main_df != on_column_lookup_df and on_column_lookup_df in merged_df.columns:
                merged_df.drop(columns=[on_column_lookup_df], inplace=True)

            self.active_df = merged_df # Update the active DataFrame
            self.output_handler.show_success(f"Columns {columns_to_add} added from '{lookup_file_path}' sheet '{lookup_sheet_name}' via lookup. Displaying head:")
            return self.active_df.head()
        except Exception as e:
            self.output_handler.show_error(f"Error performing lookup: {e}")
            return None

    @tool(description="Fills missing values in a column using advanced strategies like forward-fill, backward-fill, or interpolation. Use this for 'fill NaNs with previous/next value', 'interpolate missing data'.")
    def impute_missing_values_advanced(self, column: str, strategy: str, limit: int = None) -> pd.DataFrame:
        """
        Fills missing values in a column using advanced strategies.
        Strategies: 'ffill' (forward-fill), 'bfill' (backward-fill), 'interpolate'.
        """
        if self.active_df is None:
            self.output_handler.show_error("No active DataFrame. Please load data first.")
            return None
        if column not in self.active_df.columns:
            self.output_handler.show_error(f"Column '{column}' not found in DataFrame.")
            return None
        
        original_nan_count = self.active_df[column].isnull().sum()
        if original_nan_count == 0:
            self.output_handler.show_warning(f"No missing values found in column '{column}'. No action taken.")
            return self.active_df.head()

        try:
            if strategy == 'ffill':
                self.active_df[column].fillna(method='ffill', limit=limit, inplace=True)
                self.output_handler.show_success(f"Missing values in '{column}' forward-filled (limit={limit}).")
            elif strategy == 'bfill':
                self.active_df[column].fillna(method='bfill', limit=limit, inplace=True)
                self.output_handler.show_success(f"Missing values in '{column}' backward-filled (limit={limit}).")
            elif strategy == 'interpolate':
                if pd.api.types.is_numeric_dtype(self.active_df[column]):
                    self.active_df[column].interpolate(method='linear', limit=limit, inplace=True)
                    self.output_handler.show_success(f"Missing values in '{column}' interpolated (limit={limit}).")
                else:
                    self.output_handler.show_error(f"Cannot interpolate: Column '{column}' is not numeric.")
                    return None
            else:
                self.output_handler.show_error(f"Invalid strategy: '{strategy}'. Supported: 'ffill', 'bfill', 'interpolate'.")
                return None
            
            return self.active_df.head()
        except Exception as e:
            self.output_handler.show_error(f"Error imputing missing values in column '{column}' with strategy '{strategy}': {e}")
            return None

    @tool(description="Exports the active DataFrame to a new file in CSV, JSON, or Excel format. Use this when the user asks to 'save data', 'export data', 'create new file'.")
    def export_dataframe(self, output_file_path: str, output_format: str) -> str:
        """
        Exports the active DataFrame to a new file in specified format.
        Supported formats: 'csv', 'json', 'excel'.
        """
        if self.active_df is None:
            self.output_handler.show_error("No active DataFrame to export. Please load data first.")
            return None
        
        # Ensure the directory exists
        output_dir = os.path.dirname(output_file_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)

        try:
            if output_format == 'csv':
                self.active_df.to_csv(output_file_path, index=False)
            elif output_format == 'json':
                self.active_df.to_json(output_file_path, orient='records', indent=4)
            elif output_format == 'excel':
                self.active_df.to_excel(output_file_path, index=False)
            else:
                self.output_handler.show_error(f"Unsupported export format: '{output_format}'. Supported: 'csv', 'json', 'excel'.")
                return None
            
            self.output_handler.show_success(f"DataFrame successfully exported to '{output_file_path}' as {output_format}.")
            return output_file_path
        except Exception as e:
            self.output_handler.show_error(f"Error exporting DataFrame to '{output_file_path}': {e}")
            return None

