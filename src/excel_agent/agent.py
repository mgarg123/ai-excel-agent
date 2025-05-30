import pandas as pd
import json
import os
from typing import List, Dict, Any

from src.excel_agent.llm_interface import LLMInterface
from src.excel_agent.excel_handler import ExcelHandler
from src.excel_agent.config import Config
from src.excel_agent.utils import generate_output_filename
from src.excel_agent.tools import tool
from src.excel_agent.output.abstract_output_handler import AbstractOutputHandler

class ExcelAgent:
    """
    The core agent orchestrator.
    It takes user queries, interacts with the LLM, and executes Excel operations via tool calls.
    """
    def __init__(self, output_handler: AbstractOutputHandler):
        self.llm = LLMInterface(output_handler)
        self.output_handler = output_handler
        self.excel_handlers: Dict[str, ExcelHandler] = {}
        self.active_file_path: str = None
        self.active_sheet_name: str = None

        # Map tool names (from LLM) to methods.
        # Tools that operate on a single file/sheet will map to ExcelHandler methods.
        # Tools that operate across files (like merge) will map to ExcelAgent methods.
        self.tool_map = {
            "load_and_display_data": ExcelHandler.load_and_display_data,
            "filter_and_display_dataframe": ExcelHandler.filter_and_display_dataframe,
            "group_and_display_dataframe": ExcelHandler.group_and_display_dataframe,
            "sort_and_display_dataframe": ExcelHandler.sort_and_display_dataframe,
            "add_column_and_display_dataframe": ExcelHandler.add_column_and_display_dataframe,
            "calculate_scalar_value": ExcelHandler.calculate_scalar_value,
            "save_dataframe_to_new_excel": ExcelHandler.save_dataframe_to_new_excel,
            "apply_excel_formula": ExcelHandler.apply_excel_formula,
            "apply_formatting": ExcelHandler.apply_formatting,
            "handle_missing_values": ExcelHandler.handle_missing_values,
            "remove_duplicates": ExcelHandler.remove_duplicates,
            "rename_column": ExcelHandler.rename_column,
            "select_columns_and_display": ExcelHandler.select_columns_and_display,
            "get_descriptive_statistics": ExcelHandler.get_descriptive_statistics,
            "delete_rows_or_columns": ExcelHandler.delete_rows_or_columns,
            "pivot_table": ExcelHandler.pivot_table,
            "display_head_or_tail": ExcelHandler.display_head_or_tail,
            "compare_values": ExcelHandler.compare_values,
            "extract_date_part": ExcelHandler.extract_date_part,
            "add_lagged_column": ExcelHandler.add_lagged_column,
            "plot_dataframe": ExcelHandler.plot_dataframe,
            "plot_radar_chart": ExcelHandler.plot_radar_chart,
            "convert_column_type": ExcelHandler.convert_column_type,
            "split_column_by_delimiter": ExcelHandler.split_column_by_delimiter,
            "extract_pattern_from_column": ExcelHandler.extract_pattern_from_column,
            "clean_text_column": ExcelHandler.clean_text_column,
            "perform_lookup": ExcelHandler.perform_lookup,
            "impute_missing_values_advanced": ExcelHandler.impute_missing_values_advanced,
            "export_dataframe": ExcelHandler.export_dataframe,
            "merge_dataframes": self.merge_dataframes,
            "concatenate_dataframes": self.concatenate_dataframes,
        }

    @tool(description="Merges two DataFrames from different Excel files or sheets based on a common key. Use this when the user asks to 'combine data from two files/sheets' or 'join sheets'.")
    def merge_dataframes(self, file_path_left: str, sheet_name_left: str, file_path_right: str, sheet_name_right: str, on_column: str, how: str = 'inner') -> pd.DataFrame:
        """
        Merges two DataFrames from different Excel files/sheets based on a common column.
        'how' can be 'inner', 'left', 'right', 'outer'.
        """
        df_left = None
        df_right = None

        # Ensure handlers for both files exist
        if file_path_left not in self.excel_handlers:
            self.output_handler.show_error(f"Excel file '{file_path_left}' was not loaded. Please ensure it's provided as input.")
            return None
        if file_path_right not in self.excel_handlers:
            self.output_handler.show_error(f"Excel file '{file_path_right}' was not loaded. Please ensure it's provided as input.")
            return None

        # Load data using the respective handlers' internal method (not setting active_df for merge sources)
        df_left = self.excel_handlers[file_path_left]._load_data_internal(file_path_left, sheet_name=sheet_name_left)
        df_right = self.excel_handlers[file_path_right]._load_data_internal(file_path_right, sheet_name=sheet_name_right)

        if df_left is None or df_right is None:
            self.output_handler.show_error("Could not load one or both specified sheets for merging.")
            return None

        if on_column not in df_left.columns:
            self.output_handler.show_error(f"Merge key '{on_column}' not found in sheet '{sheet_name_left}' of file '{file_path_left}'.")
            return None
        if on_column not in df_right.columns:
            self.output_handler.show_error(f"Merge key '{on_column}' not found in sheet '{sheet_name_right}' of file '{file_path_right}'.")
            return None
        
        if how not in ['inner', 'left', 'right', 'outer']:
            self.output_handler.show_error(f"Invalid merge 'how' parameter: '{how}'. Must be 'inner', 'left', 'right', or 'outer'.")
            return None

        try:
            merged_df = pd.merge(df_left, df_right, on=on_column, how=how)
            self.output_handler.show_success(f"Sheets '{sheet_name_left}' from '{file_path_left}' and '{sheet_name_right}' from '{file_path_right}' merged on column '{on_column}' using '{how}' join.")
            
            return merged_df
        except Exception as e:
            self.output_handler.show_error(f"Error merging dataframes: {e}")
            return None

    @tool(description="Concatenates (stacks) two DataFrames vertically from different Excel files or sheets. Use this when the user asks to 'combine rows from two files/sheets' or 'stack data'.")
    def concatenate_dataframes(self, file_path_top: str, sheet_name_top: str, file_path_bottom: str, sheet_name_bottom: str) -> pd.DataFrame:
        """
        Concatenates two DataFrames vertically (stacks rows) from different Excel files/sheets.
        The DataFrames should ideally have the same column structure for meaningful concatenation.
        """
        df_top = None
        df_bottom = None

        # Ensure handlers for both files exist
        if file_path_top not in self.excel_handlers:
            self.output_handler.show_error(f"Excel file '{file_path_top}' was not loaded. Please ensure it's provided as input.")
            return None
        if file_path_bottom not in self.excel_handlers:
            self.output_handler.show_error(f"Excel file '{file_path_bottom}' was not loaded. Please ensure it's provided as input.")
            return None

        # Load data using the respective handlers' internal method
        df_top = self.excel_handlers[file_path_top]._load_data_internal(file_path_top, sheet_name=sheet_name_top)
        df_bottom = self.excel_handlers[file_path_bottom]._load_data_internal(file_path_bottom, sheet_name=sheet_name_bottom)

        if df_top is None or df_bottom is None:
            self.output_handler.show_error("Could not load one or both specified sheets for concatenation.")
            return None

        try:
            concatenated_df = pd.concat([df_top, df_bottom], ignore_index=True)
            self.output_handler.show_success(f"Sheets '{sheet_name_top}' from '{file_path_top}' and '{sheet_name_bottom}' from '{file_path_bottom}' concatenated vertically.")
            
            # Set the concatenated_df as the active_df for one of the handlers
            # For simplicity, let's set it on the handler for file_path_top
            self.excel_handlers[file_path_top].active_df = concatenated_df.copy()
            return concatenated_df
        except Exception as e:
            self.output_handler.show_error(f"Error concatenating dataframes: {e}")
            return None

    def process_query(self, file_paths: List[str], user_query: str, show_all_tool_results: bool = False):
        """
        Processes a natural language query to manipulate Excel files using tool calls.
        'show_all_tool_results': If True, displays output after each tool execution.
                                 If False, only displays the final result of the last tool.
        """
        # Initialize ExcelHandler for each file and gather context
        file_contexts = []
        for f_path in file_paths:
            handler = ExcelHandler(f_path, self.output_handler)
            self.excel_handlers[f_path] = handler # Store handler instance

            all_sheet_names = handler.get_sheet_names()
            if not all_sheet_names:
                self.output_handler.show_error(f"Could not read sheet names from Excel file: '{f_path}'. Please ensure it's a valid .xlsx or .xls file.")
                continue # Skip this file but continue with others

            file_context = {
                "file_path": f_path,
                "sheets": []
            }
            for s_name in all_sheet_names:
                column_headers = handler.get_column_headers(sheet_name=s_name)
                if not column_headers:
                    self.output_handler.show_warning(f"Could not read column headers from sheet '{s_name}' in file '{f_path}'. It might be empty or malformed.")
                    # Still include the sheet name even if headers are empty, so LLM knows it exists
                    file_context["sheets"].append({"sheet_name": s_name, "column_headers": []})
                else:
                    file_context["sheets"].append({"sheet_name": s_name, "column_headers": column_headers})
            
            if file_context["sheets"]: # Only add if there's at least one valid sheet
                file_contexts.append(file_context)

        if not file_contexts:
            self.output_handler.show_error("No valid Excel files or sheets found to process.")
            return

        # 2. Construct LLM prompt with context from all files
        context_message_parts = [
            "You are an expert Excel assistant. Based on the user's query and the provided Excel file contexts, select and call the appropriate tool(s). "
            "For operations that modify or query a DataFrame (like filtering, grouping, sorting, calculating values, etc.), you must first explicitly call `load_and_display_data` to load a specific sheet from an Excel file. This loaded sheet will become the 'active' DataFrame for all subsequent operations until a new `load_and_display_data` call is made. "
            "When calling tools like `filter_and_display_dataframe`, `group_and_display_dataframe`, `display_head_or_tail`, etc., you do NOT need to provide `file_path` or `sheet_name` parameters, as they will automatically operate on the currently active DataFrame. "
            "If the user's query implies multiple chained operations (e.g., 'Calculate the average Profit for the East region' or 'Show the top 5 records with the highest Units Sold'), ensure you call the tools in the correct logical sequence (e.g., first filter, then calculate; or first sort, then display head).",
            "For `add_column_and_display_dataframe`, the `formula` parameter can now be any valid pandas expression involving existing column names. Column names with spaces MUST be enclosed in backticks (e.g., `Column Name`). Example: '`Net Revenue` * 0.10', '(`Profit` - `Previous Month Profit`) / `Previous Month Profit`'. This tool is NOT for functions like `MONTH()` or `LAG()`. Use `extract_date_part` and `add_lagged_column` for those specific needs.",
            "For `filter_and_display_dataframe`, the `query_string` parameter uses pandas.DataFrame.query() syntax. This allows for complex boolean logic (e.g., `and`, `or`, `not`). Column names with spaces MUST be enclosed in backticks (e.g., `Column Name`). String values MUST be enclosed in single quotes (e.g., `'Value'`). For example, to filter for 'Department' being 'Sales', the query string should be `\"Department == 'Sales'\"`.",
            "For `group_and_display_dataframe`, when you aggregate, the resulting column will be named predictably: if `aggregation_type` is 'count', the column will be named 'CountOfRecords'. Otherwise, it will be named as '{target_column}_{aggregation_type}' (e.g., 'Revenue_sum', 'Profit_mean'). You must use these exact names in subsequent tool calls, especially for plotting.", # MODIFIED: Added naming convention for grouped columns
            "For `calculate_scalar_value` and `compare_values`, the `aggregation_type` parameter now supports 'sum', 'mean', 'count', 'min', 'max', and 'std' (standard deviation).",
            "Crucially, `calculate_scalar_value` now accepts an optional `query_string` parameter. Use this when you need to calculate a statistic (like mean or std) on a *subset* of the data without permanently changing the active DataFrame. For example, to get the 'average Expenses in the Sales department', you would call `calculate_scalar_value` with `column='Expenses'`, `aggregation_type='mean'`, and `query_string=\"Department == 'Sales'\"`. The active DataFrame will remain the full dataset after this calculation.",
            "When you calculate a scalar value (e.g., mean, std) using `calculate_scalar_value` and need to use it in a subsequent `filter_and_display_dataframe` tool call, you must embed the *exact numerical result* into the `query_string`. For example, if `calculate_scalar_value` returns 123.45 for the mean of 'Units Sold', your next `filter_and_display_dataframe` call should use `query_string=\"`Units Sold` > 123.45\"`. Do NOT use placeholders like `mean_value` or `MEAN_Units_Sold` in the `query_string` you generate; directly insert the number.",
            "For `compare_values`, provide a list of dictionaries, each describing a value to calculate (with a label, column, aggregation type, and optional query string). This tool will perform all calculations and present a consolidated comparison.",
            "For `plot_dataframe`, you can generate various charts like 'line', 'bar', 'scatter', 'hist', 'box', or 'pie'. You must provide an `output_filename` (e.g., 'my_chart.png'). Specify `x_column` and `y_column` for most plots, and optionally `hue_column` for grouping/coloring. For 'pie' charts, `x_column` should be the categorical column for labels and `y_column` should be the numeric column for values. The plot will be saved as an image file in the 'plots' directory. Note: For radar charts, use `plot_radar_chart` instead. When a parameter is optional and not explicitly requested by the user (e.g., `hue_column` if no grouping is asked), you MUST omit that parameter from the tool_parameters dictionary entirely, rather than setting its value to 'null' or None.",
            "For `plot_radar_chart`, this tool will automatically calculate the *mean* of the `value_columns` for each `category_column` from the currently active DataFrame before plotting. It is ideal for comparing multiple quantitative variables (average metrics) across different categories. You must provide a `category_column` (e.g., 'Region') and a `value_columns` list (e.g., `['Revenue', 'Expenses', 'Profit']`). The plot will be saved as an image file.",
            "For `convert_column_type`, use this to change a column's data type to 'numeric', 'datetime', or 'string'. This is essential for correct calculations and analysis.",
            "For `split_column_by_delimiter`, use this to break a single text column into multiple new columns based on a delimiter (e.g., splitting 'Address' into 'Street', 'City', 'Zip').",
            "For `extract_pattern_from_column`, use this to pull out specific text patterns (like numbers, emails, or codes) from a column using regular expressions.",
            "For `clean_text_column`, use this to standardize text data by applying operations like stripping whitespace, changing case, or removing digits/punctuation.",
            "For `perform_lookup`, use this to add data from a separate Excel file/sheet to your active DataFrame, similar to VLOOKUP. You'll need to specify the lookup file, sheet, and the columns to match and add.",
            "For `impute_missing_values_advanced`, use this to fill missing values in a column using more sophisticated methods like forward-fill (`ffill`), backward-fill (`bfill`), or interpolation. You can also set a `limit` for consecutive fills.",
            "For `export_dataframe`, use this to save your processed data to different formats like CSV, JSON, or Excel. Use this when the user asks to 'save the data', 'export to a new file', or 'create a new Excel file'.",
            "For `concatenate_dataframes`, use this to stack rows from two different Excel sheets or files vertically. This is useful when you have data with the same structure but from different periods or sources.",
            "Example for chained operations: To 'Calculate the average Profit for the East region' from 'extended_excel_test_data.xlsx' sheet 'Sheet1':",
            "1. `{'tool_name': 'load_and_display_data', 'tool_parameters': {'file_path': 'extended_excel_test_data.xlsx', 'sheet_name': 'Sheet1'}}`",
            "2. `{'tool_name': 'filter_and_display_dataframe', 'tool_parameters': {'query_string': \"Region == 'East'\"}}`",
            "3. `{'tool_name': 'calculate_scalar_value', 'tool_parameters': {'column': 'Profit', 'aggregation_type': 'mean'}}`",
            "\nExample for 'Show the top 5 records with the highest Units Sold':",
            "1. `{'tool_name': 'load_and_display_data', 'tool_parameters': {'file_path': 'extended_excel_test_data.xlsx', 'sheet_name': 'Sheet1'}}`",
            "2. `{'tool_name': 'sort_and_display_dataframe', 'tool_parameters': {'sort_by_columns': ['Units Sold'], 'ascending': False}}`",
            "3. `{'tool_name': 'display_head_or_tail', 'tool_parameters': {'num_rows': 5, 'from_end': False}}`",
            "\nExample for 'Calculate the average Profit margin (Profit/Net Revenue) for each Region.':",
            "1. `{'tool_name': 'load_and_display_data', 'tool_parameters': {'file_path': 'extended_excel_test_data.xlsx', 'sheet_name': 'Sheet1'}}`",
            "2. `{'tool_name': 'add_column_and_display_dataframe', 'tool_parameters': {'new_column_name': 'Profit Margin', 'formula': 'Profit / `Net Revenue`'}}`",
            "3. `{'tool_name': 'group_and_display_dataframe', 'tool_parameters': {'group_by_columns': ['Region'], 'target_column': 'Profit Margin', 'aggregation_type': 'mean'}}`",
            "\nExample for 'Show entries where the Discount Amount is greater than 500 or Net Revenue is less than 2000.':",
            "1. `{'tool_name': 'load_and_display_data', 'tool_parameters': {'file_path': 'extended_excel_test_data.xlsx', 'sheet_name': 'Sheet1'}}`",
            "2. `{'tool_name': 'filter_and_display_dataframe', 'tool_parameters': {'query_string': \"(`Discount Amount` > 500) or (`Net Revenue` < 2000)\"}}`",
            "\nExample for 'List records where the Department is HR and either Revenue is above 3000 or Units Sold is below 20.':",
            "1. `{'tool_name': 'load_and_display_data', 'tool_parameters': {'file_path': 'extended_excel_test_data.xlsx', 'sheet_name': 'Sheet1'}}`",
            "2. `{'tool_name': 'filter_and_display_dataframe', 'tool_parameters': {'query_string': \"Department == 'HR' and (Revenue > 3000 or `Units Sold` < 20)\"}}`",
            "\nExample for 'Compare the total Revenue of Gadget X and Widget B.':",
            "1. `{'tool_name': 'load_and_display_data', 'tool_parameters': {'file_path': 'extended_excel_test_data.xlsx', 'sheet_name': 'Sheet1'}}`",
            "2. `{'tool_name': 'compare_values', 'tool_parameters': {'comparisons': ["
            "    {'label': 'Total Revenue of Gadget X', 'column': 'Revenue', 'aggregation_type': 'sum', 'query_string': \"Product == 'Gadget X'\"},"
            "    {'label': 'Total Revenue of Widget B', 'column': 'Revenue', 'aggregation_type': 'sum', 'query_string': \"Product == 'Widget B'\"}"
            "]}}`",
            "\nExample for 'List entries where Units Sold is more than two standard deviations above the average.':",
            "1. `{'tool_name': 'load_and_display_data', 'tool_parameters': {'file_path': 'extended_excel_test_data.xlsx', 'sheet_name': 'Sheet1'}}`",
            "2. `{'tool_name': 'calculate_scalar_value', 'tool_parameters': {'column': 'Units Sold', 'aggregation_type': 'mean'}}` (Assume this returns 100)",
            "3. `{'tool_name': 'calculate_scalar_value', 'tool_parameters': {'column': 'Units Sold', 'aggregation_type': 'std'}}` (Assume this returns 10)",
            "4. `{'tool_name': 'filter_and_display_dataframe', 'tool_parameters': {'query_string': \"`Units Sold` > (100 + 2 * 10)\"}}` (LLM must substitute the actual numbers from previous steps)",
            "\nExample for 'Show records where Expenses are unusually high compared to the average Expenses in the Sales department.':",
            "1. `{'tool_name': 'load_and_display_data', 'tool_parameters': {'file_path': 'extended_excel_test_data.xlsx', 'sheet_name': 'Sheet1'}}`",
            "2. `{'tool_name': 'calculate_scalar_value', 'tool_parameters': {'column': 'Expenses', 'aggregation_type': 'mean', 'query_string': \"Department == 'Sales'\"}}` (Assume this returns 500.50)",
            "3. `{'tool_name': 'calculate_scalar_value', 'tool_parameters': {'column': 'Expenses', 'aggregation_type': 'std', 'query_string': \"Department == 'Sales'\"}}` (Assume this returns 50.25)",
            "4. `{'tool_name': 'filter_and_display_dataframe', 'tool_parameters': {'query_string': \"`Expenses` > (500.50 + 2 * 50.25)\"}}` (LLM must substitute the actual numbers from previous steps)",
            "5. `{'tool_name': 'export_dataframe', 'tool_parameters': {'output_file_path': 'high_expenses_sales_department.xlsx', 'output_format': 'excel'}}`",
            "\nExample for 'What is the month-over-month change in Profit for the North region?':",
            "1. `{'tool_name': 'load_and_display_data', 'tool_parameters': {'file_path': 'extended_excel_test_data.xlsx', 'sheet_name': 'Sheet1'}}`",
            "2. `{'tool_name': 'filter_and_display_dataframe', 'tool_parameters': {'query_string': \"Region == 'North'\"}}`",
            "3. `{'tool_name': 'extract_date_part', 'tool_parameters': {'date_column': 'Date', 'part': 'year', 'new_column_name': 'Year'}}`",
            "4. `{'tool_name': 'extract_date_part', 'tool_parameters': {'date_column': 'Date', 'part': 'month', 'new_column_name': 'Month'}}`",
            "5. `{'tool_name': 'sort_and_display_dataframe', 'tool_parameters': {'sort_by_columns': ['Year', 'Month', 'Date'], 'ascending': True}}`",
            "6. `{'tool_name': 'add_lagged_column', 'tool_parameters': {'column': 'Profit', 'new_column_name': 'Previous Month Profit', 'periods': 1, 'group_by_columns': ['Region']}}`",
            "7. `{'tool_name': 'add_column_and_display_dataframe', 'tool_parameters': {'new_column_name': 'MoM Profit Change', 'formula': '(`Profit` - `Previous Month Profit`) / `Previous Month Profit`'}}`",
            "8. `{'tool_name': 'select_columns_and_display', 'tool_parameters': {'columns_to_select': ['Date', 'Region', 'Profit', 'Previous Month Profit', 'MoM Profit Change']}}`",
            "\nExample for 'Plot the total sales by product as a bar chart.':",
            "1. `{'tool_name': 'load_and_display_data', 'tool_parameters': {'file_path': 'extended_excel_test_data.xlsx', 'sheet_name': 'Sheet1'}}`",
            "2. `{'tool_name': 'group_and_display_dataframe', 'tool_parameters': {'group_by_columns': ['Product'], 'target_column': 'Sales', 'aggregation_type': 'sum'}}`",
            "3. `{'tool_name': 'plot_dataframe', 'tool_parameters': {'plot_type': 'bar', 'x_column': 'Product', 'y_column': 'Sales_sum', 'title': 'Total Sales by Product', 'output_filename': 'total_sales_by_product.png'}}`", # MODIFIED: y_column name
            "\nExample for 'Show a line chart of Profit over Date, separated by Region.':",
            "1. `{'tool_name': 'load_and_display_data', 'tool_parameters': {'file_path': 'extended_excel_test_data.xlsx', 'sheet_name': 'Sheet1'}}`",
            "2. `{'tool_name': 'sort_and_display_dataframe', 'tool_parameters': {'sort_by_columns': ['Date'], 'ascending': True}}`",
            "3. `{'tool_name': 'plot_dataframe', 'tool_parameters': {'plot_type': 'line', 'x_column': 'Date', 'y_column': 'Profit', 'hue_column': 'Region', 'title': 'Profit Over Time by Region', 'output_filename': 'profit_over_time_by_region.png'}}`",
            "\nExample for 'Create a histogram of Units Sold.':",
            "1. `{'tool_name': 'load_and_display_data', 'tool_parameters': {'file_path': 'extended_excel_test_data.xlsx', 'sheet_name': 'Sheet1'}}`",
            "2. `{'tool_name': 'plot_dataframe', 'tool_parameters': {'plot_type': 'hist', 'x_column': 'Units Sold', 'title': 'Distribution of Units Sold', 'output_filename': 'units_sold_histogram.png'}}`",
            "\nExample for 'Create a pie chart showing the distribution of total Profit by Department.':",
            "1. `{'tool_name': 'load_and_display_data', 'tool_parameters': {'file_path': 'extended_excel_test_data.xlsx', 'sheet_name': 'Sheet1'}}`",
            "2. `{'tool_name': 'group_and_display_dataframe', 'tool_parameters': {'group_by_columns': ['Department'], 'target_column': 'Profit', 'aggregation_type': 'sum'}}`",
            "3. `{'tool_name': 'plot_dataframe', 'tool_parameters': {'plot_type': 'pie', 'x_column': 'Department', 'y_column': 'Profit_sum', 'title': 'Distribution of Total Profit by Department', 'output_filename': 'profit_by_department_pie.png'}}`", # MODIFIED: y_column name
            "\nExample for 'Show a line chart of Revenue over time for the Sales department.':",
            "1. `{'tool_name': 'load_and_display_data', 'tool_parameters': {'file_path': 'extended_excel_test_data.xlsx', 'sheet_name': 'Sheet1'}}`",
            "2. `{'tool_name': 'filter_and_display_dataframe', 'tool_parameters': {'query_string': \"Department == 'Sales'\"}}`",
            "3. `{'tool_name': 'sort_and_display_dataframe', 'tool_parameters': {'sort_by_columns': ['Date'], 'ascending': True}}`",
            "4. `{'tool_name': 'plot_dataframe', 'tool_parameters': {'plot_type': 'line', 'x_column': 'Date', 'y_column': 'Revenue', 'title': 'Revenue Over Time for Sales Department', 'output_filename': 'revenue_over_time_sales.png'}}`",
            "\nExample for 'Create a pie chart showing the percentage of records with negative Profit by Department.':", # NEW EXAMPLE
            "1. `{'tool_name': 'load_and_display_data', 'tool_parameters': {'file_path': 'extended_excel_test_data.xlsx', 'sheet_name': 'Sheet1'}}`",
            "2. `{'tool_name': 'filter_and_display_dataframe', 'tool_parameters': {'query_string': \"Profit < 0\"}}`",
            "3. `{'tool_name': 'group_and_display_dataframe', 'tool_parameters': {'group_by_columns': ['Department'], 'target_column': 'Profit', 'aggregation_type': 'count'}}`", # This will create 'CountOfRecords'
            "4. `{'tool_name': 'plot_dataframe', 'tool_parameters': {'plot_type': 'pie', 'x_column': 'Department', 'y_column': 'CountOfRecords', 'title': 'Percentage of Negative Profit Records by Department', 'output_filename': 'negative_profit_by_department_pie.png'}}`", # Use 'CountOfRecords'
            "\nExample for 'Show a radar chart of average Revenue, Expenses, and Profit by Region.':",
            "1. `{'tool_name': 'load_and_display_data', 'tool_parameters': {'file_path': 'extended_excel_test_data.xlsx', 'sheet_name': 'Sheet1'}}`",
            "2. `{'tool_name': 'plot_radar_chart', 'tool_parameters': {'category_column': 'Region', 'value_columns': ['Revenue', 'Expenses', 'Profit'], 'title': 'Average Metrics by Region', 'output_filename': 'avg_metrics_by_region_radar.png'}}`",
            "\nExample for 'Convert the 'Order Date' column to datetime format.':",
            "1. `{'tool_name': 'load_and_display_data', 'tool_parameters': {'file_path': 'extended_excel_test_data.xlsx', 'sheet_name': 'Sheet1'}}`",
            "2. `{'tool_name': 'convert_column_type', 'tool_parameters': {'column': 'Order Date', 'target_type': 'datetime'}}`",
            "\nExample for 'Split the 'Full Name' column into 'First Name' and 'Last Name' using space as a delimiter.':",
            "1. `{'tool_name': 'load_and_display_data', 'tool_parameters': {'file_path': 'extended_excel_test_data.xlsx', 'sheet_name': 'Sheet1'}}`",
            "2. `{'tool_name': 'split_column_by_delimiter', 'tool_parameters': {'column': 'Full Name', 'delimiter': ' ', 'new_column_names': ['First Name', 'Last Name']}}`",
            "\nExample for 'Extract all numbers from the 'Product Code' column and put them into a new column called 'Product ID'.':",
            "1. `{'tool_name': 'load_and_display_data', 'tool_parameters': {'file_path': 'extended_excel_test_data.xlsx', 'sheet_name': 'Sheet1'}}`",
            "2. `{'tool_name': 'extract_pattern_from_column', 'tool_parameters': {'column': 'Product Code', 'regex_pattern': '\\\\d+', 'new_column_name': 'Product ID'}}`",
            "\nExample for 'Clean the 'Description' column by removing leading/trailing spaces and converting to lowercase.':",
            "1. `{'tool_name': 'load_and_display_data', 'tool_parameters': {'file_path': 'extended_excel_test_data.xlsx', 'sheet_name': 'Sheet1'}}`",
            "2. `{'tool_name': 'clean_text_column', 'tool_parameters': {'column': 'Description', 'cleaning_operations': ['strip', 'lower']}}`",
            "\nExample for 'Add 'Category' and 'Price' from 'product_details.xlsx' sheet 'Products' to the current data, matching on 'Product Name'.':",
            "1. `{'tool_name': 'load_and_display_data', 'tool_parameters': {'file_path': 'extended_excel_test_data.xlsx', 'sheet_name': 'Sheet1'}}`",
            "2. `{'tool_name': 'perform_lookup', 'tool_parameters': {'lookup_file_path': 'product_details.xlsx', 'lookup_sheet_name': 'Products', 'on_column_main_df': 'Product Name', 'on_column_lookup_df': 'Product Name', 'columns_to_add': ['Category', 'Price']}}`",
            "\nExample for 'Fill missing values in the 'Sales' column using the previous valid observation.':",
            "1. `{'tool_name': 'load_and_display_data', 'tool_parameters': {'file_path': 'extended_excel_test_data.xlsx', 'sheet_name': 'Sheet1'}}`",
            "2. `{'tool_name': 'impute_missing_values_advanced', 'tool_parameters': {'column': 'Sales', 'strategy': 'ffill'}}`",
            "\nExample for 'Export the current data to a CSV file named 'processed_data.csv'.':",
            "1. `{'tool_name': 'export_dataframe', 'tool_parameters': {'output_file_path': 'processed_data.csv', 'output_format': 'csv'}}`",
            "\nExample for 'Combine sales data from 'sales_q1.xlsx' sheet 'Sheet1' and 'sales_q2.xlsx' sheet 'Sheet1'.':",
            "1. `{'tool_name': 'concatenate_dataframes', 'tool_parameters': {'file_path_top': 'sales_q1.xlsx', 'sheet_name_top': 'Sheet1', 'file_path_bottom': 'sales_q2.xlsx', 'sheet_name_bottom': 'Sheet1'}}`",
            "\nAvailable Excel Files and their Structures:"
        ]

        for f_ctx in file_contexts:
            context_message_parts.append(f"\nFile: {f_ctx['file_path']}")
            for s_ctx in f_ctx['sheets']:
                context_message_parts.append(f"  Sheet: {s_ctx['sheet_name']}")
                context_message_parts.append(f"    Column Headers: {', '.join(s_ctx['column_headers']) if s_ctx['column_headers'] else 'No headers found'}")
        
        context_message_parts.append(f"\nUser Query: {user_query}")
        context_message = "\n".join(context_message_parts)
        
        # 3. Get tool call(s) from LLM
        tool_calls_response = self.llm.get_tool_call(context_message)

        if "error" in tool_calls_response:
            self.output_handler.show_error(f"LLM Tool Call Error: {tool_calls_response['error']}")
            return

        if not tool_calls_response:
            self.output_handler.show_warning("LLM did not suggest any tools for the given query.")
            return

        # 4. Execute each tool call
        # Store scalar results to be used in subsequent steps (e.g., for dynamic query strings)
        scalar_results = {} 
        last_tool_output = None # Store the output of the last executed tool

        for i, tool_call in enumerate(tool_calls_response):
            tool_name = tool_call.get("tool_name")
            tool_parameters = tool_call.get("tool_parameters", {})

            if show_all_tool_results:
                self.output_handler.print_message(f"\nExecuting Tool Call {i+1}:", style='warning')
                self.output_handler.print_message(f"LLM selected tool: {tool_name}", style='info')
                self.output_handler.print_message(f"Parameters: {json.dumps(tool_parameters, indent=2)}", style='dim')

            if tool_name not in self.tool_map:
                self.output_handler.show_error(f"LLM requested an unknown tool: '{tool_name}'")
                continue

            try:
                tool_function = self.tool_map[tool_name]
                result = None

                # Special handling for filter_and_display_dataframe to substitute scalar results
                if tool_name == "filter_and_display_dataframe" and "query_string" in tool_parameters:
                    original_query = tool_parameters["query_string"]
                    substituted_query = original_query
                    pass

                if tool_name == "load_and_display_data":
                    target_file_path = tool_parameters.get("file_path")
                    target_sheet_name = tool_parameters.get("sheet_name")
                    if not target_file_path or target_file_path not in self.excel_handlers:
                        self.output_handler.show_error(f"Tool '{tool_name}' requires a valid 'file_path' parameter which was not provided or is invalid: '{target_file_path}'.")
                        continue
                    
                    excel_handler_instance = self.excel_handlers[target_file_path]
                    result = tool_function(excel_handler_instance, **tool_parameters)
                    
                    # Update active file/sheet tracking
                    if result is not None:
                        self.active_file_path = target_file_path
                        self.active_sheet_name = target_sheet_name
                        if show_all_tool_results:
                            self.output_handler.show_success(f"Active DataFrame set to: '{self.active_file_path}' sheet '{self.active_sheet_name}'.")

                elif tool_name in ["merge_dataframes", "concatenate_dataframes"]:
                    # These tools are handled directly by ExcelAgent and take full paths
                    result = tool_function(**tool_parameters)
                elif tool_name == "perform_lookup":
                    # This tool takes a lookup file path but operates on the active_df
                    if self.active_file_path is None:
                        self.output_handler.show_error("No active Excel file/sheet. Please use 'load_and_display_data' first.")
                        continue
                    excel_handler_instance = self.excel_handlers[self.active_file_path]
                    result = tool_function(excel_handler_instance, **tool_parameters)
                else:
                    # For all other tools, they operate on the active_df of the active handler
                    if self.active_file_path is None:
                        self.output_handler.show_error("No active Excel file/sheet. Please use 'load_and_display_data' first.")
                        continue
                    
                    excel_handler_instance = self.excel_handlers[self.active_file_path]
                    result = tool_function(excel_handler_instance, **tool_parameters)
                    
                # 5. Handle the result
                if result is not None:
                    # Handle plot_dataframe specifically to get full path
                    if tool_name == "plot_dataframe" and isinstance(result, str) and (result.lower().endswith(('.png', '.jpg', '.jpeg'))):
                        # 'result' here is the output_filename (basename) returned by plot_dataframe,
                        # which is expected to be relative to the current working directory (Excel file's directory).
                        # The ExcelHandler's plot_dataframe method is assumed to save to Config.PLOTS_DIR within that CWD
                        # and return the path relative to the CWD (e.g., "plots/my_plot.png").
                        
                        # Construct the full absolute path using os.getcwd() and the result
                        full_plot_path = os.path.abspath(os.path.join(os.getcwd(), result))
                        self.output_handler.display_plot(full_plot_path, title="Generated Plot")
                        last_tool_output = full_plot_path # Store full path for final display
                        if show_all_tool_results:
                            self.output_handler.show_success(f"Operation successful! Plot saved to: {full_plot_path}")
                    elif show_all_tool_results:
                        if isinstance(result, pd.DataFrame):
                            self.output_handler.show_success("Operation successful! Here's a preview of the result:")
                            if self.excel_handlers:
                                self.output_handler.display_dataframe(result)
                            else:
                                self.output_handler.show_warning("No ExcelHandler available to display DataFrame.")
                        elif isinstance(result, str) and (result.lower().endswith(('.csv', '.json', '.xlsx', '.xls'))):
                            self.output_handler.show_success(f"File generated: {result}")
                        else:
                            self.output_handler.show_success("Operation successful! Here's the result:")
                            self.output_handler.print_message(f"Result: {result}", style='info')
                    
                    # If calculate_scalar_value was called, store its result for potential future use
                    if tool_name == "calculate_scalar_value":
                        column_name = tool_parameters.get("column")
                        agg_type = tool_parameters.get("aggregation_type")
                        query_for_scalar = tool_parameters.get("query_string")
                        
                        if column_name and agg_type:
                            key_suffix = column_name.replace(' ', '_')
                            if query_for_scalar:
                                query_hash = hash(query_for_scalar) % 1000
                                scalar_results[f"{agg_type}_{key_suffix}_{query_hash}"] = result
                            else:
                                scalar_results[f"{agg_type}_{key_suffix}"] = result
                        if show_all_tool_results:
                            self.output_handler.print_message(f"Stored scalar result: {scalar_results}", style='dim')


            except TypeError as e:
                self.output_handler.show_error(f"Error executing tool '{tool_name}': Missing or invalid parameters. Details: {e}")
                self.output_handler.print_message(f"Requested parameters: {json.dumps(tool_parameters)}", style='dim')
            except Exception as e:
                self.output_handler.show_error(f"An unexpected error occurred during tool execution: {e}")
                self.output_handler.print_message("Please review the tool's parameters or the Excel file content.", style='dim')
        
        # After all tools are executed, if not showing all results, display only the last one
        if not show_all_tool_results and last_tool_output is not None:
            self.output_handler.show_success("All operations completed. Here is the final result:")
            # Handle plot/export paths for final display
            if isinstance(last_tool_output, pd.DataFrame):
                if self.excel_handlers:
                    self.output_handler.display_dataframe(last_tool_output)
                else:
                    self.output_handler.show_warning("No ExcelHandler available to display DataFrame.")
            elif isinstance(last_tool_output, str) and (last_tool_output.lower().endswith(('.png', '.jpg', '.jpeg'))):
                self.output_handler.display_plot(last_tool_output, title="Generated Plot")
            elif isinstance(last_tool_output, str) and (last_tool_output.lower().endswith(('.csv', '.json', '.xlsx', '.xls'))):
                self.output_handler.show_success(f"File generated: {last_tool_output}")
            else:
                self.output_handler.print_message(f"Result: {last_tool_output}", style='info')
        elif not show_all_tool_results and last_tool_output is None:
            self.output_handler.show_warning("All operations completed, but no final result to display.")
