import inspect
import json
from functools import wraps
from typing import Any, Callable, Dict, List
from rich.console import Console

console = Console()

_registered_tools = []

def tool(description: str = "") -> Callable:
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs):
            return func(*args, **kwargs)
        
        signature = inspect.signature(func)
        properties = {}
        required = []

        for name, param in signature.parameters.items():
            if name == 'self':
                continue

            # Determine if this parameter should be skipped for schema generation
            # for tools that operate on the active_df
            if func.__name__ not in ["load_and_display_data", "apply_excel_formula", "apply_formatting", "merge_dataframes", "concatenate_dataframes", "perform_lookup"] and name in ["file_path", "sheet_name"]:
                continue # Skip adding these parameters to the schema

            param_schema = None # Initialize param_schema to None at the start of each iteration
            param_type = None
            items_type = None
            items_properties = None
            items_required = None

            param_description = f"The {name} parameter."
            
            # Specific tool parameter handling
            if func.__name__ == "filter_and_display_dataframe" and name == "query_string":
                param_schema = {
                    "description": "A string representing the query to filter the DataFrame, using pandas.DataFrame.query() syntax. This allows for complex boolean logic (e.g., 'and', 'or', 'not'). Column names with spaces must be enclosed in backticks (e.g., `Column Name`). Example: \"(`Discount Amount` > 500) and (`Net Revenue` < 2000)\"",
                    "type": "string"
                }
            # Special handling for display_head_or_tail parameters
            elif func.__name__ == "display_head_or_tail":
                if name == "num_rows":
                    param_schema = {"type": "integer", "description": "The number of rows to display from the head or tail of the DataFrame. Defaults to 5."}
                elif name == "from_end":
                    param_schema = {"type": "boolean", "description": "If true, displays rows from the end (tail); otherwise, from the beginning (head). Defaults to false."}
                else: # Fallback for other parameters of this specific tool if any
                    param_schema = {"type": "string", "description": param_description}
            # Special handling for calculate_scalar_value aggregation_type and query_string
            elif func.__name__ == "calculate_scalar_value":
                if name == "aggregation_type":
                    param_schema = {
                        "type": "string",
                        "enum": ["sum", "mean", "count", "min", "max", "std"], # Added 'std'
                        "description": "The type of aggregation to perform."
                    }
                elif name == "query_string":
                    param_schema = {
                        "type": "string",
                        "description": "Optional: A pandas query string to temporarily filter the DataFrame *before* aggregation. Column names with spaces must be enclosed in backticks (e.g., `Product Name`). This filter is only for the calculation and does not change the active DataFrame."
                    }
                else: # For 'column' parameter
                    param_schema = {"type": "string", "description": param_description}
            # Special handling for compare_values parameters
            elif func.__name__ == "compare_values" and name == "comparisons":
                param_schema = {
                    "description": "A list of dictionaries, each specifying a value to calculate and compare. Each dictionary must contain 'label' (string), 'column' (string), 'aggregation_type' (string: 'sum', 'mean', 'count', 'min', 'max', 'std'), and optionally 'query_string' (string for filtering).", # Added 'std'
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "label": {"type": "string", "description": "A descriptive label for the value (e.g., 'Total Sales for Region A')."},
                            "column": {"type": "string", "description": "The column name to aggregate."},
                            "aggregation_type": {"type": "string", "enum": ["sum", "mean", "count", "min", "max", "std"], "description": "The type of aggregation to perform."}, # Added 'std'
                            "query_string": {"type": "string", "description": "Optional: A pandas query string to filter the DataFrame before aggregation. Column names with spaces must be enclosed in backticks."}
                        },
                        "required": ["label", "column", "aggregation_type"]
                    }
                }
            # New tool: extract_date_part
            elif func.__name__ == "extract_date_part":
                if name == "date_column":
                    param_schema = {"type": "string", "description": "The name of the column containing date values."}
                elif name == "part":
                    param_schema = {"type": "string", "enum": ["year", "month", "day", "quarter"], "description": "The date part to extract."}
                elif name == "new_column_name":
                    param_schema = {"type": "string", "description": "The name for the new column that will store the extracted part."}
                else: # Fallback for other parameters of this specific tool if any
                    param_schema = {"type": "string", "description": param_description}
            # New tool: add_lagged_column
            elif func.__name__ == "add_lagged_column":
                if name == "column":
                    param_schema = {"type": "string", "description": "The column for which to calculate lagged values."}
                elif name == "new_column_name":
                    param_schema = {"type": "string", "description": "The name for the new column (e.g., 'Previous Month Sales')."}
                elif name == "periods":
                    param_schema = {"type": "integer", "description": "The number of periods to shift (default is 1 for previous row)."}
                elif name == "group_by_columns":
                    param_schema = {"type": "array", "items": {"type": "string"}, "description": "Optional list of columns to group by before applying the lag. This ensures the lag is calculated within each group (e.g., per region)."}
                else: # Fallback for other parameters of this specific tool if any
                    param_schema = {"type": "string", "description": param_description}
            # New tool: plot_dataframe
            elif func.__name__ == "plot_dataframe":
                if name == "plot_type":
                    param_schema = {"type": "string", "enum": ["line", "bar", "scatter", "hist", "box", "pie"], "description": "The type of plot to generate. Note: For radar charts, use 'plot_radar_chart' instead."}
                elif name == "output_filename":
                    param_schema = {"type": "string", "description": "The name of the file to save the plot (e.g., 'my_chart.png'). Must end with .png, .jpg, or .jpeg."}
                elif name == "x_column":
                    param_schema = {"type": "string", "description": "Optional: The column for the x-axis (or labels for pie chart). Required for line, bar, scatter, pie plots."}
                elif name == "y_column":
                    param_schema = {"type": "string", "description": "Optional: The column for the y-axis (or values for pie chart). Required for line, bar, scatter, pie plots."}
                elif name == "title":
                    param_schema = {"type": "string", "description": "Optional: The title of the plot."}
                elif name == "x_label":
                    param_schema = {"type": "string", "description": "Optional: Label for the x-axis."}
                elif name == "y_label":
                    param_schema = {"type": "string", "description": "Optional: Label for the y-axis."}
                elif name == "hue_column":
                    param_schema = {"type": "string", "description": "Optional: A column to use for color encoding (e.g., grouping in scatter plots or different lines in line plot)."}
                else: # Fallback for other parameters of this specific tool if any
                    param_schema = {"type": "string", "description": param_description}
            # New tool: plot_radar_chart
            elif func.__name__ == "plot_radar_chart":
                if name == "category_column":
                    param_schema = {"type": "string", "description": "The column containing the categories to compare (e.g., 'Region')."}
                elif name == "value_columns":
                    param_schema = {"type": "array", "items": {"type": "string"}, "description": "A list of numeric columns representing the metrics to plot on the radar axes (e.g., ['Revenue', 'Expenses', 'Profit']). This tool will automatically calculate the *mean* of these columns for each category."}
                elif name == "output_filename":
                    param_schema = {"type": "string", "description": "The name of the file to save the plot (e.g., 'radar_chart.png'). Must end with .png, .jpg, or .jpeg."}
                elif name == "title":
                    param_schema = {"type": "string", "description": "Optional: The title of the radar chart."}
                else: # Fallback for other parameters of this specific tool if any
                    param_schema = {"type": "string", "description": param_description}
            # New tool: convert_column_type
            elif func.__name__ == "convert_column_type":
                if name == "column":
                    param_schema = {"type": "string", "description": "The name of the column to convert."}
                elif name == "target_type":
                    param_schema = {"type": "string", "enum": ["numeric", "datetime", "string"], "description": "The target data type for the column."}
                else: # Fallback for other parameters of this specific tool if any
                    param_schema = {"type": "string", "description": param_description}
            # New tool: split_column_by_delimiter
            elif func.__name__ == "split_column_by_delimiter":
                if name == "column":
                    param_schema = {"type": "string", "description": "The name of the column to split."}
                elif name == "delimiter":
                    param_schema = {"type": "string", "description": "The character or string to split by (e.g., ', ', '-')."}
                elif name == "new_column_names":
                    param_schema = {"type": "array", "items": {"type": "string"}, "description": "A list of names for the new columns created by the split. The number of names should match the expected number of parts after splitting."}
                else: # Fallback for other parameters of this specific tool if any
                    param_schema = {"type": "string", "description": param_description}
            # New tool: extract_pattern_from_column
            elif func.__name__ == "extract_pattern_from_column":
                if name == "column":
                    param_schema = {"type": "string", "description": "The name of the column to extract from."}
                elif name == "regex_pattern":
                    param_schema = {"type": "string", "description": "The regular expression pattern to use (e.g., r'\\d+' for numbers, r'\\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\\.[A-Z|a-z]{2,}\\b' for emails)."}
                elif name == "new_column_name":
                    param_schema = {"type": "string", "description": "The name for the new column that will store the extracted data."}
                elif name == "group_index":
                    param_schema = {"type": "integer", "description": "The index of the capturing group in the regex to extract (default is 0 for the entire match)."}
                else: # Fallback for other parameters of this specific tool if any
                    param_schema = {"type": "string", "description": param_description}
            # New tool: clean_text_column
            elif func.__name__ == "clean_text_column":
                if name == "column":
                    param_schema = {"type": "string", "description": "The name of the text column to clean."}
                elif name == "cleaning_operations":
                    param_schema = {"type": "array", "items": {"type": "string", "enum": ["strip", "lower", "upper", "remove_digits", "remove_punctuation"]}, "description": "A list of operations to apply ('strip', 'lower', 'upper', 'remove_digits', 'remove_punctuation')."}
                else: # Fallback for other parameters of this specific tool if any
                    param_schema = {"type": "string", "description": param_description}
            # New tool: perform_lookup
            elif func.__name__ == "perform_lookup":
                if name == "lookup_file_path":
                    param_schema = {"type": "string", "description": "Path to the Excel file containing the lookup data."}
                elif name == "lookup_sheet_name":
                    param_schema = {"type": "string", "description": "Name of the sheet in the lookup file."}
                elif name == "on_column_main_df":
                    param_schema = {"type": "string", "description": "The column in the active DataFrame to match on."}
                elif name == "on_column_lookup_df":
                    param_schema = {"type": "string", "description": "The column in the lookup DataFrame to match on."}
                elif name == "columns_to_add":
                    param_schema = {"type": "array", "items": {"type": "string"}, "description": "A list of column names from the lookup DataFrame to add to the active DataFrame."}
                elif name == "how":
                    param_schema = {"type": "string", "enum": ["left", "inner", "right", "outer"], "description": "Type of merge ('left', 'inner', 'right', 'outer'). Defaults to 'left' for VLOOKUP-like behavior."}
                else: # Fallback for other parameters of this specific tool if any
                    param_schema = {"type": "string", "description": param_description}
            # New tool: impute_missing_values_advanced
            elif func.__name__ == "impute_missing_values_advanced":
                if name == "column":
                    param_schema = {"type": "string", "description": "The name of the column to impute."}
                elif name == "strategy":
                    param_schema = {"type": "string", "enum": ["ffill", "bfill", "interpolate"], "description": "The imputation method ('ffill' for forward-fill, 'bfill' for backward-fill, 'interpolate')."}
                elif name == "limit":
                    param_schema = {"type": "integer", "description": "Optional. The maximum number of consecutive NaN values to fill."}
                else: # Fallback for other parameters of this specific tool if any
                    param_schema = {"type": "string", "description": param_description}
            # New tool: export_dataframe
            elif func.__name__ == "export_dataframe":
                if name == "output_file_path":
                    param_schema = {"type": "string", "description": "The path and filename for the output file (e.g., 'output_data.csv')."}
                elif name == "output_format":
                    param_schema = {"type": "string", "enum": ["csv", "json", "excel"], "description": "The desired output format ('csv', 'json', or 'excel')."}
                else: # Fallback for other parameters of this specific tool if any
                    param_schema = {"type": "string", "description": param_description}
            # New tool: concatenate_dataframes (in agent.py, but schema defined here)
            elif func.__name__ == "concatenate_dataframes":
                if name == "file_path_top":
                    param_schema = {"type": "string", "description": "Path to the first Excel file."}
                elif name == "sheet_name_top":
                    param_schema = {"type": "string", "description": "Name of the sheet in the first Excel file."}
                elif name == "file_path_bottom":
                    param_schema = {"type": "string", "description": "Path to the second Excel file."}
                elif name == "sheet_name_bottom":
                    param_schema = {"type": "string", "description": "Name of the sheet in the second Excel file."}
                else: # Fallback for other parameters of this specific tool if any
                    param_schema = {"type": "string", "description": param_description}
            # Generic parameter handling
            else: 
                if param.annotation is inspect.Parameter.empty:
                    param_type = "string"
                elif hasattr(param.annotation, '__origin__') and param.annotation.__origin__ is list:
                    param_type = "array"
                    if hasattr(param.annotation, '__args__') and param.annotation.__args__:
                        arg_type = param.annotation.__args__[0]
                        if arg_type is str:
                            items_type = "string"
                        elif arg_type is int or arg_type is float:
                            items_type = "number"
                        elif arg_type is bool:
                            items_type = "boolean"
                        elif arg_type is dict or arg_type is Any:
                            items_type = "object"
                            items_properties = {} # Default for generic List[Dict]
                        else:
                            items_type = "string"
                    else:
                        items_type = "string"
                elif param.annotation is str:
                    param_type = "string"
                elif param.annotation is int or param.annotation is float:
                    param_type = "number"
                elif param.annotation is bool:
                    param_type = "boolean"
                elif param.annotation is dict:
                    param_type = "object"
                elif param.annotation is Any:
                    param_type = "string" # Default for other Any types

                param_schema = {"description": param_description}
                if param_type:
                    param_schema["type"] = param_type
                    if param_type == "array" and items_type:
                        param_schema["items"] = {"type": items_type}
                        if items_type == "object" and items_properties:
                            param_schema["items"]["properties"] = items_properties
                            if items_required:
                                param_schema["items"]["required"] = items_required
            
            # Add the generated schema to properties if it was successfully created
            if param_schema is not None:
                properties[name] = param_schema

            # Only add to required if no default value is provided
            # and if the parameter was not skipped (e.g., file_path/sheet_name for active_df tools)
            if param.default is inspect.Parameter.empty and name in properties: # Check if parameter was actually added to properties
                required.append(name)

        tool_schema = {
            "type": "function",
            "function": {
                "name": func.__name__,
                "description": description or func.__doc__ or f"Execute the {func.__name__} function.",
                "parameters": {
                    "type": "object",
                    "properties": properties,
                    "required": required
                }
            }
        }

        _registered_tools.append(tool_schema)
        # console.print(f"[bold yellow]Generated Tool Schema for {func.__name__}:[/bold yellow] {json.dumps(tool_schema, indent=2)}")
        return wrapper
    return decorator

def get_registered_tools():
    return _registered_tools
