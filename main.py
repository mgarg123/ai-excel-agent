import argparse
from src.excel_agent.agent import ExcelAgent # MODIFIED
from src.excel_agent.utils import validate_excel_path # MODIFIED
from src.excel_agent.output.console_output_handler import ConsoleOutputHandler # MODIFIED
import pandas as pd

def main():
    parser = argparse.ArgumentParser(
        description="An AI agent for interacting with Excel files using natural language."
    )
    parser.add_argument(
        "file_paths",
        type=str,
        nargs='+',
        help="Path(s) to the Excel file(s) (e.g., data/sales.xlsx data/customers.xlsx)"
    )
    parser.add_argument(
        "query",
        type=str,
        help="Natural language query for Excel manipulation (e.g., 'Filter sales greater than 1000' or 'List all transactions for Gadget X')"
    )
    parser.add_argument(
        "--show-tools-execution",
        "-v",
        action="store_true", # This makes it a boolean flag, True if present, False otherwise
        help="Show detailed output for each tool execution step."
    )
    args = parser.parse_args()

    if not validate_excel_path(args.file_paths):
        return

    # NEW: Instantiate ConsoleOutputHandler for CLI mode
    output_handler = ConsoleOutputHandler()
    agent = ExcelAgent(output_handler) # MODIFIED: Pass the output_handler
    
    # Pass the new flag to the process_query method
    agent.process_query(args.file_paths, args.query, show_all_tool_results=args.show_tools_execution)

if __name__ == "__main__":
    main()
