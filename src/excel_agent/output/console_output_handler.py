import pandas as pd
from rich.console import Console
from rich.table import Table
from rich.box import SIMPLE_HEAD
from typing import Any

from src.excel_agent.output.abstract_output_handler import AbstractOutputHandler # MODIFIED

class ConsoleOutputHandler(AbstractOutputHandler):
    """
    Concrete implementation of AbstractOutputHandler for console output using rich.
    """
    def __init__(self):
        self.console = Console()

    def print_message(self, message: str, style: str = None):
        """
        Prints a general message to the console with optional styling.
        """
        if style == 'info':
            self.console.print(f"[blue]{message}[/blue]")
        elif style == 'warning':
            self.console.print(f"[bold yellow]Warning:[/bold yellow] {message}")
        elif style == 'error':
            self.console.print(f"[bold red]Error:[/bold red] {message}")
        elif style == 'success':
            self.console.print(f"[bold green]{message}[/bold green]")
        elif style == 'dim':
            self.console.print(f"[dim]{message}[/dim]")
        else:
            self.console.print(message)

    def display_dataframe(self, df: pd.DataFrame, title: str = None):
        """
        Displays a pandas DataFrame in the console using rich.
        """
        if df.empty:
            self.print_message("DataFrame is empty.", style='warning')
            return

        if title:
            self.print_message(f"\n[bold magenta]{title}[/bold magenta]")

        # Limit rows for display to avoid overwhelming the console
        display_df = df.head(10)
        if len(df) > 10:
            self.print_message(f"Displaying first 10 rows of {len(df)} total rows.", style='dim')

        table = Table(box=SIMPLE_HEAD, show_header=True, header_style="bold magenta")
        
        # Add columns to the table
        for col in display_df.columns:
            table.add_column(str(col))
        
        # Add rows to the table
        for _, row in display_df.iterrows():
            table.add_row(*[str(item) for item in row.values])
        
        self.console.print(table)

    def display_plot(self, image_path: str, title: str = None):
        """
        Informs the user where the plot has been saved.
        """
        if title:
            self.print_message(f"[bold green]{title}[/bold green]")
        self.print_message(f"Plot saved to: [cyan]{image_path}[/cyan]", style='success')

    def get_user_input(self, prompt: str) -> str:
        """
        Gets input from the user via the console.
        """
        return self.console.input(f"[bold blue]{prompt}[/bold blue] ")

    def show_error(self, message: str):
        self.print_message(message, style='error')

    def show_warning(self, message: str):
        self.print_message(message, style='warning')

    def show_success(self, message: str):
        self.print_message(message, style='success')
