from abc import ABC, abstractmethod
import pandas as pd
from typing import Any

class AbstractOutputHandler(ABC):
    """
    Abstract base class defining the interface for output handling.
    Concrete implementations will handle output to console, GUI, etc.
    """

    @abstractmethod
    def print_message(self, message: str, style: str = None):
        """
        Prints a general message to the output.
        'style' can be used to indicate different message types (e.g., 'info', 'warning', 'error', 'success').
        """
        pass

    @abstractmethod
    def display_dataframe(self, df: pd.DataFrame, title: str = None):
        """
        Displays a pandas DataFrame.
        """
        pass

    @abstractmethod
    def display_plot(self, image_path: str, title: str = None):
        """
        Displays a generated plot image.
        """
        pass

    @abstractmethod
    def get_user_input(self, prompt: str) -> str:
        """
        Gets input from the user. (Primarily for CLI, GUI will use widgets)
        """
        pass

    @abstractmethod
    def show_error(self, message: str):
        """
        Displays an error message.
        """
        pass

    @abstractmethod
    def show_warning(self, message: str):
        """
        Displays a warning message.
        """
        pass

    @abstractmethod
    def show_success(self, message: str):
        """
        Displays a success message.
        """
        pass
