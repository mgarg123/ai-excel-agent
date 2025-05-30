import pandas as pd
from typing import Any
from PyQt6.QtCore import QObject, pyqtSignal, Qt
from PyQt6.QtGui import QPixmap, QImage
import io
import os

from src.excel_agent.output.abstract_output_handler import AbstractOutputHandler
from abc import ABCMeta # Added: Import ABCMeta for metaclass handling

# Get the metaclass of QObject
QObjectMetaclass = type(QObject) # Added: Get QObject's metaclass

# Define a custom metaclass that inherits from both QObject's metaclass and ABCMeta
class CombinedMetaclass(QObjectMetaclass, ABCMeta): # Added: Custom metaclass to resolve conflict
    pass

class GuiOutputHandler(QObject, AbstractOutputHandler, metaclass=CombinedMetaclass): # MODIFIED: Apply the custom metaclass
    """
    Concrete implementation of AbstractOutputHandler for GUI output using PyQt6 signals.
    Inherits from QObject to enable signals.
    """
    # Define signals to emit different types of output to the GUI
    message_signal = pyqtSignal(str, str) # message, style
    dataframe_signal = pyqtSignal(pd.DataFrame, str) # dataframe, title
    plot_signal = pyqtSignal(str, str) # image_path, title
    input_request_signal = pyqtSignal(str) # prompt
    error_signal = pyqtSignal(str) # message
    warning_signal = pyqtSignal(str) # message
    success_signal = pyqtSignal(str) # message

    def __init__(self):
        # Call QObject's constructor first.
        # super() will correctly handle the MRO for both QObject and AbstractOutputHandler.
        super().__init__()

    def print_message(self, message: str, style: str = None):
        """
        Emits a message to be displayed in the GUI.
        """
        self.message_signal.emit(message, style or 'info')

    def display_dataframe(self, df: pd.DataFrame, title: str = None):
        """
        Emits a DataFrame to be displayed in a GUI table or text area.
        """
        self.dataframe_signal.emit(df, title or "DataFrame Result")

    def display_plot(self, image_path: str, title: str = None):
        """
        Emits the path to a generated plot image to be displayed in the GUI.
        """
        self.plot_signal.emit(image_path, title or "Generated Plot")

    def get_user_input(self, prompt: str) -> str:
        """
        For GUI, this would typically involve showing a dialog and waiting for user input.
        However, for simplicity in this initial setup, we'll assume direct input via a text field
        and not block the main thread with a modal dialog here.
        If interactive input is strictly needed mid-process, a QInputDialog or custom dialog
        would be shown, and the worker thread would need to pause.
        For now, this method will raise an error if called, as GUI input is handled differently.
        """
        self.show_error("Interactive input via get_user_input is not supported in GUI mode for mid-process queries.")
        raise NotImplementedError("get_user_input is not directly supported for interactive GUI flow.")

    def show_error(self, message: str):
        self.error_signal.emit(message)
        self.print_message(f"Error: {message}", style='error') # Also print to general message area

    def show_warning(self, message: str):
        self.warning_signal.emit(message)
        self.print_message(f"Warning: {message}", style='warning') # Also print to general message area

    def show_success(self, message: str):
        self.success_signal.emit(message)
        self.print_message(f"Success: {message}", style='success') # Also print to general message area
