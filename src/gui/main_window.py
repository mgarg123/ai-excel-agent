import sys
import os
import pandas as pd
import shutil
from PyQt6.QtWidgets import (
    QMainWindow, QVBoxLayout, QHBoxLayout, QWidget, QPushButton,
    QTextEdit, QLineEdit, QFileDialog, QLabel, QTableWidget,
    QTableWidgetItem, QHeaderView, QMessageBox, QProgressDialog,
    QAbstractItemView, QGroupBox, QSplitter, QTabWidget,
    QStatusBar, QSpacerItem, QSizePolicy
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QPixmap, QImage, QTextCharFormat, QColor, QFont, QTextCursor, QIcon, QAction

from src.excel_agent.agent import ExcelAgent
from src.excel_agent.output.gui_output_handler import GuiOutputHandler
from src.excel_agent.utils import validate_excel_path

# Worker thread for running the ExcelAgent to keep the GUI responsive
class AgentWorker(QThread):
    finished = pyqtSignal()
    error = pyqtSignal(str)
    
    def __init__(self, agent: ExcelAgent, file_paths: list, query: str, show_all_tool_results: bool):
        super().__init__()
        self.agent = agent
        self.file_paths = file_paths
        self.query = query
        self.show_all_tool_results = show_all_tool_results

    def run(self):
        try:
            self.agent.process_query(self.file_paths, self.query, self.show_all_tool_results)
        except Exception as e:
            self.error.emit(f"An unexpected error occurred during agent processing: {e}")
        finally:
            self.finished.emit()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel AI Agent GUI")
        self.setGeometry(100, 100, 1200, 800) # Initial window size (will be overridden by showMaximized)

        self.output_handler = GuiOutputHandler()
        self.excel_agent = ExcelAgent(self.output_handler)

        self.current_file_paths = [] # Stores full paths
        self.original_cwd = os.getcwd() # Store original CWD
        self.agent_worker = None # To hold the worker thread instance
        self.progress_dialog = None # To hold the progress dialog
        self.current_plot_path = None # New attribute to store the path of the current plot

        self.setWindowIcon(QIcon(os.path.join(os.path.dirname(__file__), 'icons', 'app_icon.png'))) # Set application icon (requires app_icon.png)

        self.init_ui()
        self.connect_signals()
        
        self.showMaximized() # MODIFIED: Open the application in maximized mode

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(15, 15, 15, 15) # Add margins
        main_layout.setSpacing(10) # Add spacing

        # --- Query Input & File Selection Area ---
        query_group_box = QGroupBox("1. Enter Query & Select File(s)") # Combined GroupBox
        query_layout = QHBoxLayout(query_group_box)
        query_layout.setSpacing(10)

        self.query_input = QLineEdit()
        self.query_input.setPlaceholderText("Enter your natural language query here...")
        query_layout.addWidget(self.query_input)

        # Add file browse action directly to QLineEdit with the new icon path
        self.browse_action = QAction(QIcon(os.path.join(os.path.dirname(__file__), 'images', 'file_upload.png')), "Browse Files", self)
        self.query_input.addAction(self.browse_action, QLineEdit.ActionPosition.TrailingPosition) # Place on the right side

        self.process_button = QPushButton("Process Query")
        self.process_button.setIcon(QIcon(os.path.join(os.path.dirname(__file__), 'icons', 'play.png')))
        query_layout.addWidget(self.process_button)
        main_layout.addWidget(query_group_box) # Add group box to main layout

        # --- Output Display Area (Tabs) ---
        output_group_box = QGroupBox("2. Results & Plots") # Renumbered GroupBox and title
        output_layout = QVBoxLayout(output_group_box)
        output_layout.setSpacing(10)

        self.output_tab_widget = QTabWidget() # New QTabWidget

        # --- Results Tab ---
        results_tab = QWidget()
        results_layout = QVBoxLayout(results_tab)
        results_layout.setContentsMargins(0, 0, 0, 0) # No extra margins inside tab
        results_layout.setSpacing(5)

        results_splitter = QSplitter(Qt.Orientation.Vertical) # Splitter for text and table
        
        # Text output (messages, warnings, errors)
        self.output_text_edit = QTextEdit()
        self.output_text_edit.setReadOnly(True)
        results_splitter.addWidget(self.output_text_edit)

        # DataFrame display
        self.dataframe_table = QTableWidget()
        self.dataframe_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.dataframe_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.dataframe_table.verticalHeader().setVisible(False) # Hide row numbers
        results_splitter.addWidget(self.dataframe_table)

        results_splitter.setSizes([250, 550]) # Initial sizes for text and table within results tab
        results_layout.addWidget(results_splitter)
        self.output_tab_widget.addTab(results_tab, "Results") # Add Results tab

        # --- Plots Tab ---
        plots_tab = QWidget()
        plots_layout = QVBoxLayout(plots_tab)
        plots_layout.setContentsMargins(10, 10, 10, 10) # Add some margins for the tab content
        plots_layout.setSpacing(10)

        # Layout for the export button
        export_button_layout = QHBoxLayout()
        export_button_layout.addStretch(1) # Pushes button to the right

        self.export_plot_button = QPushButton("Export Plot")
        self.export_plot_button.setIcon(QIcon(os.path.join(os.path.dirname(__file__), 'icons', 'save.png'))) # Assuming save.png
        self.export_plot_button.setFixedWidth(120) # Give it a fixed width
        self.export_plot_button.setEnabled(False) # Initially disabled
        export_button_layout.addWidget(self.export_plot_button)
        
        plots_layout.addLayout(export_button_layout) # Add button layout to plots tab layout

        self.plot_label = QLabel("Generated Plot will appear here")
        self.plot_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.plot_label.setMinimumSize(200, 150) # Set a minimum size but allow stretching
        self.plot_label.setStyleSheet("border: 1px solid lightgrey; background-color: #f0f0f0;")
        plots_layout.addWidget(self.plot_label, alignment=Qt.AlignmentFlag.AlignCenter) # Add plot label to plots tab layout

        self.output_tab_widget.addTab(plots_tab, "Plots")

        output_layout.addWidget(self.output_tab_widget) # Add the tab widget to the output group box
        main_layout.addWidget(output_group_box) # Add output group box to main layout

        # Status Bar
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        self.statusBar.showMessage("Ready. Please select Excel file(s) and enter a query.")

    def connect_signals(self):
        self.browse_action.triggered.connect(self.browse_files)
        self.process_button.clicked.connect(self.process_user_query)
        self.export_plot_button.clicked.connect(self.export_plot)

        # Connect signals from GuiOutputHandler to UI update slots
        self.output_handler.message_signal.connect(self.append_output_message)
        self.output_handler.dataframe_signal.connect(self.display_dataframe_in_table)
        self.output_handler.plot_signal.connect(self.display_plot_image)
        self.output_handler.error_signal.connect(self.show_error_messagebox)
        self.output_handler.warning_signal.connect(self.show_warning_messagebox)
        self.output_handler.success_signal.connect(self.show_success_messagebox)

    def browse_files(self):
        # Allow selecting multiple Excel files
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.FileMode.ExistingFiles)
        file_dialog.setNameFilter("Excel Files (*.xlsx *.xls)")
        
        if file_dialog.exec():
            selected_files = file_dialog.selectedFiles()
            if selected_files:
                # Validate selected files
                if not validate_excel_path(selected_files):
                    QMessageBox.warning(self, "Invalid Files", "Some selected files are not valid Excel files or do not exist.")
                    self.current_file_paths = []
                    self.statusBar.showMessage("Error: Invalid file(s) selected.", 5000)
                    return

                self.current_file_paths = selected_files # Store full paths
                
                self.output_text_edit.clear()
                # Display selected files in output text area
                file_names_display = [os.path.basename(f) for f in selected_files]
                self.append_output_message(f"Selected file(s): {'; '.join(file_names_display)}", style='info')
                
                self.clear_dataframe_table()
                self.clear_plot_display() # Clear plot label content and disable button
                self.output_tab_widget.setCurrentIndex(0) # Switch to Results tab
                self.statusBar.showMessage(f"Selected {len(selected_files)} file(s). Ready for query.", 3000)
            else:
                self.current_file_paths = []
                self.statusBar.showMessage("No files selected.", 3000)

    def process_user_query(self):
        if not self.current_file_paths:
            QMessageBox.warning(self, "No File Selected", "Please select at least one Excel file first.")
            self.statusBar.showMessage("Error: No Excel file selected.", 5000)
            return

        user_query = self.query_input.text().strip()
        if not user_query:
            QMessageBox.warning(self, "Empty Query", "Please enter a query.")
            self.statusBar.showMessage("Error: Query is empty.", 5000)
            return

        # Clear previous results
        self.output_text_edit.clear()
        self.clear_dataframe_table()
        self.clear_plot_display() # Clear plot label content and disable button
        self.output_tab_widget.setCurrentIndex(0) # Switch to Results tab before processing
        self.append_output_message("Processing query...", style='info')
        self.statusBar.showMessage("Processing query...", 0)
        self.set_ui_enabled(False) # Disable UI during processing

        # Change CWD and prepare file paths for the agent
        try:
            # Get the directory of the first selected file
            first_file_dir = os.path.dirname(self.current_file_paths[0])
            # Change the current working directory to where the Excel files are located
            os.chdir(first_file_dir)
            self.append_output_message(f"Changed working directory to: {first_file_dir}", style='dim')
            
            # Pass only basenames to the agent, as the CWD is now set
            file_paths_for_agent = [os.path.basename(p) for p in self.current_file_paths]
            
        except Exception as e:
            self.set_ui_enabled(True)
            if self.progress_dialog:
                self.progress_dialog.close()
            self.show_error_messagebox(f"Failed to set working directory or process file paths: {e}")
            self.append_output_message(f"\nError: Failed to set working directory or process file paths: {e}", style='error')
            self.statusBar.showMessage("Error: Failed to set working directory.", 5000)
            return

        # Show progress dialog
        self.progress_dialog = QProgressDialog("Processing...", "Cancel", 0, 0, self)
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setCancelButton(None) # No cancel button for now, as agent doesn't support interruption
        self.progress_dialog.setWindowTitle("AI Agent Progress")
        self.progress_dialog.setLabelText("AI Agent is working on your request. This may take a moment...")
        self.progress_dialog.show()

        # Create and start the worker thread
        self.agent_worker = AgentWorker(self.excel_agent, file_paths_for_agent, user_query, show_all_tool_results=True)
        self.agent_worker.finished.connect(self.on_agent_finished)
        self.agent_worker.error.connect(self.on_agent_error)
        self.agent_worker.start()

    def on_agent_finished(self):
        self.set_ui_enabled(True) # Re-enable UI
        if self.progress_dialog:
            self.progress_dialog.close()
        self.append_output_message("\nQuery processing completed.", style='success')
        self.statusBar.showMessage("Query processing completed successfully.", 3000)
        # Restore original working directory
        os.chdir(self.original_cwd)
        self.append_output_message(f"Restored working directory to: {self.original_cwd}", style='dim')


    def on_agent_error(self, error_message: str):
        self.set_ui_enabled(True) # Re-enable UI
        if self.progress_dialog:
            self.progress_dialog.close()
        self.show_error_messagebox(error_message)
        self.append_output_message(f"\nError during processing: {error_message}", style='error')
        self.statusBar.showMessage("Error during query processing. Check output for details.", 5000)
        # Restore original working directory
        os.chdir(self.original_cwd)
        self.append_output_message(f"Restored working directory to: {self.original_cwd}", style='dim')

    def set_ui_enabled(self, enabled: bool):
        self.browse_action.setEnabled(enabled)
        self.query_input.setEnabled(enabled)
        self.process_button.setEnabled(enabled)
        # The export_plot_button's enabled state is managed by display_plot_image/clear_plot_display

    def append_output_message(self, message: str, style: str):
        # Define text formats for different styles
        text_format = QTextCharFormat()
        if style == 'info':
            text_format.setForeground(QColor("#007bff"))
        elif style == 'warning':
            text_format.setForeground(QColor("#ffc107"))
            text_format.setFontWeight(QFont.Weight.Bold)
        elif style == 'error':
            text_format.setForeground(QColor("#dc3545"))
            text_format.setFontWeight(QFont.Weight.Bold)
        elif style == 'success':
            text_format.setForeground(QColor("#28a745"))
            text_format.setFontWeight(QFont.Weight.Bold)
        elif style == 'dim':
            text_format.setForeground(QColor("#6c757d"))
        else:
            text_format.setForeground(QColor("#333333"))

        cursor = self.output_text_edit.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        cursor.insertText(message + "\n", text_format)
        self.output_text_edit.setTextCursor(cursor)
        self.output_text_edit.ensureCursorVisible()

    def display_dataframe_in_table(self, df: pd.DataFrame, title: str):
        self.clear_dataframe_table()
        self.append_output_message(f"\n--- {title} ---", style='info')

        if df.empty:
            self.append_output_message("DataFrame is empty.", style='warning')
            return

        self.dataframe_table.setRowCount(df.shape[0])
        self.dataframe_table.setColumnCount(df.shape[1])
        self.dataframe_table.setHorizontalHeaderLabels(df.columns.astype(str))

        # Limit rows for display to avoid overwhelming the GUI
        display_rows = min(len(df), 100) # Display max 100 rows
        if len(df) > display_rows:
            self.append_output_message(f"Displaying first {display_rows} rows of {len(df)} total rows.", style='dim')

        for i in range(display_rows):
            for j in range(df.shape[1]):
                item = QTableWidgetItem(str(df.iloc[i, j]))
                self.dataframe_table.setItem(i, j, item)
        
        self.dataframe_table.resizeColumnsToContents()
        self.dataframe_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive) # Allow manual resize
        self.dataframe_table.horizontalHeader().setStretchLastSection(True) # Stretch last section to fill space

    def clear_dataframe_table(self):
        self.dataframe_table.clearContents()
        self.dataframe_table.setRowCount(0)
        self.dataframe_table.setColumnCount(0)
        self.dataframe_table.setHorizontalHeaderLabels([])

    def display_plot_image(self, image_path: str, title: str):
        self.clear_plot_display() # Clear existing plot content and disable button
        self.append_output_message(f"\n--- {title} ---", style='info')
        self.append_output_message(f"Plot saved to: {image_path}", style='success')

        if os.path.exists(image_path):
            pixmap = QPixmap(image_path)
            if not pixmap.isNull():
                # Scale pixmap to fit the label while maintaining aspect ratio
                scaled_pixmap = pixmap.scaled(self.plot_label.size(), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
                self.plot_label.setPixmap(scaled_pixmap)
                self.plot_label.setText("") # Clear "Generated Plot will appear here" text
                self.output_tab_widget.setCurrentIndex(1) # Switch to Plots tab (index 1)
                self.current_plot_path = image_path # Store the path
                self.export_plot_button.setEnabled(True) # Enable export button
            else:
                self.plot_label.setText("Failed to load image.")
                self.append_output_message(f"Error: Could not load image from '{image_path}'.", style='error')
                self.current_plot_path = None # Clear path on failure
                self.export_plot_button.setEnabled(False) # Disable button
        else:
            self.plot_label.setText("Plot file not found.")
            self.append_output_message(f"Error: Plot file not found at '{image_path}'.", style='error')
            self.current_plot_path = None # Clear path on failure
            self.export_plot_button.setEnabled(False) # Disable button

    def clear_plot_display(self):
        self.plot_label.clear()
        self.plot_label.setText("Generated Plot will appear here")
        self.current_plot_path = None # Clear the stored path
        self.export_plot_button.setEnabled(False) # Disable the export button

    def export_plot(self):
        if not self.current_plot_path or not os.path.exists(self.current_plot_path):
            QMessageBox.warning(self, "No Plot to Export", "There is no plot currently displayed or the file does not exist.")
            self.statusBar.showMessage("No plot to export.", 3000)
            return

        # Suggest a default filename based on the current plot's filename
        default_filename = os.path.basename(self.current_plot_path)
        
        # Open a save file dialog
        file_dialog = QFileDialog()
        # Set initial directory to the directory of the current plot
        initial_dir = os.path.dirname(self.current_plot_path) if os.path.exists(self.current_plot_path) else os.getcwd()
        
        save_path, _ = file_dialog.getSaveFileName(self, "Save Plot As", 
                                                   os.path.join(initial_dir, default_filename), 
                                                   "PNG Image (*.png);;JPEG Image (*.jpg *.jpeg);;All Files (*)")

        if save_path:
            try:
                shutil.copy(self.current_plot_path, save_path)
                QMessageBox.information(self, "Plot Exported", f"Plot successfully saved to:\n{save_path}")
                self.statusBar.showMessage(f"Plot exported to {os.path.basename(save_path)}", 3000)
            except Exception as e:
                QMessageBox.critical(self, "Export Error", f"Failed to save plot:\n{e}")
                self.statusBar.showMessage(f"Error exporting plot: {e}", 5000)
        else:
            self.statusBar.showMessage("Plot export cancelled.", 3000)

    def show_error_messagebox(self, message: str):
        QMessageBox.critical(self, "Error", message)
        self.statusBar.showMessage(f"Error: {message}", 5000)

    def show_warning_messagebox(self, message: str):
        QMessageBox.warning(self, "Warning", message)
        self.statusBar.showMessage(f"Warning: {message}", 5000)

    def show_success_messagebox(self, message: str):
        # For success, just append to text edit, no need for a popup unless critical
        pass # Handled by append_output_message already
