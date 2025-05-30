import sys
import os
from PyQt6.QtWidgets import QApplication

# Add the project root to sys.path to enable absolute imports
# This assumes app.py is located at <project_root>/src/gui/app.py
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.abspath(os.path.join(current_dir, '..', '..'))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

# Now, absolute imports from the 'src' package should work
from src.gui.main_window import MainWindow

def run_gui():
    app = QApplication(sys.argv)

    # MODIFIED: Load and apply Qt Style Sheet
    qss_path = os.path.join(current_dir, 'styles', 'style.qss')
    if os.path.exists(qss_path):
        with open(qss_path, "r") as f:
            _style = f.read()
            app.setStyleSheet(_style)
    else:
        print(f"Warning: style.qss not found at {qss_path}. UI will use default styles.")

    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    run_gui()
