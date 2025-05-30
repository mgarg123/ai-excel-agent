# Excel AI Agent

## Overview

The Excel AI Agent is a powerful desktop application built with Python and PyQt6 that allows users to interact with and analyze Excel (`.xlsx`, `.xls`) and CSV (`.csv`) files using natural language queries. Leveraging a Large Language Model (LLM) from Groq, it translates your requests into data manipulation and visualization operations, providing insights through tabular data and interactive plots.

This application aims to simplify data analysis for users who may not be proficient in programming or complex spreadsheet formulas, offering an intuitive graphical interface to unlock the power of their data.

## Features

*   **Natural Language Interaction**: Ask questions and give commands in plain English.
*   **Multi-file Support**: Load and process data from single or multiple Excel (`.xlsx`, `.xls`) and CSV (`.csv`) files.
*   **Comprehensive Data Manipulation**:
    *   Load and display data from specific sheets.
    *   Filter data based on complex criteria.
    *   Group and aggregate data (sum, mean, count, min, max, std) with predictable column naming.
    *   Sort data by multiple columns.
    *   Add new columns using custom formulas.
    *   Handle missing values (fill, drop, advanced imputation).
    *   Remove duplicate rows.
    *   Rename columns.
    *   Select specific columns.
    *   Generate descriptive statistics.
    *   Delete rows or columns.
    *   Create pivot tables.
    *   Extract date parts (year, month, day, quarter).
    *   Add lagged columns for time-series analysis.
    *   Convert column data types (numeric, datetime, string).
    *   Split text columns by delimiter.
    *   Extract patterns from text columns using regex.
    *   Clean text columns (strip, lower, upper, remove digits/punctuation).
    *   Perform VLOOKUP-like operations to merge data from other files.
    *   Concatenate (stack) data from multiple files/sheets.
*   **Interactive Plotting**: Generate various chart types (line, bar, scatter, histogram, box, pie, radar) and view them directly within the application.
*   **Plot Export**: Export generated plots to your desired location as PNG or JPEG images.
*   **Intuitive GUI**: A clean, modern user interface built with PyQt6, featuring:
    *   Combined query input and file selection.
    *   Tabbed output for "Results" (text and tables) and "Plots".
    *   Resizable output panels using `QSplitter`.
    *   Real-time status updates via a status bar.
    *   Application opens in maximized mode for optimal workspace.
*   **Error Handling & Feedback**: Provides clear messages for successful operations, warnings, and errors.

## Installation

Follow these steps to set up and run the Excel AI Agent on your local machine.

### Prerequisites

*   Python 3.8 or higher
*   `pip` (Python package installer)

### Steps

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/your-username/excel-ai-agent.git
    cd excel-ai-agent
    ```
    (Replace `your-username` with the actual GitHub username if this is a public repository.)

2.  **Create a virtual environment (recommended):**
    ```bash
    python -m venv venv
    ```

3.  **Activate the virtual environment:**
    *   **Windows:**
        ```bash
        .\venv\Scripts\activate
        ```
    *   **macOS/Linux:**
        ```bash
        source venv/bin/activate
        ```

4.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```
    This will install all necessary libraries, including `PyQt6`, `pandas`, `matplotlib`, `seaborn`, and `groq`.

5.  **Set up your Groq API Key:**
    The application uses the Groq API for its LLM capabilities. You need to obtain an API key from [Groq Cloud](https://console.groq.com/keys).

    Create a file named `.env` in the root directory of your project (the same directory as `requirements.txt`) and add your API key:
    ```
    GROQ_API_KEY="your_groq_api_key_here"
    # Optional: Specify a different Groq model if available and desired
    # GROQ_MODEL_NAME="llama-3.1-70b-versatile"
    ```
    Replace `"your_groq_api_key_here"` with your actual Groq API key.

6.  **Prepare Icons (Optional but Recommended for full UI experience):**
    The GUI uses icons for buttons and the application window. Please ensure you have the following `.png` files in the `src/gui/icons/` and `src/gui/images/` directories:
    *   `src/gui/icons/app_icon.png` (for the main window icon)
    *   `src/gui/icons/play.png` (for the "Process Query" button)
    *   `src/gui/icons/save.png` (for the "Export Plot" button)
    *   `src/gui/images/file_upload.png` (for the file upload icon in the query input)

    If these files are missing, the application will still run, but the icons will not appear.

## Usage

To launch the application, run the `app.py` file from the `src/gui` directory:

