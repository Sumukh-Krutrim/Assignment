# README

Overview

This project consists of two Python scripts, `data.py` and `plott.py`, which work together to process and visualize remark statistics from an Excel file. 

1. `data.py`

- Purpose: Reads data from an input Excel file and creates a new Excel file (`Statistics.xlsx`) with processed statistics.
- Functionality:
  - Reads data from the specified Excel file for various languages.
  - Processes remarks data for different models and computes frequencies and percentages.
  - Writes the processed data to a new Excel file, `Statistics.xlsx`.

2. `plott.py`

- Purpose: Visualizes the processed data from `Statistics.xlsx` using Dash and Plotly.
- Functionality:
  - Initializes a Dash app with dropdowns for selecting language, model, and remark.
  - Generates a bar chart showing the percentage of remarks based on the selected criteria.
  - Updates dropdown options and the graph based on user selections.

-> Usage

Prerequisites

- Python environment with the following libraries installed:
  - pandas
  - plotly
  - dash
  - openpyxl (for `data.py`)

Steps to Run

1. Run `data.py`
   - Ensure the input Excel file is at the specified path.
   - This script processes the data and creates a new file named `Statistics.xlsx` in the current directory.
   - Command: `python data.py`
   
2. Run `plott.py`
   - This script uses the processed data in `Statistics.xlsx` to visualize the statistics.
   - Command: `python plott.py`

Notes

- Ensure `data.py` runs successfully to generate `Statistics.xlsx` before running `plott.py`.
- `plott.py` relies on the structure and content of `Statistics.xlsx` created by `data.py`.

Conclusion

This project provides an interactive way to visualize remark statistics using Dash and Plotly. By following the steps above, you can explore different languages, models, and remarks through a user-friendly interface.

