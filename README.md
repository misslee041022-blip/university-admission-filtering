# 🎓 University Admission Filtering & Simulation System
(대입 수능 최저학력기준 분석 및 등급 시뮬레이터)

## 📌 Project Overview
This project is a Python-based application that filters university departments based on CSAT (Suneung) scores.
It goes beyond simple filtering by providing a **"What-if Simulation"** feature to analyze potential admission chances based on grade changes and generates **formatted Excel reports**.

## ✨ Key Features (주요 기능)
1.  **Data Loading**: Supports `.xlsx` and `.csv` files containing university admission data.
2.  **Basic Filtering**: Filters departments based on user's current CSAT grades (Korean, Math, English, History, Exploration).
3.  **Detailed Filtering**: Filter results dynamically by Category (Humanities/Natural), University, and Admission Type.
4.  **📈 Grade Simulation (What-if Analysis)**:
    * Simulates scenarios where grades improve or decline (e.g., English +1 grade).
    * Provides a **Comparative Analysis (Delta)**: Shows newly added departments vs. removed departments.
    * Visualizes results using Tabs (Added/Removed).
5.  **📊 Excel Reporting**: Exports analysis results into a clean, formatted Excel file using the `xlsxwriter` engine.

## 🛠️ Tech Stack & Modules
* **Python 3.9+**
* **Pandas**: Data processing and filtering engine.
* **Openpyxl**: Engine for reading Excel database files.
* **XlsxWriter**: Engine for generating formatted Excel reports.
* **Tkinter**: GUI (Graphical User Interface) implementation.

## 🚀 How to Run
1.  Install dependencies:
    ```bash
    pip install pandas openpyxl xlsxwriter
    ```
2.  Run the application:
    ```bash
    python main.py
    ```
3.  Load your data file (`data.xlsx`) and input your scores.

## 👨‍💻 Developer Note
This program was developed to help students establish strategic admission plans by visualizing data-driven possibilities.
