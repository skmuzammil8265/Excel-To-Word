# Excel-To-Word
A simple Python desktop app that generates multiple Word documents from a template by replacing placeholders with data from an Excel sheet. Built with Tkinter, python-docx, and openpyxl.
# 📝 Excel to Word Generator (Tkinter App)

This is a Python desktop application built with **Tkinter** that allows users to generate multiple Word documents from a single Word template by replacing placeholders with data from an Excel spreadsheet.

## 🔧 Features

- GUI-based tool using Tkinter
- Reads data from Excel (`.xlsx`) files
- Replaces placeholders like `{Name}`, `{Date}`, etc. in a Word (`.docx`) template
- Saves each customized document separately
- Automatic file naming and conflict resolution

## 📂 How to Use

1. Run `App.py` with Python.
2. Select an Excel file and a Word template using the GUI.
3. Click the **"Generate Word Files"** button.
4. Generated files will appear in the `output/` folder.

## 🧠 Requirements

- Python 3.x
- `python-docx`
- `openpyxl`

Install dependencies with:

```bash
pip install python-docx openpyxl
