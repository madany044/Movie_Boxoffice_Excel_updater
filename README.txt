# Movie Excel Updater 📊

A Python automation tool that converts raw, real-world movie show updates into a structured Excel analytics report.

## ✨ Features
- Parses messy multi-city update text
- Generates city-wise Excel sheets
- Auto-creates summary & grand total
- Builds charts automatically
- Creates daily Excel backups
- Zero manual Excel work

## 📁 Input Format
The tool accepts update files like:

--- Bengaluru ---
{ JSON block }

--- Mysuru ---
{ JSON block }

(No format changes required.)

## 🚀 How to Use

1. Install dependencies
```bash
pip install -r requirements.txt

2.Place update file in
```bash
data/latest_update.txt

3. Run
```bash
python scripts/update_excel.py
