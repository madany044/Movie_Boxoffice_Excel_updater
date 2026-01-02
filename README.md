# Movie Boxoffice Excel Updater 📊

A Python-based automation tool that converts real-world movie show updates into a structured Excel analytics report.

This project is designed for **personal use** to eliminate manual Excel work when handling frequently changing, multi-city movie show data.

---

## ✨ What This Tool Does

- Accepts **raw update files** containing multiple cities
- Handles **non-standard / messy input formats**
- Automatically generates:
  - City-wise Excel sheets
  - A consolidated Summary sheet
  - A highlighted GRAND TOTAL row
  - Charts for quick insights
- Creates **daily backups** before every update
- Updates everything with **one command**

---
## 🛠 Tech Stack

- Python 3

- Pandas

- OpenPyXL

## 📄 Input Format

The input file can contain **multiple city sections**, each with its own JSON block.

Example:

--- Bengaluru ---

 { JSON block }

--- Mysuru (Mysore) ---

 { JSON block }



✔ No need to modify the incoming format  
✔ Each city can have any number of theaters  
✔ Updates can change every time  

---

## 📊 Excel Output

After running the script, the Excel file will contain:

### 📄 City-wise Sheets
- One sheet per city
- Complete show-level data

### 📄 Summary Sheet
For each city:
- Total venues
- Total shows
- Total seats
- Sold seats
- Available seats
- Total gross
- Average occupancy

### ⭐ GRAND TOTAL
- One highlighted row combining **all cities**
- Overall totals for seats, sales, and gross

### 📈 Charts
- Automatically generated charts (e.g., Total Gross by City)

### 📦 Backups
- Every run creates a timestamped backup
- Stored safely in `excel/backups/`

---

## 🚀 How to Use

### 1️⃣ Install dependencies
```bash
pip install -r requirements.txt
```
## 2️⃣Place update file
```
data/latest_update.txt
```

## 3️⃣ Run the script
```
python scripts/update_excel.py
```

<div align="center">
  
## Designed & Developed By 
[ MADAN Y ]

 **Email**: madanmadany2004@gmail.com 

</div>
