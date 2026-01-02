
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
