# 📊 Test Dial Extraction Automation using Python & OpenPyXL

This project automates the process of finding and highlighting **Test dial entries** from a large detailed Excel report.  
It cross-checks names or identifiers between two Excel sheets — one containing target names and another containing the detailed data — and highlights all matches automatically.

---

## 🚀 Key Features

- 📂 **Two-File Comparison:** Compares data from two Excel files (source & detailed report)
- 🟧 **Automatic Highlighting:** Highlights matching entries in orange
- ⚙️ **OpenPyXL Powered:** Uses the OpenPyXL library for Excel reading and writing
- 🔁 **Fast & Scalable:** Handles large datasets efficiently
- 🧠 **Simple Logic:** Clean and easily customizable script

---

## 🧠 How It Works

1. The script reads a list of PPT names (or identifiers) from the first Excel file (`Sheet1`).  
2. It loads the second Excel file (the detailed report) from `Sheet2`.  
3. For every name found in the list, it searches through the detailed report.  
4. When a match is found, the corresponding cell is **highlighted in orange**.  
5. The updated report is saved automatically.

---

## 👨‍💻 Author

Ahmed Essam
Automation Developer | Python & Excel Automation

