# SQL to Excel Automation ( Python)

## 📌 Overview
This repository contains two sample automation scripts (Java and Python) that:
1. Fetch data from a SQL database.
2. Write the data into an Excel sheet.
3. Send the Excel file via Outlook email to specified recipients.

---

## 🐍 Python Script

### Features
- Uses `pyodbc` to connect to SQL Server.
- Uses `pandas` to export query results into Excel.
- Uses `win32com.client` to send email via Outlook.

### Requirements
- Python 3.x
- Libraries:
  ```bash
  pip install pyodbc pandas pywin32
