import pyodbc
import pandas as pd
import win32com.client as win32

# Step 1: Fetch data from SQL
conn = pyodbc.connect("Driver={SQL Server};Server=YOUR_SERVER;Database=YOUR_DB;Trusted_Connection=yes;")
query = "SELECT * FROM your_table"
df = pd.read_sql(query, conn)

# Step 2: Save data to Excel
excel_file = "output.xlsx"
df.to_excel(excel_file, index=False)

# Step 3: Send via Outlook
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = "recipient@example.com"
mail.Subject = "Automated SQL Report"
mail.Body = "Hi,\n\nPlease find the attached report.\n\nRegards,\nAutomation Script"
mail.Attachments.Add(excel_file)
mail.Send()

print("Email sent successfully!")
