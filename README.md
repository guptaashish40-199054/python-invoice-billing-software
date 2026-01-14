Invoice Generator Application (Python)

A professional GUI-based Invoice & Billing Application built using Python to generate GST-compliant invoices in PDF format for small businesses.

🚀 Features

GUI-based Invoice Generator (Tkinter)

Vendor (Company) Details

Customer Details (Company, Name, Address, Phone)

Multiple Item Billing

Automatic Invoice Number Generation

CGST + SGST Calculation

Professional PDF Invoice Generation

Customer Database (Excel)

Invoice History (Excel)

Company Logo Support

Windows Executable (.exe)

🛠️ Technologies Used

Python

Tkinter

ReportLab

OpenPyXL

PyInstaller

📂 Project Structure
invoice_app/
│
├── invoice_gui.py
├── company_logo.png
├── customers.xlsx
├── invoice_history.xlsx
├── invoices/

▶️ How to Run the Project

Install required libraries:

pip install reportlab openpyxl


Run the application:

python invoice_gui.py

🪟 Create Windows EXE
pyinstaller --onefile --windowed invoice_gui.py

📸 Screenshots

(Add screenshots of GUI and generated PDF here)

👨‍💻 Author

Ashish Gupta
