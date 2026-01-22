import openpyxl
from datetime import datetime
import pandas as pd
import smtplib
from email.message import EmailMessage
import time

# --- YOUR DETAILS ---
# SENDER_EMAIL = "your_email@gmail.com"
SENDER_EMAIL = "christopherkoochinfong@gmail.com"
# Use a Google App Password (16 characters), NOT your regular password
# Use the link https://myaccount.google.com/apppasswords to generate your App password
SENDER_PASSWORD = "xxxx xxxx xxxx" 
EXCEL_FILE = 'Test1.xlsx'

def send_bulk_emails():
    try:
        # 1. Read data with pandas 
        data = pd.read_excel(EXCEL_FILE, sheet_name='2026 Tracker', header=2)
        data.columns = data.columns.str.strip()
        data = data.dropna(subset=['Email'])

        # 2. Open the actual workbook
        wb = openpyxl.load_workbook(EXCEL_FILE)
        sheet = wb['2026 Tracker']

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            
            for index, row in data.iterrows():
                name = str(row.get('First Name', 'Friend'))
                recipient = str(row['Email']).strip()

                # Send Email
                msg = EmailMessage()
                msg['Subject'] = f"Quick Update for {name}"
                msg['From'] = SENDER_EMAIL
                msg['To'] = recipient
                msg.set_content(f"Hey {name},\n\nTesting the new safe-save logic!")
                server.send_message(msg)
                print(f"Sent to {name}")

                excel_row = index + 4 
                sheet.cell(row=excel_row, column=6).value = datetime.now().strftime('%Y-%m-%d')
                
                time.sleep(2)

        # 4. Save the Workbook safely
        wb.save(EXCEL_FILE)
        print("Excel updated safely without messing up formatting!")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    send_bulk_emails()