import pandas as pd
import smtplib
from email.message import EmailMessage
import time 

# --- YOUR DETAILS ---
# SENDER_EMAIL = "your_email@gmail.com"
SENDER_EMAIL = "christopherkoochinfong@gmail.com"
# Use a Google App Password (16 characters), NOT your regular password
# Use the link https://myaccount.google.com/apppasswords to generate your App password
SENDER_PASSWORD = "xx" 
EXCEL_FILE = 'Test1.xlsx'

def send_bulk_emails():
    try:
        # Change header=1 to header=0 if your column names are on the VERY FIRST row
        # Change header=1 to header=2 if there are two title rows above your column names
        data = pd.read_excel(EXCEL_FILE, sheet_name='2026 Tracker', header=2)
        
        # This will clean up any extra spaces automatically
        data.columns = data.columns.str.strip()

        # Let's check what it finds now
        print("Now seeing these columns:", data.columns.tolist())

        if 'Email' not in data.columns:
            print("Still can't find 'Email'. Try changing header=2 to header=0 in the code.")
            return

        data = data.dropna(subset=['Email']) 

        # 4. Connect to Gmail
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            print("Successfully logged into Gmail!")

            for index, row in data.iterrows():
                # Using .get() prevents the script from crashing if a cell is empty
                name = str(row.get('Name', 'Friend')) 
                recipient = str(row['Email']).strip()

                msg = EmailMessage()
                msg['Subject'] = f"Quick Update for {name}"
                msg['From'] = SENDER_EMAIL
                msg['To'] = recipient
                
                body = f"Hey {name},\n\nI hope you're having a great day! I'm testing out a new automation tool for Mecha Mayhem.\n\nBest,\nChristopher"
                msg.set_content(body)

                server.send_message(msg)
                print(f"Email sent to {name} ({recipient})")
                
                # Pause for 2 seconds (Anti-Spam)
                time.sleep(2)

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    send_bulk_emails()