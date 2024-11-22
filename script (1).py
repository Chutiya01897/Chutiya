import time
import openpyxl
import os
from urllib.parse import quote_plus

# Path to your contacts Excel file
file_path = "contacts.xlsx"  # Ensure this file is in the same directory as your script

# Open the Excel file
workbook = openpyxl.load_workbook(file_path)
sheet = workbook.active

# Extract contacts (Phone Number and Name)
contacts = []
for row in sheet.iter_rows(min_row=2, max_col=2, values_only=True):
    phone_number = str(row[0])  # Phone number (make sure it's a string)
    name = row[1]  # Contact name
    contacts.append((phone_number, name))

# Define a function to send message to WhatsApp via Intent
def send_whatsapp_message(phone_number, name):
    # Construct the WhatsApp API URL
    message = f"Hello {name}, this is an automated message!"  # Customize your message here
    url = f"https://api.whatsapp.com/send?phone={phone_number}&text={quote_plus(message)}"
    
    # Use Termux's am command to open the WhatsApp URL directly in WhatsApp
    # The command 'am start' will open the URL within the app
    command = f"am start -a android.intent.action.VIEW -d '{url}'"
    os.system(command)
    print(f"Sent message to {name} ({phone_number}) via WhatsApp")

# Infinite loop to send messages to all contacts
while True:
    for phone_number, name in contacts:
        send_whatsapp_message(phone_number, name)
        time.sleep(5)  # Wait for 5 seconds before sending the next message