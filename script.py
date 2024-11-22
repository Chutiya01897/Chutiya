import time
import openpyxl
import os
from urllib.parse import quote_plus

# Path to your contacts Excel file
file_path = "contacts.xlsx"  # Ensure this file is in the same directory as your script

# Path to your file that you want to send (ensure the file is publicly accessible)
file_url = "VAHAN_PARIVAHAN.apk"  # Replace this with your actual file URL

# Open the Excel file
workbook = openpyxl.load_workbook(file_path)
sheet = workbook.active

# Extract contacts (Phone Number, Name, and Vehicle Number)
contacts = []
for row in sheet.iter_rows(min_row=2, max_col=3, values_only=True):  # Assuming 3 columns: Phone Number, Name, Vehicle Number
    phone_number = str(row[0])  # Phone number (make sure it's a string)
    name = row[1]  # Contact name
    contacts.append((phone_number, name))

# Define a function to send message with file to WhatsApp via Intent
def send_whatsapp_message(phone_number, name):
    # Construct the custom message with caption
    message = (
        f"Dear Sir/Madam,\n\n"
        f"We would like to inform you that a traffic ticket (No. MH46894230933070073) has been issued for your vehicle, "
        f"registered under the number *{name}*. The evidence for this ticket was captured by our CCTV cameras.\n\n"
        f"If you believe this ticket has been issued in error, we encourage you to download the Vahan Parivahan app using the link below. "
        f"This app will allow you to review the evidence in detail and report any discrepancies.\n\n"
        f"[Download Vahan Parivahan App]\n\n"
        f"Thank you for your cooperation.\n\n"
    )
    
    # Construct the WhatsApp API URL with file attachment
    url = f"https://api.whatsapp.com/send?phone={phone_number}&text={quote_plus(message)}&attachment={quote_plus(file_url)}"
    
    # Use Termux's am command to open the WhatsApp URL directly in WhatsApp
    # The command 'am start' will open the URL within the app
    command = f"am start -a android.intent.action.VIEW -d '{url}'"
    os.system(command)
    print(f"Sent message to {name} ({phone_number}) with vehicle {vehicle_number} and file via WhatsApp")

# Infinite loop to send messages to all contacts
while True:
    for phone_number, name in contacts:
        send_whatsapp_message(phone_number, name)
        time.sleep(5)  # Wait for 5 seconds before sending the next message