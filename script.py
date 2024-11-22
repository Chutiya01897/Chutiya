import time
import pyautogui
import webbrowser
import openpyxl

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

# Loop through contacts and send messages
while True:
    for phone_number, name in contacts:
        # Construct the WhatsApp API URL with the contact's phone number and personalized message
        url = f"https://api.whatsapp.com/send?phone={phone_number}&text=Hello%20{name}%2C%20this%20is%20an%20automated%20message!"

        # Open the WhatsApp link in the default web browser
        print(f"Sending message to {name} ({phone_number})...")
        webbrowser.open(url)

        # Wait for 5 seconds to let the page load
        time.sleep(5)