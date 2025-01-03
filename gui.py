import json
import gspread
import pyautogui
import pyperclip
import time
from datetime import datetime

class WeddingInvitation:

    def __init__(self, sheet_url, service_account_path):
        self.sheet_url = sheet_url
        self.service_account_path = service_account_path

    def send_messages(self, data, message):
        # Open WhatsApp Desktop App
        pyautogui.hotkey('win', 's')  # Open Windows search
        time.sleep(1)
        pyautogui.write('WhatsApp')  # Type WhatsApp in the search bar
        pyautogui.press('enter')  # Open WhatsApp
        time.sleep(5)  # Wait for WhatsApp to open completely

        # date = datetime.today()

        for sadhak in data:
            name = sadhak.get('Name').split(' ')[0]
            contact = str(sadhak.get('Contact'))
            pyperclip.copy(contact)  # Copy contact to clipboard

            # Search for the contact
            pyautogui.click(200, 100)  # Coordinates for the search box (adjust as per your screen)
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'v')  # Paste contact
            time.sleep(1)
            pyautogui.press('down')  # Navigate to the first result
            pyautogui.press('enter')  # Open chat with the contact
            time.sleep(1)

            # Type and send the message
            pyperclip.copy(message)  # Copy message to clipboard
            pyautogui.hotkey('ctrl', 'v')  # Paste message
            time.sleep(1)
            pyautogui.press('enter')  # Select the first result
            time.sleep(2) # Send the message

            # Add a small pause to ensure smooth operation
            time.sleep(2)

    def extract_data(self):
        # Load message template from JSON
        with open("wedding_message.json", encoding="utf-8") as file:
            data = json.load(file)

        message = data["message"]

        # Access Google Sheet
        gc = gspread.service_account(filename=self.service_account_path)
        gSheet = gc.open_by_url(self.sheet_url)
        worksheet = gSheet.worksheet('wedding')
        data = worksheet.get_all_records()

        # Send messages
        self.send_messages(data, message)

if __name__ == "__main__":
    sheet_url = "https://docs.google.com/spreadsheets/d/14M3lwPv2woPgqbGWASJV_R6z11Rj7TpCMejo7lv8o5M/edit?usp=sharing"
    service_account_path = "E:\\Development\\WhatsBot\\service_account.json"
    wedding_invitation = WeddingInvitation(sheet_url, service_account_path)
    wedding_invitation.extract_data()
