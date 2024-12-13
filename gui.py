import json
import gspread
import pyautogui
import pyperclip
import time

def send_messages(data, imessage):
    # Open WhatsApp Desktop App
    pyautogui.hotkey('win', 's')  # Open Windows search
    time.sleep(1)
    pyautogui.write('WhatsApp')  # Type WhatsApp in the search bar
    pyautogui.press('enter')  # Open WhatsApp
    time.sleep(5)  # Wait for WhatsApp to open completely

    for sadhak in data:
        name = sadhak.get('Name').split(' ')[0]
        contact = '+91 ' + str(sadhak.get('Contact'))
        message = imessage.format(name, name)

        # Search for the contact
        pyautogui.click(200, 100)  # Coordinates for the search box (adjust as per your screen)
        time.sleep(1)
        pyperclip.copy(contact)  # Copy contact to clipboard
        pyautogui.hotkey('ctrl', 'v')  # Paste contact
        time.sleep(1)
        pyautogui.press('enter')  # Open chat with the contact
        time.sleep(2)

        # Type and send the message
        pyperclip.copy(message)  # Copy message to clipboard
        pyautogui.hotkey('ctrl', 'v')  # Paste message
        time.sleep(1)
        pyautogui.press('enter')  # Send the message
        time.sleep(1)

        # Add a small pause to ensure smooth operation
        time.sleep(2)

def data_extract(sheet_url):
    # Load message template from JSON
    with open("message.json", encoding="utf-8") as file:
        data = json.load(file)

    greeting = data["greeting"]
    message = data["message"]
    imessage = f"{greeting}\n\n{message}\n\n🌸समर्पण आश्रम अरड़का🌸"

    # Access Google Sheet
    gc = gspread.service_account(filename=r"C:\Users\WFH\Desktop\WhaBot\service_account.json")
    gSheet = gc.open_by_url(sheet_url)
    worksheet = gSheet.worksheet('test')
    data = worksheet.get_all_records()

    # Send messages
    send_messages(data, imessage)

if __name__ == "__main__":
    sheet_url = "https://docs.google.com/spreadsheets/d/14M3lwPv2woPgqbGWASJV_R6z11Rj7TpCMejo7lv8o5M/edit?usp=sharing"
    data_extract(sheet_url)
