import json
import gspread
from playwright.sync_api import sync_playwright

def send_messages(data, imessage):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        page.goto("https://web.whatsapp.com")
        
        print("Please scan the QR code on WhatsApp Web and press Enter after login.")
        input()

        for sadhak in data:
            name = sadhak.get('Name').split(' ')[0]
            contact = '+91 ' + str(sadhak.get('Contact'))
            message = imessage.format(name)

            # Search for the contact number
            search_box = page.wait_for_selector("div[contenteditable='true'][data-tab='3']")
            search_box.click()
            search_box.fill(contact)
            
            page.keyboard.press("Enter")

            # Wait for the chat to open and verify the contact
            message_box = page.wait_for_selector("div[contenteditable='true'][data-tab='10']")

            message_box.click()
            message_box.fill(message)
            page.keyboard.press("Enter")
        
        browser.close()

def data_extract(sheet_url):
    with open("message.json", encoding="utf-8") as file:
        data = json.load(file)

    greeting = data["greeting"]
    message = data["message"]
    imessage = f"{greeting}\n\n{message}\n\n🌸गुरु पूर्णिमा समिति🌸"

    gc = gspread.service_account(filename=r"C:\Aditya\Study\Development\WhatsBot\service_account.json")
    gSheet = gc.open_by_url(sheet_url)
    worksheet = gSheet.worksheet('Whatsapp')
    data = worksheet.get_all_records()
    
    send_messages(data, imessage)

if __name__ == "__main__":
    sheet_url = "https://docs.google.com/spreadsheets/d/1Ma3EFn2vsULjtG9koy_y-tHhDbTyU4Z-WszUWzBvkTk/edit#gid=1075867701"
    data_extract(sheet_url)
