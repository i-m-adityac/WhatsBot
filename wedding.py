import json
import gspread
from playwright.sync_api import sync_playwright
import pandas as pd
import time

class Gurukarya:
    def send_messages(self, data, imessage):
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False)
            page = browser.new_page()
            page.goto("https://web.whatsapp.com")
            
            print("Please scan the QR code on WhatsApp Web and press Enter after login.")
            input()

            # print(data)
            # import pdb;pdb.set_trace()
            date = data[0].get('Date', '')
            for sadhak in data:
                name = sadhak.get('Name').split(' ')[0]
                contact = '+91 ' + str(sadhak.get('Contact'))
                message = imessage.format(name, name, date)

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

    def data_extract(self, sheet_url):
        with open("message.json", encoding="utf-8") as file:
            data = json.load(file)

        greeting = data["greeting"]
        message = data["message"]
        imessage = f"{greeting}\n\n{message}\n\n🌸समर्पण आश्रम अरड़का🌸"

        gc = gspread.service_account(filename=r"E:\Development\WhatsBot\service_account.json")
        gSheet = gc.open_by_url(sheet_url)
        worksheet = gSheet.worksheet('test')
        data = worksheet.get_all_records()
        
        self.send_messages(data, imessage)


class WeddingInvitation:

    def create_broadcast(self, data):
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False, channel="msedge")
            context = browser.new_context(storage_state="whatsapp_state.json") 
            page = context.new_page()

            page.goto("https://web.whatsapp.com")
            
            print("Please scan the QR code on WhatsApp Web and press Enter after login.")
            input()

            import pdb;pdb.set_trace()

            page.click('div[title="New chat"]')  # Click on the "New chat" button
            page.click('div[title="New broadcast"]')  # Click on "New broadcast"

            print(data)
            for sadhak in data:
                contact = '+91 ' + str(sadhak.get('Contact'))
                search_box = page.wait_for_selector("div[contenteditable='true'][data-tab='3']")
                search_box.click()
                search_box.fill(contact)
                page.keyboard.press("Enter")
                time.sleep(1)  # Ensure the contact is added to the broadcast

            # Press enter to create the broadcast
            page.keyboard.press("Enter")

            time.sleep(5)
            # input()
            context.storage_state(path="whatsapp_state.json")
            browser.close()


    def wedding_message(self, data, message):
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False, channel="msedge")
            try:
                context = browser.new_context(storage_state="whatsapp_state.json") 
                page = context.new_page()
            except:
                page = browser.new_page()
            page.goto("https://web.whatsapp.com")
            
            print("Please scan the QR code on WhatsApp Web and press Enter after login.")
            input()

            print(data)
            for sadhak in data:
                contact = '+91 ' + str(sadhak.get('Contact'))

                # Search for the contact number
                search_box = page.wait_for_selector("div[contenteditable='true'][data-tab='3']")
                search_box.click()
                search_box.fill(contact)
                
                page.keyboard.press("Enter")

                # Wait for the chat to open and verify the contact
                message_box = page.wait_for_selector("div[contenteditable='true'][data-tab='10']")

                message_box.click()
                message_box.fill(message)

                video_path = r'E:\Development\WhatsBot\wed_vid.mp4'  # Provide the path to your video

                attach_button = page.locator('button[aria-label="Attach"]')  # XPath for attachment button
                attach_button.click()
                
                # Select the file input element and upload file
                file_input = page.locator('input[type="file"][accept*="image/"][accept*="video/"]')
                file_input.set_input_files(video_path)

                page.wait_for_timeout(2000)

                page.keyboard.press("Enter")

            time.sleep(5)
            # input()
            context.storage_state(path="whatsapp_state.json")
            browser.close()


    def wedding_invitation(self, sheet_url):
        with open("wedding_message.json", encoding="utf-8") as file:
            data = json.load(file)

        message = data["message"]

        gc = gspread.service_account(filename=r"E:\Development\WhatsBot\service_account.json")
        gSheet = gc.open_by_url(sheet_url)
        worksheet = gSheet.worksheet('wedding')
        data = worksheet.get_all_records()

        # Function to send messages (to be defined)
        # self.wedding_message(data, message)
        self.create_broadcast(data)

if __name__ == "__main__":
    sheet_url = "https://docs.google.com/spreadsheets/d/14M3lwPv2woPgqbGWASJV_R6z11Rj7TpCMejo7lv8o5M/edit?usp=sharing"
    # data_extract(sheet_url)
    weeding_sheet = "https://interrasystems-my.sharepoint.com/:x:/p/sakshij_in/EVkYqjch-tdKvSAWla6IZIQBGCo77kbwX89ktwnu6aBKmw?e=kpTiDQ"
    wdObj = WeddingInvitation()
    wdObj.wedding_invitation(sheet_url)
