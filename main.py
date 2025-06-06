import json
import time
import gspread
from playwright.sync_api import sync_playwright

def send_messages(data, messages):
    skipped_contacts = []  # List to keep track of skipped contacts

    try:
        with sync_playwright() as p:
            user_data_dir = r"C:\Users\MSI\AppData\Local\Google\Chrome\User Data"
            # Launch browser using Chromium, edge or chrome depending on your setup
            browser = p.chromium.launch_persistent_context(
                user_data_dir=user_data_dir,
                headless=False,  # Headless mode is False for debugging purposes
                channel="chrome",  # Use "chrome" or "msedge" depending on your browser
            )
            page = browser.new_page()
            page.goto("https://web.whatsapp.com")
            time.sleep(10)  # Adjust if necessary based on your connection speed
            # import pdb;pdb.set_trace()

            # Wait for the chat list to load, signaling WhatsApp Web is ready
            page.wait_for_selector("div[aria-label='Chat list']", timeout=10000)

            for sadhak in data:
                try:
                    name = sadhak.get('Name')
                    contact = '+91 ' + str(sadhak.get('Contact')).lstrip()

                    # Search for the contact number
                    search_box = page.wait_for_selector("div[contenteditable='true'][data-tab='3']", timeout=5000)
                    search_box.click()
                    search_box.fill(contact)
                    time.sleep(0.5)  # Adjust if necessary based on your connection speed
                    page.keyboard.press("Enter")
                    time.sleep(1)

                    # Send each message in the list
                    for key, msg in messages.items():
                        formatted_message = msg.format(name)  # Format the message with the recipient's first name

                        # Wait for the message input box and fill the message
                        message_box = page.wait_for_selector("div[contenteditable='true'][data-tab='10']", timeout=5000)
                        message_box.click()
                        message_box.fill(formatted_message)
                        time.sleep(0.1)
                        page.keyboard.press("Enter")
                        time.sleep(1)

                    print(f"Message sent to {contact} ({name})")
                except Exception as e:
                    print(f"Error while sending message to {contact}: {e}")
                    skipped_contacts.append(contact)  # Add to skipped list in case of error
            time.sleep(5) 
            page.close()
            browser.close()  # Close the browser once done
    except Exception as e:
        print(f"An error occurred: {e}")

def data_extract(sheet_url):
    with open("message_single.json", encoding="utf-8") as file:
        messages = json.load(file)  # Load messages from JSON

    gc = gspread.service_account(filename=r"service_account.json")
    gSheet = gc.open_by_url(sheet_url)
    worksheet = gSheet.worksheet('test')  # Adjust the sheet name if necessary
    data = worksheet.get_all_records()  # Fetch all records

    send_messages(data, messages)

if __name__ == "__main__":
    sheet_url = "https://docs.google.com/spreadsheets/d/14M3lwPv2woPgqbGWASJV_R6z11Rj7TpCMejo7lv8o5M/edit?usp=sharing"
    data_extract(sheet_url)
