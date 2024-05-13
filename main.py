import pywhatkit
import gspread
from datetime import datetime


def send_message(name, message, contact, hour, minute):
    now = datetime.now()

    pywhatkit.sendwhatmsg_instantly(contact, f"Name: {name}\nMessage: {message}", wait_time=6)


def data_extract(sheet_url):

    message = "Jai Baba Swami {}"
    gc = gspread.service_account()
    gSheet = gc.open_by_url(sheet_url)

    worksheet = gSheet.worksheet('Whatsapp')
    data = worksheet.get_all_records()
    
    for sadhak in data:
        name = sadhak.get('Name')
        contact = '+91 ' + str(sadhak.get('Contact'))

    send_message(name, message.format(name), contact, 23, 0)

if __name__ == "__main__":
    sheet_url = "https://docs.google.com/spreadsheets/d/1Ma3EFn2vsULjtG9koy_y-tHhDbTyU4Z-WszUWzBvkTk/edit#gid=1075867701"
    data_extract(sheet_url)