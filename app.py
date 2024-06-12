"""
Automation of messages so that customers can contact the supplier, find out prices and even send billing messages on a specific day to different customers
"""
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

webbrowser.open('https://web.whatsapp.com/')
sleep(30)
# Read spreadsheet and save information about name, telephone number and due date
workbook = openpyxl.load_workbook('clients.xlsx')
page_clients = workbook['Página1']

for line in page_clients.iter_rows(min_row=2):
    #name, phone, due date
    name = line[0].value
    phone = line[1].value
    due_date = line[2].value
    message = f'message = Your bill expires on {due_date.strftime("%d/%m/%Y")}. Please pay through the link https://www.link_pagamento.com'

# Create personalized Whatsapp links and send messages to each customer based on spreadsheet dataworkbook
    link_message_whatsapp = f'https://web.whatsapp.com/send?phone={phone}&text={quote(message)}'
    webbrowser.open(link_message_whatsapp)
    sleep(10)
    try:
        # Find the arrow image and click
        arrow = pyautogui.locateCenterOnScreen('arrow.png')
        sleep(5)
        pyautogui.click(arrow[0],arrow[1])
        sleep(5)
        pyautogui.hotkey('ctrl','w')
        sleep(5)
    except:
        print(f'Was not possible to send message to {name}')
        with open('error.csv','a',newline='',encoding='utf-8') as archieve:
            archieve.write(f'{name},{phone}')
    


