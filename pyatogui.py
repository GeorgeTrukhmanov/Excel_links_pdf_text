import webbrowser
import pyautogui as pg
import time
import pandas as pd
import pdfplumber
import requests
import io
import os
import pyperclip
excel_file_path = 'Data_sample.xlsx'


def read_links_from_excel(file_path):
    df = pd.read_excel(file_path)
    return df[df.columns[1]].tolist()


def download_pdf(url):
    response = requests.get(url)
    response.raise_for_status()
    return response.content


def extract_text_from_pdf(pdf_content):
    text = ''
    with pdfplumber.open(io.BytesIO(pdf_content)) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
    if page_text:
        text += page_text + '\n'
    return text

# Чтение ссылок из Excel файла
links = read_links_from_excel(excel_file_path)
#print(pg.position())


for url in links:
    print(url)
    webbrowser.open(f'{url}', new=2)
    time.sleep(3)
    pg.hotkey("ctrl", "s")
    time.sleep(3)
    pg.hotkey("ctrl", "c")
    nazv = pyperclip.paste()
    print(nazv)
    pg.leftClick(727, 562, duration=0.2)
    time.sleep(1)
    short_name = url[72:]
    mega_short = short_name[:10]
    print(mega_short)
    try:
        os.rename(f'C:/Users/truhm/Работа/pdf/{nazv}.pdf', f'C:/Users/truhm/Работа/pdf/{mega_short}.pdf')
    except Exception as e:
        print(f'Произошла ошибка: {e}')
#    write_text_to_pdf(f'C:/Users/truhm/Работа/pdf/{nazv}.pdf', {url})
    pg.hotkey("ctrl", "w")