import pdfplumber
import os
import pandas as pd
import re

directory = 'C://Users//truhm//Работа//pdf'  # папка с pdf-файлами после скачивания
file_p = 'C://Users//truhm//Работа//Data_sample.xlsx'  # файл с ссылками


def find_keywords(text):
    words = ['определил', 'решил']
    new_words = []

    for word in words:
        new_word = '\s*'.join([letter for letter in word])
        new_words.append(new_word)

    pattern = '|'.join(new_words)

    matches = [match for match in re.finditer(pattern, text, re.IGNORECASE)]

    if not matches:
        return 'Нет ключевых слов в тексте'
    else:
        last_match = matches[-1]
        return text[last_match.end():].strip()


def extract_text_from_pdf(file_path):
    t = ''
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:  # проверяем, есть ли текст на странице
                t += page_text + '\n'
    return t



def insert_text_to_excel(file_path, case, text_ins):
    df = pd.read_excel(file_path)
    for index, row in df.iterrows():
        df['text'] = df['text'].where((df['link'].str[72:82]) != case[:10], text_ins)
        df.to_excel(file_path, index=False)
    return 0


for filename in os.listdir(directory):
    text = extract_text_from_pdf(f'C:/Users/truhm/Работа/pdf/{filename}')
    print(filename)
    ans = find_keywords(text)
    print(ans[:500])
    insert_text_to_excel(file_p, filename, ans[:500])
