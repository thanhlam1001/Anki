import requests
from bs4 import BeautifulSoup
import json
import urllib.request
import pandas as pd
# Hàm lấy link phát âm từ Oxford Learner's Dictionaries
def get_pronunciation_a_phonetics_links(word):
    base_url = "https://www.oxfordlearnersdictionaries.com/definition/english/"
    url = base_url + word

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
    }

    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"Lỗi kết nối: {e}")
        return None, None

    soup = BeautifulSoup(response.text, 'html.parser')

    # Tìm thẻ chứa phát âm tiếng Anh-Anh (UK)
    pronunciation_uk_div = soup.find('div', class_='sound audio_play_button pron-uk icon-audio')
    if pronunciation_uk_div and 'data-src-mp3' in pronunciation_uk_div.attrs:
        uk_pronunciation = pronunciation_uk_div['data-src-mp3']
    else:
        uk_pronunciation = ""
        print("Can't get the uk_pronunciation!")

    uk_phonetics_span = soup.find('div', class_='phons_br').find('span', class_='phon')
    if uk_phonetics_span:
        uk_phonetics = uk_phonetics_span.text 
    else:
        uk_phonetics = ""
        print("Can't get the uk_phonetics!")

    # Tìm thẻ chứa phát âm tiếng Anh-Mỹ (US)
    pronunciation_us_div = soup.find('div', class_='sound audio_play_button pron-us icon-audio')
    if pronunciation_us_div and 'data-src-mp3' in pronunciation_us_div.attrs:
        us_pronunciation = pronunciation_us_div['data-src-mp3']
    else:
        us_pronunciation = ""
        print("Can't get the us_pronunciation!")
    
    us_phonetics_span = soup.find('div', class_='phons_n_am').find('span', class_='phon')
    if us_phonetics_span:
        us_phonetics = us_phonetics_span.text 
    else:
        us_phonetics = ""
        print("Can't get the us_phonetics!")

    # Trả về cả link phát âm Anh-Anh và Anh-Mỹ
    return [uk_pronunciation, us_pronunciation], [uk_phonetics, us_phonetics]



# Hàm thêm từ vào Anki với link phát âm
def add_note_to_anki(word, example, meaning, uk_pronunciation, us_pronunciation, uk_phonetics, us_phonetics):
    anki_connect_url = 'http://localhost:8765'
    
    note = {
        "action": "addNote",
        "version": 6,
        "params": {
            "note": {
                "deckName": "Default",  # Tên của deck bạn muốn thêm thẻ vào
                "modelName": "Basic",   # Mô hình thẻ có các trường tùy chỉnh
                "fields": {
                    "Word": word.lower(),
                    "Pronunciation UK": uk_pronunciation,
                    "Pronunciation US": us_pronunciation,
                    "UK Phonetics": uk_phonetics,
                    "US Phonetics": us_phonetics,
                    "Example": example,
                    "Back": meaning
                }
            }
        }
    }

    response = requests.post(anki_connect_url, json=note)
    return response.json()

# Hàm main để điều khiển luồng chính của chương trình
def main():


    


    df = pd.read_excel(r'D:\Git_Workspace\Anki_Prj\Anki_Template\Python\vocab.xlsx')

    for index, row in df.iterrows():
        word = row['Word']
        example = row['Example']
        meaning = row['Meaning']
        print(f"********_______{word}_______********")
        pronunciation_lst, phonetics_lst = get_pronunciation_a_phonetics_links(word)

        result = add_note_to_anki(word, example, meaning, pronunciation_lst[0], pronunciation_lst[1], phonetics_lst[0], phonetics_lst[1])

        if result['error'] is None:
            print(f"Adding new word successfully!!")
        else:
            print(result['error'])

if __name__ == "__main__":
    main()
