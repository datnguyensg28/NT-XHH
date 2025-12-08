# modules/utils.py
from num2words import num2words

def to_vietnamese_words(number):
    try:
        return num2words(number, lang="vi").capitalize()
    except Exception:
        return ""
