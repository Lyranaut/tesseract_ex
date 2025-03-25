import pytesseract
from PIL import Image
from docx import Document
import os
import sys
import tkinter as tk
from tkinter import filedialog

# Устанавливаем кодировку UTF-8
sys.stdout.reconfigure(encoding='utf-8')

# Указываем путь к tesseract.exe (если требуется)
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Определяем путь к рабочему столу
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
word_filename = os.path.join(desktop_path, "recognized_text.docx")

# Открываем диалог выбора файла
root = tk.Tk()
root.withdraw()  # Скрыть главное окно
file_path = filedialog.askopenfilename(title="Выберите изображение", filetypes=[("Изображения", "*.png;*.jpg;*.jpeg;*.bmp;*.tiff")])

if not file_path:
    print("Файл не выбран. Программа завершена.")
    sys.exit()

try:
    # Загружаем изображение
    img = Image.open(file_path)
    print(f"Изображение '{file_path}' успешно загружено.")

    # Распознаем текст
    text = pytesseract.image_to_string(img)
    print("Распознанный текст:\n", text)

    # Создаем Word-файл и записываем в него текст
    doc = Document()
    doc.add_paragraph(text)
    doc.save(word_filename)

    print(f"Файл '{word_filename}' успешно создан на рабочем столе.")

except pytesseract.TesseractNotFoundError:
    print("Tesseract не найден. Проверьте путь к tesseract.exe.")
except Exception as e:
    print("Произошла ошибка:", e)