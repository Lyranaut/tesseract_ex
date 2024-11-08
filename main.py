import pytesseract
from PIL import Image
import sys

# Устанавливаем кодировку вывода на UTF-8
sys.stdout.reconfigure(encoding='utf-8')

# Проверьте версию Tesseract для отладки
print("Версия Tesseract:", pytesseract.get_tesseract_version())

# Укажите путь к tesseract.exe (если требуется)
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

try:
    # Проверьте наличие изображения
    img_path = "text.png"  # укажите путь к изображению
    img = Image.open(img_path)
    print(f"Изображение '{img_path}' успешно загружено.")
    
    # Попробуйте распознать текст
    text = pytesseract.image_to_string(img)
    print("Распознанный текст:\n", text)

except FileNotFoundError:
    print(f"Файл '{img_path}' не найден. Убедитесь, что указали правильный путь к изображению.")
except pytesseract.TesseractNotFoundError:
    print("Tesseract не найден. Проверьте путь к tesseract.exe.")
except Exception as e:
    print("Произошла ошибка:", e)
