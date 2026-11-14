# 🔄 Word Number Converter | مبدل اعداد ورد

**Version 1.1.0**

A Python GUI application to convert English digits to Persian and vice
versa in Microsoft Word (`.docx`) files while preserving the original
formatting.

یک ابزار گرافیکی (GUI) با پایتون برای تبدیل اعداد انگلیسی به فارسی و
بالعکس در فایل‌های مایکروسافت ورد (`.docx`) با حفظ کامل فرمت.

---

## ✨ Features | امکانات

- **Bidirectional Conversion | تبدیل دوطرفه:** One application with a
  direction selector — English digits to Persian, and Persian/Arabic
  digits to English.
- **Format Preservation | حفظ فرمت:** Only the text of each run is
  changed, so fonts, colors, boldness and other styles remain intact.
- **Tables Support | پشتیبانی از جداول:** Converts numbers inside
  tables, including nested tables.
- **Headers & Footers | هدر و فوتر:** Default, first-page and
  even-page headers and footers are processed.
- **Text Boxes | جعبه‌های متن:** Numbers inside text boxes are
  converted too, as are footnotes, endnotes and comments when the
  document has them.
- **Drag & Drop | کشیدن و رها کردن:** Drop your file directly into
  the app (requires `tkinterdnd2`; the app works without it too).
- **Safe Output | خروجی امن:** By default creates a new file with a
  `_persian` or `_english` suffix without modifying the original file.
  Overwriting the original is possible but always asks for
  confirmation first.

### What is *not* converted | چه چیزهایی تبدیل نمی‌شوند

- Field results that Word generates itself, such as automatic page
  numbers and tables of contents.
- Digits inside embedded objects or images.

---

## 📂 Project Structure | ساختار پروژه

```
word-number-converter/
│
├── converter.py                  # Shared conversion core (both directions, no GUI)
├── word_number_converter_gui.py  # The single GUI application
├── requirements.txt
├── tests/
│   └── test_converter.py
├── .gitignore
├── LICENSE
└── README.md
```

---

## 🛠 Requirements | پیش‌نیازها

- Python 3.9 or higher
- Packages listed in `requirements.txt`

نصب پیش‌نیازها:

```
pip install -r requirements.txt
```

---

## 🚀 How to Run | نحوه اجرا

```
python word_number_converter_gui.py
```

Then choose the conversion direction inside the app
(English → Persian or Persian → English), pick your `.docx` file and
press «شروع تبدیل».

برای اجرا کافی است دستور بالا را بزنید و جهت تبدیل را داخل خود
برنامه انتخاب کنید.

---

## 🧪 Tests | تست‌ها

```
python -m unittest discover -s tests -v
```

---

## 📦 Build EXE | ساخت فایل اجرایی (EXE)

To create a standalone .exe file using PyInstaller:

```
python -m PyInstaller --onefile --windowed --name "WordNumberConverter" --hidden-import tkinterdnd2 --collect-all tkinterdnd2 word_number_converter_gui.py
```

---

## 👨‍💻 Developer | توسعه‌دهنده

- GitHub: @WAI490

---

⭐ If you find this project useful, please give it a star! | اگر این پروژه برایتان مفید بود، به آن ستاره بدهید! ⭐
