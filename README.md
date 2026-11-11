# 🔄 Word Number Converter | مبدل اعداد ورد

A Python GUI application to convert English digits to Persian and vice versa in Microsoft Word (`.docx`) files while preserving the original formatting.

یک ابزار گرافیکی (GUI) کاربردی با پایتون برای تبدیل اعداد انگلیسی به فارسی و بالعکس در فایل‌های مایکروسافت ورد (`.docx`) با حفظ کامل فرمت.

---

## ✨ Features | امکانات
- **Bidirectional Conversion | تبدیل دوطرفه:** Convert English digits to Persian and Persian/Arabic digits to English.
- **Format Preservation | حفظ فرمت:** Fonts, colors, boldness, and other styles remain completely intact.
- **Tables Support | پشتیبانی از جداول:** Converts numbers inside tables, including nested tables.
- **Headers & Footers | هدر و فوتر:** Processes numbers in all parts of the document (paragraphs, tables, headers, footers).
- **Modern GUI | رابط کاربری مدرن:** Built with Tkinter with a clean and user-friendly interface.
- **Drag & Drop | کشیدن و رها کردن:** Drop your file directly into the app (requires `tkinterdnd2`).
- **Safe Output | خروجی امن:** Creates a new file with `_persian` or `_english` suffix without modifying the original file.

---

## 📂 Project Structure | ساختار پروژه
This repository contains two separate GUI applications:

```text
word-number-converter/
│
├── انگلیسی به فارسی/         # English to Persian converter
│   └── word_number_converter_gui.py
│
├── فارسی به انگلیسی/         # Persian to English converter
│   └── word_to_english_converter_gui.py
│
├── .gitignore
└── README.md