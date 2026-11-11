import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from docx import Document

# نگاشت اعداد فارسی و عربی به انگلیسی
ENGLISH_DIGITS = {
    '۰': '0', '۱': '1', '۲': '2', '۳': '3', '۴': '4',
    '۵': '5', '۶': '6', '۷': '7', '۸': '8', '۹': '9',
    # اضافه کردن اعداد عربی برای اطمینان بیشتر
    '٠': '0', '١': '1', '٢': '2', '٣': '3', '٤': '4',
    '٥': '5', '٦': '6', '٧': '7', '٨': '8', '٩': '9'
}

def convert_to_english(text):
    """تبدیل اعداد فارسی/عربی به انگلیسی در یک رشته"""
    return ''.join(ENGLISH_DIGITS.get(ch, ch) for ch in text)

def process_paragraph(paragraph):
    """پردازش پاراگراف با حفظ فرمت اولیه"""
    for run in paragraph.runs:
        if any(c in ENGLISH_DIGITS for c in run.text):
            run.text = convert_to_english(run.text)

def process_table(table):
    """پردازش جداول (شامل جداول تودرتو)"""
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                process_paragraph(paragraph)
            for nested_table in cell.tables:
                process_table(nested_table)

def process_document(input_path, output_path):
    doc = Document(input_path)

    # پردازش پاراگراف‌های اصلی
    for paragraph in doc.paragraphs:
        process_paragraph(paragraph)

    # پردازش جداول
    for table in doc.tables:
        process_table(table)

    # پردازش هدر و فوتر
    for section in doc.sections:
        for header in [section.header, section.first_page_header, section.even_page_header]:
            if header:
                for p in header.paragraphs:
                    process_paragraph(p)
                for t in header.tables:
                    process_table(t)

        for footer in [section.footer, section.first_page_footer, section.even_page_footer]:
            if footer:
                for p in footer.paragraphs:
                    process_paragraph(p)
                for t in footer.tables:
                    process_table(t)

    doc.save(output_path)


class ConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("تبدیل اعداد فارسی به انگلیسی در Word")
        self.root.geometry("560x420")
        self.root.resizable(False, False)
        self.root.configure(bg="#f5f5f5")

        self.input_path = None
        self.use_overwrite = tk.BooleanVar(value=False)
        self.dnd_available = False
        
        # بررسی و فعال‌سازی Drag & Drop
        try:
            from tkinterdnd2 import DND_FILES
            if hasattr(self.root, 'drop_target_register'):
                self.root.drop_target_register(DND_FILES)
                self.root.dnd_bind('<<Drop>>', self._on_drop)
                self.dnd_available = True
        except Exception:
            self.dnd_available = False

        self._build_ui()

    def _build_ui(self):
        # ---------- عنوان ----------
        title_frame = tk.Frame(self.root, bg="#2c3e50", height=70)
        title_frame.pack(fill="x")
        title_frame.pack_propagate(False)

        tk.Label(
            title_frame,
            text="🔄 تبدیل اعداد فارسی به انگلیسی",
            bg="#2c3e50", fg="white",
            font=("Tahoma", 14, "bold")
        ).pack(pady=18)

        # ---------- بدنه اصلی ----------
        main = tk.Frame(self.root, bg="#f5f5f5", padx=20, pady=15)
        main.pack(fill="both", expand=True)

        # انتخاب فایل
        file_frame = tk.LabelFrame(
            main, text=" فایل Word ",
            bg="#f5f5f5", font=("Tahoma", 10, "bold"),
            padx=10, pady=10
        )
        file_frame.pack(fill="x")

        self.file_label = tk.Label(
            file_frame,
            text="هیچ فایلی انتخاب نشده",
            bg="#ffffff", fg="#666",
            anchor="w", relief="sunken", bd=1,
            padx=8, pady=6, font=("Tahoma", 9),
            wraplength=400, justify="left"
        )
        self.file_label.pack(fill="x", pady=(0, 8))

        btn_row = tk.Frame(file_frame, bg="#f5f5f5")
        btn_row.pack(fill="x")

        tk.Button(
            btn_row, text="📂 انتخاب فایل",
            command=self.select_file,
            bg="#3498db", fg="white",
            font=("Tahoma", 10, "bold"),
            relief="flat", padx=15, pady=6, cursor="hand2"
        ).pack(side="left")

        tk.Button(
            btn_row, text="✖ پاک کردن",
            command=self.clear_file,
            bg="#95a5a6", fg="white",
            font=("Tahoma", 10),
            relief="flat", padx=10, pady=6, cursor="hand2"
        ).pack(side="left", padx=5)

        if self.dnd_available:
            tk.Label(
                btn_row, text="(یا فایل را بکشید و رها کنید)",
                bg="#f5f5f5", fg="#888", font=("Tahoma", 8)
            ).pack(side="right")

        # گزینه‌های خروجی
        options_frame = tk.LabelFrame(
            main, text=" تنظیمات خروجی ",
            bg="#f5f5f5", font=("Tahoma", 10, "bold"),
            padx=10, pady=10
        )
        options_frame.pack(fill="x", pady=12)

        tk.Radiobutton(
            options_frame, text="ساخت فایل جدید (نام اصلی + _english)",
            variable=self.use_overwrite, value=False,
            bg="#f5f5f5", font=("Tahoma", 9), anchor="w"
        ).pack(fill="x")

        tk.Radiobutton(
            options_frame, text="⚠ بازنویسی روی فایل اصلی (پشتیبان بگیرید!)",
            variable=self.use_overwrite, value=True,
            bg="#f5f5f5", fg="#c0392b", font=("Tahoma", 9), anchor="w"
        ).pack(fill="x")

        # دکمه تبدیل
        self.convert_btn = tk.Button(
            main, text="🚀 شروع تبدیل",
            command=self.start_conversion,
            bg="#27ae60", fg="white",
            font=("Tahoma", 12, "bold"),
            relief="flat", pady=10, cursor="hand2",
            state="disabled"
        )
        self.convert_btn.pack(fill="x", pady=(0, 10))

        # وضعیت
        self.status_label = tk.Label(
            main, text="آماده",
            bg="#f5f5f5", fg="#555",
            font=("Tahoma", 9), anchor="center"
        )
        self.status_label.pack(fill="x")

        self.progress = ttk.Progressbar(main, mode="indeterminate")
        self.progress.pack(fill="x", pady=(5, 0))

    # ---------- رویدادها ----------
    def select_file(self):
        path = filedialog.askopenfilename(
            title="انتخاب فایل Word",
            filetypes=[("Word Document", "*.docx"), ("همه فایل‌ها", "*.*")]
        )
        if path:
            self._set_file(path)

    def _on_drop(self, event):
        path = event.data.strip()
        if path.startswith("{") and path.endswith("}"):
            path = path[1:-1]
        if path.lower().endswith(".docx"):
            self._set_file(path)
        else:
            messagebox.showwarning("فایل نامعتبر", "فقط فایل‌های .docx پشتیبانی می‌شوند.")

    def _set_file(self, path):
        self.input_path = path
        self.file_label.config(text=path, fg="#2c3e50")
        self.convert_btn.config(state="normal")
        self.status_label.config(text="فایل انتخاب شد. آماده تبدیل.", fg="#555")

    def clear_file(self):
        self.input_path = None
        self.file_label.config(text="هیچ فایلی انتخاب نشده", fg="#666")
        self.convert_btn.config(state="disabled")
        self.status_label.config(text="آماده", fg="#555")

    def start_conversion(self):
        if not self.input_path:
            return

        if not os.path.exists(self.input_path):
            messagebox.showerror("خطا", "فایل انتخاب‌شده پیدا نشد.")
            return

        if self.use_overwrite.get():
            confirm = messagebox.askyesno(
                "تایید بازنویسی",
                "فایل اصلی بازنویسی خواهد شد و امکان بازگشت نیست.\nآیا مطمئن هستید؟"
            )
            if not confirm:
                return

        self.convert_btn.config(state="disabled")
        self.status_label.config(text="در حال تبدیل...", fg="#2980b9")
        self.progress.start(10)

        t = threading.Thread(target=self._run_conversion, daemon=True)
        t.start()

    def _run_conversion(self):
        try:
            input_path = self.input_path

            if self.use_overwrite.get():
                output_path = input_path
            else:
                base, ext = os.path.splitext(input_path)
                output_path = f"{base}_english{ext}"

            process_document(input_path, output_path)

            self.root.after(0, self._on_success, output_path)

        except Exception as e:
            self.root.after(0, self._on_error, str(e))

    def _on_success(self, output_path):
        self.progress.stop()
        self.convert_btn.config(state="normal")
        self.status_label.config(text=f"✅ تبدیل موفق! → {os.path.basename(output_path)}", fg="#27ae60")
        messagebox.showinfo(
            "موفق",
            f"تبدیل با موفقیت انجام شد.\n\nفایل خروجی:\n{output_path}"
        )

    def _on_error(self, err):
        self.progress.stop()
        self.convert_btn.config(state="normal")
        self.status_label.config(text="❌ خطا در تبدیل", fg="#c0392b")
        messagebox.showerror("خطا", f"عملیات ناموفق بود:\n\n{err}")


def main():
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except ImportError:
        root = tk.Tk()

    ConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()