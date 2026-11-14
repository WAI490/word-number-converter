"""Word Number Converter — GUI.

One application for both directions:
  * English digits → Persian digits
  * Persian/Arabic digits → English digits

All conversion logic lives in converter.py; this file is only the
Tkinter interface.
"""

import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from converter import output_path_for, process_document


class ConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("مبدل اعداد ورد | Word Number Converter")
        self.root.geometry("560x480")
        self.root.resizable(False, False)
        self.root.configure(bg="#f5f5f5")

        self.input_path = None
        self.direction = tk.StringVar(value="en2fa")
        self.use_overwrite = tk.BooleanVar(value=False)
        self.dnd_available = False

        try:
            from tkinterdnd2 import DND_FILES

            if hasattr(self.root, "drop_target_register"):
                self.root.drop_target_register(DND_FILES)
                self.root.dnd_bind("<<Drop>>", self._on_drop)
                self.dnd_available = True
        except Exception:
            pass

        self._build_ui()

    def _build_ui(self):
        title_frame = tk.Frame(self.root, bg="#2c3e50", height=70)
        title_frame.pack(fill="x")
        title_frame.pack_propagate(False)

        tk.Label(
            title_frame,
            text="مبدل اعداد در فایل ورد",
            bg="#2c3e50", fg="white",
            font=("Tahoma", 14, "bold"),
        ).pack(pady=18)

        main = tk.Frame(self.root, bg="#f5f5f5", padx=20, pady=15)
        main.pack(fill="both", expand=True)

        file_frame = tk.LabelFrame(
            main, text=" فایل Word ",
            bg="#f5f5f5", font=("Tahoma", 10, "bold"),
            padx=10, pady=10,
        )
        file_frame.pack(fill="x")

        self.file_label = tk.Label(
            file_frame,
            text="هیچ فایلی انتخاب نشده",
            bg="#ffffff", fg="#666",
            anchor="w", relief="sunken", bd=1,
            padx=8, pady=6, font=("Tahoma", 9),
            wraplength=400, justify="left",
        )
        self.file_label.pack(fill="x", pady=(0, 8))

        btn_row = tk.Frame(file_frame, bg="#f5f5f5")
        btn_row.pack(fill="x")

        tk.Button(
            btn_row, text="انتخاب فایل",
            command=self.select_file,
            bg="#3498db", fg="white",
            font=("Tahoma", 10, "bold"),
            relief="flat", padx=15, pady=6, cursor="hand2",
        ).pack(side="left")

        tk.Button(
            btn_row, text="پاک کردن",
            command=self.clear_file,
            bg="#95a5a6", fg="white",
            font=("Tahoma", 10),
            relief="flat", padx=10, pady=6, cursor="hand2",
        ).pack(side="left", padx=5)

        if self.dnd_available:
            tk.Label(
                btn_row, text="(یا فایل را بکشید و رها کنید)",
                bg="#f5f5f5", fg="#888", font=("Tahoma", 8),
            ).pack(side="right")

        dir_frame = tk.LabelFrame(
            main, text=" جهت تبدیل ",
            bg="#f5f5f5", font=("Tahoma", 10, "bold"),
            padx=10, pady=10,
        )
        dir_frame.pack(fill="x", pady=(12, 0))

        tk.Radiobutton(
            dir_frame, text="(123 → ۱۲۳)",
            variable=self.direction, value="en2fa",
            bg="#f5f5f5", font=("Tahoma", 9), anchor="w",
        ).pack(fill="x")

        tk.Radiobutton(
            dir_frame, text="(١٢٣ → 123)",
            variable=self.direction, value="fa2en",
            bg="#f5f5f5", font=("Tahoma", 9), anchor="w",
        ).pack(fill="x")

        options_frame = tk.LabelFrame(
            main, text=" تنظیمات خروجی ",
            bg="#f5f5f5", font=("Tahoma", 10, "bold"),
            padx=10, pady=10,
        )
        options_frame.pack(fill="x", pady=12)

        tk.Radiobutton(
            options_frame, text="ساخت فایل جدید (نام اصلی + پسوند جهت تبدیل)",
            variable=self.use_overwrite, value=False,
            bg="#f5f5f5", font=("Tahoma", 9), anchor="w",
        ).pack(fill="x")

        tk.Radiobutton(
            options_frame, text="بازنویسی روی فایل اصلی (پشتیبان بگیرید!)",
            variable=self.use_overwrite, value=True,
            bg="#f5f5f5", fg="#c0392b", font=("Tahoma", 9), anchor="w",
        ).pack(fill="x")

        self.convert_btn = tk.Button(
            main, text="شروع تبدیل",
            command=self.start_conversion,
            bg="#27ae60", fg="white",
            font=("Tahoma", 12, "bold"),
            relief="flat", pady=10, cursor="hand2",
            state="disabled",
        )
        self.convert_btn.pack(fill="x", pady=(0, 10))

        self.status_label = tk.Label(
            main, text="آماده",
            bg="#f5f5f5", fg="#555",
            font=("Tahoma", 9), anchor="center",
        )
        self.status_label.pack(fill="x")

        self.progress = ttk.Progressbar(main, mode="indeterminate")
        self.progress.pack(fill="x", pady=(5, 0))

    def select_file(self):
        path = filedialog.askopenfilename(
            title="انتخاب فایل Word",
            filetypes=[("Word Document", "*.docx")],
        )
        if path:
            self._set_file(path)

    def _on_drop(self, event):
        path = event.data.strip()
        if path.startswith("{") and path.endswith("}"):
            path = path[1:-1]
        self._set_file(path)

    def _set_file(self, path):
        if not path.lower().endswith(".docx"):
            messagebox.showwarning("فایل نامعتبر", "فقط فایل‌های .docx پشتیبانی می‌شوند.")
            return
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
                "فایل اصلی بازنویسی خواهد شد و امکان بازگشت نیست.\nآیا مطمئن هستید؟",
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
            direction = self.direction.get()

            if self.use_overwrite.get():
                output_path = input_path
            else:
                output_path = output_path_for(input_path, direction)

            process_document(input_path, output_path, direction)

            self.root.after(0, self._on_success, output_path)

        except Exception as e:
            self.root.after(0, self._on_error, str(e))

    def _on_success(self, output_path):
        self.progress.stop()
        self.convert_btn.config(state="normal")
        self.status_label.config(
            text=f"تبدیل موفق! → {os.path.basename(output_path)}", fg="#27ae60"
        )
        messagebox.showinfo(
            "موفق",
            f"تبدیل با موفقیت انجام شد.\n\nفایل خروجی:\n{output_path}",
        )

    def _on_error(self, err):
        self.progress.stop()
        self.convert_btn.config(state="normal")
        self.status_label.config(text="خطا در تبدیل", fg="#c0392b")
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
