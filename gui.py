# -*- coding: utf-8 -*-
"""
واجهة رسومية بسيطة لأداة AliExpress Product Scraper
====================================================
شغّلها بنقرة مزدوجة على start.bat أو بالأمر: python gui.py

لا تحتاج لكتابة أي أوامر: الصق الرابط، اكتب العدد، واضغط "ابدأ الاستخراج".
"""

import queue
import sys
import threading
import tkinter as tk
from tkinter import messagebox, scrolledtext

import scraper
from playwright.sync_api import sync_playwright


class QueueWriter:
    """يلتقط مخرجات print() من scraper.py ويعرضها في نافذة السجل."""

    def __init__(self, q: queue.Queue):
        self.q = q

    def write(self, text):
        if text.strip():
            self.q.put(text.strip())

    def flush(self):
        pass


class ScraperApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.q = queue.Queue()
        self.running = False

        root.title("AliExpress Product Scraper 🛒")
        root.geometry("680x520")
        root.minsize(560, 420)

        pad = {"padx": 12, "pady": 6}

        # --- خانة الرابط ---
        tk.Label(root, text="رابط الصفحة (الصق هنا):", anchor="e").pack(fill="x", **pad)
        self.url_var = tk.StringVar()
        url_entry = tk.Entry(root, textvariable=self.url_var, justify="left")
        url_entry.pack(fill="x", **pad)
        url_entry.focus()

        # --- خانة عدد المنتجات ---
        row = tk.Frame(root)
        row.pack(fill="x", **pad)
        tk.Label(row, text="عدد المنتجات:").pack(side="right")
        self.count_var = tk.StringVar(value="50")
        tk.Entry(row, textvariable=self.count_var, width=8, justify="center").pack(side="right", padx=8)

        # --- خيار إخفاء المتصفح ---
        self.headless_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            row,
            text="إخفاء المتصفح أثناء العمل (قد يزيد احتمال ظهور CAPTCHA)",
            variable=self.headless_var,
        ).pack(side="left")

        # --- زر البدء ---
        self.start_btn = tk.Button(
            root, text="🚀 ابدأ الاستخراج", font=("Segoe UI", 12, "bold"),
            bg="#e62e04", fg="white", command=self.start,
        )
        self.start_btn.pack(fill="x", padx=12, pady=10)

        # --- نافذة السجل ---
        tk.Label(root, text="سجل العملية:", anchor="e").pack(fill="x", padx=12)
        self.log = scrolledtext.ScrolledText(root, state="disabled", height=14)
        self.log.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        self.poll_queue()

    # ------------------------------------------------------------------
    def log_line(self, text: str):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def poll_queue(self):
        """نقل الرسائل من خيط العمل إلى الواجهة (يجب أن يتم في الخيط الرئيسي)."""
        try:
            while True:
                msg = self.q.get_nowait()
                if msg == "__DONE__":
                    self.running = False
                    self.start_btn.configure(state="normal", text="🚀 ابدأ الاستخراج")
                else:
                    self.log_line(msg)
        except queue.Empty:
            pass
        self.root.after(200, self.poll_queue)

    # ------------------------------------------------------------------
    def start(self):
        if self.running:
            return
        url = self.url_var.get().strip()
        if not url.startswith("http"):
            messagebox.showwarning("تنبيه", "الصق رابطاً صحيحاً يبدأ بـ http أو https")
            return
        try:
            count = int(self.count_var.get().strip())
            if count <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("تنبيه", "أدخل عدد منتجات صحيحاً (رقم أكبر من صفر)")
            return

        self.running = True
        self.start_btn.configure(state="disabled", text="⏳ جاري الاستخراج ...")
        self.log_line(f"▶️ البدء: {count} منتجاً من:\n{url}")

        threading.Thread(
            target=self.worker, args=(url, count, self.headless_var.get()), daemon=True
        ).start()

    def worker(self, url: str, count: int, headless: bool):
        """تنفيذ الاستخراج في خيط منفصل حتى لا تتجمد الواجهة."""
        old_stdout = sys.stdout
        sys.stdout = QueueWriter(self.q)
        try:
            scraper.HEADLESS = headless
            with sync_playwright() as p:
                browser, page = scraper.open_page(p, url)
                try:
                    products = scraper.collect_products(page, count)
                finally:
                    browser.close()

            if products:
                scraper.save_csv(products)
                scraper.save_xlsx(products)
                scraper.save_json(products)
                print(f"🎉 انتهى! تم استخراج {len(products)} منتجاً.")
                print("📂 الملفات محفوظة بجانب البرنامج: products.csv / products.xlsx / products.json")
            else:
                print("❌ لم يتم العثور على منتجات. تأكد من الرابط أو ألغِ خيار إخفاء المتصفح.")
        except Exception as exc:
            print(f"❌ خطأ: {exc}")
        finally:
            sys.stdout = old_stdout
            self.q.put("__DONE__")


def main():
    root = tk.Tk()
    ScraperApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
