import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import asyncio
import threading
import sys
import os

import scrape_fasih


class TextRedirector:
    """Redirect stdout ke widget log."""
    def __init__(self, widget):
        self.widget = widget

    def write(self, text):
        try:
            self.widget.after(0, self._insert, text)
        except Exception:
            pass

    def _insert(self, text):
        self.widget.configure(state="normal")
        self.widget.insert(tk.END, text)
        self.widget.see(tk.END)
        self.widget.configure(state="disabled")

    def flush(self):
        pass


class FasihScraperGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("FASIH Scraper — BPS")
        self.root.geometry("750x620")
        self.root.resizable(True, True)

        self._stop_event = threading.Event()
        self._thread = None
        self._pass_entry = None  # referensi ke widget password

        self._build_ui()

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=14)
        main.pack(fill=tk.BOTH, expand=True)

        # ── Input file ──────────────────────────────────────────────────────
        ttk.Label(main, text="File Petugas:").grid(row=0, column=0, sticky="w", pady=5)
        self.input_var = tk.StringVar()
        ttk.Entry(main, textvariable=self.input_var, width=52).grid(
            row=0, column=1, sticky="ew", padx=6
        )
        ttk.Button(main, text="Browse…", command=self._browse).grid(row=0, column=2, padx=2)

        # ── Username ─────────────────────────────────────────────────────────
        ttk.Label(main, text="Username:").grid(row=1, column=0, sticky="w", pady=5)
        self.username_var = tk.StringVar()
        ttk.Entry(main, textvariable=self.username_var, width=52).grid(
            row=1, column=1, sticky="ew", padx=6
        )

        # ── Password ─────────────────────────────────────────────────────────
        ttk.Label(main, text="Password:").grid(row=2, column=0, sticky="w", pady=5)
        self.password_var = tk.StringVar()
        self._pass_entry = ttk.Entry(
            main, textvariable=self.password_var, show="●", width=40
        )
        self._pass_entry.grid(row=2, column=1, sticky="ew", padx=6)

        self.show_pass_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            main,
            text="Tampilkan",
            variable=self.show_pass_var,
            command=self._toggle_pass,
        ).grid(row=2, column=2, padx=2)

        # ── UPI + UP3 (same row) ─────────────────────────────────────────────
        upi_up3_frame = ttk.Frame(main)
        upi_up3_frame.grid(row=3, column=0, columnspan=3, sticky="ew", padx=6, pady=5)
        ttk.Label(upi_up3_frame, text="UPI:").pack(side=tk.LEFT)
        self.upi_var = tk.StringVar(value="[55]")
        ttk.Entry(upi_up3_frame, textvariable=self.upi_var, width=12).pack(side=tk.LEFT, padx=(4, 16))
        ttk.Label(upi_up3_frame, text="UP3:").pack(side=tk.LEFT)
        self.up3_var = tk.StringVar(value="[55UTR]")
        ttk.Entry(upi_up3_frame, textvariable=self.up3_var, width=12).pack(side=tk.LEFT, padx=(4, 0))

        # ── Sheets URL ───────────────────────────────────────────────────────
        ttk.Label(main, text="Sheets URL:").grid(row=4, column=0, sticky="w", pady=5)
        self.sheets_var = tk.StringVar()
        ttk.Entry(main, textvariable=self.sheets_var, width=52).grid(
            row=4, column=1, sticky="ew", padx=6, columnspan=2
        )

        # ── Tombol Run / Stop ────────────────────────────────────────────────
        btn_frame = ttk.Frame(main)
        btn_frame.grid(row=5, column=0, columnspan=3, pady=10, sticky="w")

        self.run_btn = ttk.Button(btn_frame, text="▶  Run", command=self._run, width=12)
        self.run_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.stop_btn = ttk.Button(
            btn_frame, text="■  Stop", command=self._stop, width=12, state="disabled"
        )
        self.stop_btn.pack(side=tk.LEFT)

        self.status_var = tk.StringVar(value="")
        ttk.Label(btn_frame, textvariable=self.status_var, foreground="#555").pack(
            side=tk.LEFT, padx=14
        )

        # ── Log ───────────────────────────────────────────────────────────────
        ttk.Label(main, text="Log:").grid(row=6, column=0, sticky="w", pady=(8, 2))
        self.log = scrolledtext.ScrolledText(
            main, height=20, state="disabled", font=("Consolas", 9), wrap=tk.WORD
        )
        self.log.grid(row=7, column=0, columnspan=3, sticky="nsew")

        main.columnconfigure(1, weight=1)
        main.rowconfigure(7, weight=1)

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _browse(self):
        path = filedialog.askopenfilename(
            title="Pilih file daftar petugas",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("Semua file", "*.*")],
        )
        if path:
            self.input_var.set(path)

    def _toggle_pass(self):
        self._pass_entry.config(show="" if self.show_pass_var.get() else "●")

    def _log_clear(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", tk.END)
        self.log.configure(state="disabled")

    # ── Run / Stop ────────────────────────────────────────────────────────────

    def _run(self):
        input_file = self.input_var.get().strip()
        username   = self.username_var.get().strip()
        password   = self.password_var.get().strip()
        upi_text   = self.upi_var.get().strip()
        up3_text   = self.up3_var.get().strip()
        sheets_url = self.sheets_var.get().strip()

        if not input_file:
            messagebox.showwarning("Input", "Pilih file daftar petugas terlebih dahulu.")
            return
        if not os.path.exists(input_file):
            messagebox.showerror("File tidak ditemukan", f"{input_file}")
            return
        if not username or not password:
            messagebox.showwarning("Input", "Username dan password harus diisi.")
            return

        self._log_clear()
        sys.stdout = TextRedirector(self.log)

        self._stop_event = threading.Event()
        self.run_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.status_var.set("Berjalan…")

        self._thread = threading.Thread(
            target=self._run_in_thread,
            args=(input_file, username, password, upi_text, up3_text, sheets_url, self._stop_event),
            daemon=True,
        )
        self._thread.start()
        self._poll()

    def _run_in_thread(self, input_file, username, password, upi_text, up3_text, sheets_url, stop_event):
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            loop.run_until_complete(
                scrape_fasih.main_with_stop(
                    stop_event,
                    input_file=input_file,
                    username=username,
                    password=password,
                    sheets_url=sheets_url,
                    upi_text=upi_text,
                    up3_text=up3_text,
                )
            )
        except Exception as exc:
            print(f"\n[ERROR] {exc}")
        finally:
            loop.close()

    def _poll(self):
        if self._thread and self._thread.is_alive():
            self.root.after(500, self._poll)
        else:
            self._on_done()

    def _on_done(self):
        sys.stdout = sys.__stdout__
        self.run_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        stopped = self._stop_event.is_set()
        self.status_var.set("Dihentikan." if stopped else "Selesai ✓")

    def _stop(self):
        self._stop_event.set()
        self.stop_btn.configure(state="disabled")
        self.status_var.set("Menghentikan…")

    def _on_close(self):
        self._stop_event.set()  # signal any running scrape to stop
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    FasihScraperGUI(root)
    root.mainloop()
