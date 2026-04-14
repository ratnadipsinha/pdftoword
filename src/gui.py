"""
PDF2Word Pro — Desktop GUI (Tkinter)
Features: drag-and-drop, batch folder, OCR language selector, progress bar.
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path


# Try to enable drag-and-drop (optional dependency)
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_SUPPORTED = True
except ImportError:
    DND_SUPPORTED = False

from .converter import PDF2WordConverter


# -----------------------------------------------------------------------
# Color palette
# -----------------------------------------------------------------------
BG        = "#1e1e2e"
SURFACE   = "#2a2a3e"
ACCENT    = "#7c6af7"
ACCENT_LT = "#a89cf7"
TEXT      = "#cdd6f4"
TEXT_DIM  = "#6c7086"
SUCCESS   = "#a6e3a1"
ERROR     = "#f38ba8"
BORDER    = "#45475a"


class PDF2WordApp:
    def __init__(self):
        # Use DnD-capable root if available
        if DND_SUPPORTED:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()

        self.root.title("PDF2Word Pro")
        self.root.geometry("700x580")
        self.root.resizable(True, True)
        self.root.configure(bg=BG)
        self.root.minsize(600, 500)

        self._set_icon()
        self._build_ui()
        self._files: list[str] = []
        self._running = False

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self):
        """Assemble all UI sections."""
        self._build_header()
        self._build_drop_zone()
        self._build_file_list()
        self._build_options()
        self._build_output_row()
        self._build_action_row()
        self._build_progress()
        self._build_log()

    def _build_header(self):
        frame = tk.Frame(self.root, bg=BG)
        frame.pack(fill="x", padx=20, pady=(20, 0))

        tk.Label(
            frame, text="PDF2Word Pro",
            font=("Segoe UI", 22, "bold"),
            fg=ACCENT_LT, bg=BG
        ).pack(side="left")

        tk.Label(
            frame, text="  — Preserve formatting, convert to .docx",
            font=("Segoe UI", 10),
            fg=TEXT_DIM, bg=BG
        ).pack(side="left", pady=6)

    def _build_drop_zone(self):
        outer = tk.Frame(self.root, bg=BG)
        outer.pack(fill="x", padx=20, pady=12)

        self.drop_frame = tk.Frame(
            outer, bg=SURFACE,
            highlightbackground=BORDER,
            highlightthickness=1,
            cursor="hand2",
        )
        self.drop_frame.pack(fill="x", ipady=18)

        self.drop_label = tk.Label(
            self.drop_frame,
            text="Drag & Drop PDF files here\nor click to browse",
            font=("Segoe UI", 11),
            fg=TEXT_DIM, bg=SURFACE,
            justify="center",
        )
        self.drop_label.pack(pady=10)

        # Click to browse
        self.drop_frame.bind("<Button-1>", lambda e: self._browse_files())
        self.drop_label.bind("<Button-1>", lambda e: self._browse_files())

        # Drag-and-drop
        if DND_SUPPORTED:
            self.drop_frame.drop_target_register(DND_FILES)
            self.drop_frame.dnd_bind("<<Drop>>", self._on_drop)

    def _build_file_list(self):
        frame = tk.Frame(self.root, bg=BG)
        frame.pack(fill="both", expand=False, padx=20)

        header = tk.Frame(frame, bg=BG)
        header.pack(fill="x")
        tk.Label(header, text="Files to convert:", font=("Segoe UI", 9),
                 fg=TEXT_DIM, bg=BG).pack(side="left")
        tk.Button(
            header, text="Clear all", font=("Segoe UI", 9),
            fg=ACCENT_LT, bg=BG, bd=0, cursor="hand2",
            command=self._clear_files, activebackground=BG, activeforeground=ACCENT
        ).pack(side="right")

        list_frame = tk.Frame(frame, bg=SURFACE, highlightbackground=BORDER,
                              highlightthickness=1)
        list_frame.pack(fill="both", expand=False)

        self.file_listbox = tk.Listbox(
            list_frame, bg=SURFACE, fg=TEXT, selectbackground=ACCENT,
            font=("Segoe UI", 9), bd=0, height=5,
            selectforeground=BG, highlightthickness=0,
        )
        scrollbar = ttk.Scrollbar(list_frame, command=self.file_listbox.yview)
        self.file_listbox.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.file_listbox.pack(fill="both", expand=True, padx=4, pady=4)

    def _build_options(self):
        frame = tk.Frame(self.root, bg=BG)
        frame.pack(fill="x", padx=20, pady=(10, 0))

        tk.Label(frame, text="OCR Language:", font=("Segoe UI", 9),
                 fg=TEXT_DIM, bg=BG).pack(side="left")

        self.ocr_lang_var = tk.StringVar(value="eng")
        ocr_cb = ttk.Combobox(
            frame, textvariable=self.ocr_lang_var, width=10,
            values=["eng", "fra", "deu", "spa", "ita", "por", "hin", "chi_sim"],
            state="readonly",
        )
        ocr_cb.pack(side="left", padx=8)

        self.batch_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            frame, text="Batch folder mode", variable=self.batch_var,
            font=("Segoe UI", 9), fg=TEXT, bg=BG,
            selectcolor=SURFACE, activebackground=BG, activeforeground=TEXT,
            command=self._toggle_batch,
        ).pack(side="left", padx=16)

    def _build_output_row(self):
        frame = tk.Frame(self.root, bg=BG)
        frame.pack(fill="x", padx=20, pady=6)

        tk.Label(frame, text="Output folder:", font=("Segoe UI", 9),
                 fg=TEXT_DIM, bg=BG).pack(side="left")

        self.output_var = tk.StringVar(value=str(Path.home() / "Documents"))
        tk.Entry(
            frame, textvariable=self.output_var,
            font=("Segoe UI", 9), bg=SURFACE, fg=TEXT,
            insertbackground=TEXT, relief="flat", bd=4,
        ).pack(side="left", fill="x", expand=True, padx=8)

        tk.Button(
            frame, text="Browse", font=("Segoe UI", 9),
            fg=TEXT, bg=BORDER, bd=0, padx=8, pady=3,
            cursor="hand2", command=self._browse_output,
            activebackground=ACCENT, activeforeground=BG,
        ).pack(side="left")

    def _build_action_row(self):
        frame = tk.Frame(self.root, bg=BG)
        frame.pack(fill="x", padx=20, pady=10)

        self.convert_btn = tk.Button(
            frame, text="Convert to Word",
            font=("Segoe UI", 11, "bold"),
            fg=BG, bg=ACCENT, bd=0, padx=20, pady=8,
            cursor="hand2", command=self._start_conversion,
            activebackground=ACCENT_LT, activeforeground=BG,
        )
        self.convert_btn.pack(side="left")

        self.open_output_btn = tk.Button(
            frame, text="Open Output Folder",
            font=("Segoe UI", 10),
            fg=TEXT, bg=SURFACE, bd=0, padx=14, pady=8,
            cursor="hand2", command=self._open_output_folder,
            activebackground=BORDER, activeforeground=TEXT,
        )
        self.open_output_btn.pack(side="left", padx=10)

    def _build_progress(self):
        frame = tk.Frame(self.root, bg=BG)
        frame.pack(fill="x", padx=20)

        self.progress_var = tk.DoubleVar(value=0)
        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "Custom.Horizontal.TProgressbar",
            troughcolor=SURFACE, background=ACCENT, thickness=8,
        )
        self.progress_bar = ttk.Progressbar(
            frame, variable=self.progress_var,
            style="Custom.Horizontal.TProgressbar",
            maximum=100,
        )
        self.progress_bar.pack(fill="x", pady=4)

        self.status_label = tk.Label(
            frame, text="Ready", font=("Segoe UI", 9),
            fg=TEXT_DIM, bg=BG, anchor="w",
        )
        self.status_label.pack(fill="x")

    def _build_log(self):
        frame = tk.Frame(self.root, bg=BG)
        frame.pack(fill="both", expand=True, padx=20, pady=(6, 16))

        self.log_text = tk.Text(
            frame, bg=SURFACE, fg=TEXT, font=("Consolas", 9),
            bd=0, relief="flat", state="disabled", height=6,
            highlightbackground=BORDER, highlightthickness=1,
        )
        scrollbar = ttk.Scrollbar(frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True)

        # Color tags
        self.log_text.tag_config("ok", foreground=SUCCESS)
        self.log_text.tag_config("err", foreground=ERROR)
        self.log_text.tag_config("info", foreground=ACCENT_LT)

    # ------------------------------------------------------------------
    # File management
    # ------------------------------------------------------------------

    def _browse_files(self):
        files = filedialog.askopenfilenames(
            title="Select PDF files",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
        )
        for f in files:
            self._add_file(f)

    def _browse_output(self):
        folder = filedialog.askdirectory(title="Select output folder")
        if folder:
            self.output_var.set(folder)

    def _on_drop(self, event):
        """Handle drag-and-drop file drop."""
        raw = event.data
        # tkinterdnd2 wraps paths with spaces in braces
        files = self.root.tk.splitlist(raw)
        for f in files:
            if f.lower().endswith(".pdf"):
                self._add_file(f)
            elif os.path.isdir(f):
                for pdf in Path(f).rglob("*.pdf"):
                    self._add_file(str(pdf))

    def _add_file(self, path: str):
        if path not in self._files:
            self._files.append(path)
            self.file_listbox.insert("end", Path(path).name)
        self.drop_label.config(text=f"{len(self._files)} file(s) queued")

    def _clear_files(self):
        self._files.clear()
        self.file_listbox.delete(0, "end")
        self.drop_label.config(text="Drag & Drop PDF files here\nor click to browse")
        self.progress_var.set(0)
        self.status_label.config(text="Ready", fg=TEXT_DIM)

    def _toggle_batch(self):
        if self.batch_var.get():
            folder = filedialog.askdirectory(title="Select folder with PDFs")
            if folder:
                self._clear_files()
                for pdf in Path(folder).glob("*.pdf"):
                    self._add_file(str(pdf))
        else:
            self._clear_files()

    # ------------------------------------------------------------------
    # Conversion
    # ------------------------------------------------------------------

    def _start_conversion(self):
        if self._running:
            return
        if not self._files:
            messagebox.showwarning("No files", "Please add PDF files first.")
            return

        output_dir = self.output_var.get().strip()
        if not output_dir:
            messagebox.showwarning("No output", "Please select an output folder.")
            return

        self._running = True
        self.convert_btn.config(state="disabled", text="Converting...")
        self.progress_var.set(0)
        self._log("Starting conversion...\n", "info")

        thread = threading.Thread(
            target=self._run_conversion,
            args=(list(self._files), output_dir, self.ocr_lang_var.get()),
            daemon=True,
        )
        thread.start()

    def _run_conversion(self, files: list, output_dir: str, ocr_lang: str):
        total = len(files)
        ok_count = 0

        def progress_cb(current, tot, filename):
            pct = (current / tot) * 100
            self.root.after(0, lambda: self.progress_var.set(pct))
            self.root.after(0, lambda: self.status_label.config(
                text=f"Converting {current}/{tot}: {filename}", fg=TEXT_DIM
            ))

        converter = PDF2WordConverter(ocr_lang=ocr_lang, progress_cb=progress_cb)

        for idx, pdf_path in enumerate(files, start=1):
            out_path = str(Path(output_dir) / Path(pdf_path).with_suffix(".docx").name)
            progress_cb(idx, total, Path(pdf_path).name)

            result = converter.convert(pdf_path, out_path)

            if result.success:
                ok_count += 1
                msg = f"[OK] {Path(pdf_path).name}  ({result.page_count}p, {result.duration:.1f}s)\n"
                self.root.after(0, lambda m=msg: self._log(m, "ok"))
            else:
                msg = f"[FAIL] {Path(pdf_path).name}  — {result.error}\n"
                self.root.after(0, lambda m=msg: self._log(m, "err"))

        # Done
        def finish():
            self.progress_var.set(100)
            self.status_label.config(
                text=f"Done — {ok_count}/{total} converted successfully",
                fg=SUCCESS if ok_count == total else ERROR,
            )
            self.convert_btn.config(state="normal", text="Convert to Word")
            self._running = False
            self._log(f"\nFinished: {ok_count}/{total} files converted to {output_dir}\n", "info")

        self.root.after(0, finish)

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _log(self, message: str, tag: str = ""):
        self.log_text.config(state="normal")
        self.log_text.insert("end", message, tag)
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def _open_output_folder(self):
        folder = self.output_var.get().strip()
        if os.path.isdir(folder):
            os.startfile(folder)
        else:
            messagebox.showinfo("Not found", f"Folder does not exist:\n{folder}")

    def _set_icon(self):
        """Set a simple window icon (optional)."""
        try:
            # Create a simple colored icon
            icon = tk.PhotoImage(width=32, height=32)
            self.root.iconphoto(True, icon)
        except Exception:
            pass

    def run(self):
        self.root.mainloop()
