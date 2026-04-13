"""GUI application for document-to-Markdown conversion."""

from __future__ import annotations

import locale
import logging
import os
import sys
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

# Windows: ensure UTF-8 subprocess encoding for Korean/CJK paths
if sys.platform == "win32":
    os.environ.setdefault("PYTHONIOENCODING", "utf-8")

# Ensure bundled PyInstaller paths work
if getattr(sys, "frozen", False):
    _exe_dir = os.path.dirname(sys.executable)
    os.chdir(_exe_dir)

# Setup file logging for diagnostics
_log_path = Path.home() / "Documents" / "DocToMarkdown_error.log"
logging.basicConfig(
    filename=str(_log_path),
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    encoding="utf-8",
)
_logger = logging.getLogger("DocToMarkdown")

from convert_to_md import (
    MARKITDOWN_EXTENSIONS,
    LEGACY_OFFICE_EXTENSIONS,
    HWP_EXTENSIONS,
    PDF_EXTENSIONS,
    convert_one,
    is_supported,
)

ALL_SUPPORTED = PDF_EXTENSIONS | HWP_EXTENSIONS | LEGACY_OFFICE_EXTENSIONS | MARKITDOWN_EXTENSIONS
DEFAULT_OUTPUT_DIR = Path.home() / "Documents" / "ConvertedMD"

FILETYPES = [
    ("All supported", " ".join(f"*{ext}" for ext in sorted(ALL_SUPPORTED))),
    ("PDF", "*.pdf"),
    ("HWP", "*.hwp"),
    ("Word", "*.doc *.docx"),
    ("Excel", "*.xls *.xlsx"),
    ("PowerPoint", "*.ppt *.pptx"),
    ("Images", "*.jpg *.jpeg *.png *.gif *.bmp *.tif *.tiff"),
    ("Web/Text", "*.html *.htm *.csv *.json *.xml *.txt *.rtf"),
    ("All files", "*.*"),
]


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("PDF/Document to Markdown Converter")
        self.geometry("750x580")
        self.minsize(600, 480)
        self.resizable(True, True)

        self.selected_files: list[Path] = []
        self.output_dir = tk.StringVar(value=str(DEFAULT_OUTPUT_DIR))
        self.extract_images = tk.BooleanVar(value=True)
        self.converting = False

        self._build_ui()

    # ── UI Construction ──────────────────────────────────────────────

    def _build_ui(self) -> None:
        # Title
        title_frame = tk.Frame(self)
        title_frame.pack(fill="x", padx=12, pady=(12, 4))
        tk.Label(
            title_frame,
            text="Document to Markdown Converter",
            font=("Segoe UI", 14, "bold"),
        ).pack(anchor="w")
        tk.Label(
            title_frame,
            text="PDF, HWP, Office, images, and more -> Markdown",
            font=("Segoe UI", 9),
            fg="#666",
        ).pack(anchor="w")

        # ── Input section ────────────────────────────────────────────
        input_frame = tk.LabelFrame(self, text=" Input Files ", padx=8, pady=6)
        input_frame.pack(fill="both", expand=True, padx=12, pady=4)

        btn_row = tk.Frame(input_frame)
        btn_row.pack(fill="x", pady=(0, 4))

        tk.Button(btn_row, text="Add Files...", width=14, command=self._add_files).pack(side="left", padx=(0, 4))
        tk.Button(btn_row, text="Add Folder...", width=14, command=self._add_folder).pack(side="left", padx=(0, 4))
        tk.Button(btn_row, text="Clear All", width=10, command=self._clear_files).pack(side="left")
        self.file_count_label = tk.Label(btn_row, text="0 files selected", fg="#888")
        self.file_count_label.pack(side="right")

        list_frame = tk.Frame(input_frame)
        list_frame.pack(fill="both", expand=True)

        scrollbar = tk.Scrollbar(list_frame, orient="vertical")
        self.file_listbox = tk.Listbox(
            list_frame,
            selectmode="extended",
            yscrollcommand=scrollbar.set,
            font=("Consolas", 9),
        )
        scrollbar.config(command=self.file_listbox.yview)
        self.file_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        tk.Button(input_frame, text="Remove Selected", command=self._remove_selected).pack(anchor="e", pady=(4, 0))

        # ── Output section ───────────────────────────────────────────
        output_frame = tk.LabelFrame(self, text=" Output Location ", padx=8, pady=6)
        output_frame.pack(fill="x", padx=12, pady=4)

        self.use_default = tk.BooleanVar(value=True)
        tk.Radiobutton(
            output_frame,
            text=f"Default folder:  {DEFAULT_OUTPUT_DIR}",
            variable=self.use_default,
            value=True,
            command=self._toggle_output,
        ).pack(anchor="w")

        custom_row = tk.Frame(output_frame)
        custom_row.pack(fill="x", pady=(2, 0))
        tk.Radiobutton(
            custom_row,
            text="Custom folder:",
            variable=self.use_default,
            value=False,
            command=self._toggle_output,
        ).pack(side="left")
        self.output_entry = tk.Entry(custom_row, textvariable=self.output_dir, state="disabled")
        self.output_entry.pack(side="left", fill="x", expand=True, padx=4)
        self.browse_btn = tk.Button(custom_row, text="Browse...", command=self._browse_output, state="disabled")
        self.browse_btn.pack(side="left")

        # ── Options ──────────────────────────────────────────────────
        opt_frame = tk.Frame(self)
        opt_frame.pack(fill="x", padx=12, pady=4)
        tk.Checkbutton(opt_frame, text="Extract images from PDF/HWP", variable=self.extract_images).pack(anchor="w")

        # ── Convert button ───────────────────────────────────────────
        self.convert_btn = tk.Button(
            self,
            text="Convert",
            font=("Segoe UI", 11, "bold"),
            bg="#0078D4",
            fg="white",
            activebackground="#005A9E",
            activeforeground="white",
            height=1,
            command=self._start_conversion,
        )
        self.convert_btn.pack(fill="x", padx=12, pady=(4, 2))

        # ── Progress ─────────────────────────────────────────────────
        self.progress = ttk.Progressbar(self, mode="determinate")
        self.progress.pack(fill="x", padx=12, pady=(0, 2))

        self.status_label = tk.Label(self, text="Ready", anchor="w", fg="#444", font=("Segoe UI", 9))
        self.status_label.pack(fill="x", padx=12, pady=(0, 4))

        # ── Footer ──────────────────────────────────────────────────
        footer = tk.Label(
            self,
            text="made by \uc774\ub300\ud615 with Claude Code",
            font=("Segoe UI", 8),
            fg="#999",
        )
        footer.pack(side="bottom", pady=(0, 6))

    # ── Input handlers ───────────────────────────────────────────────

    @staticmethod
    def _safe_path(p: str | Path) -> Path:
        """Normalize and resolve a path to handle Korean/CJK characters and long Windows paths."""
        path = Path(os.path.normpath(str(p))).resolve()
        # Windows long path support (>260 chars)
        s = str(path)
        if sys.platform == "win32" and len(s) > 240 and not s.startswith("\\\\?\\"):
            path = Path("\\\\?\\" + s)
        return path

    def _add_files(self) -> None:
        paths = filedialog.askopenfilenames(title="Select files to convert", filetypes=FILETYPES)
        if not paths:
            return
        _logger.debug("filedialog returned %d path(s): %s", len(paths), paths)
        for p in paths:
            try:
                path = self._safe_path(p)
                _logger.debug("  safe_path: %r -> %r (exists=%s, supported=%s)",
                              p, path, path.exists(), is_supported(path))
                if path not in self.selected_files and path.exists() and is_supported(path):
                    self.selected_files.append(path)
            except Exception:
                _logger.exception("  _safe_path failed for: %r", p)
        self._refresh_list()

    def _add_folder(self) -> None:
        folder = filedialog.askdirectory(title="Select folder to scan")
        if not folder:
            return
        _logger.debug("askdirectory returned: %r", folder)
        try:
            folder_path = self._safe_path(folder)
            _logger.debug("  folder safe_path: %r (exists=%s)", folder_path, folder_path.exists())
            for f in sorted(folder_path.rglob("*")):
                try:
                    resolved = self._safe_path(f)
                    if resolved.is_file() and is_supported(resolved) and resolved not in self.selected_files:
                        self.selected_files.append(resolved)
                except Exception:
                    _logger.exception("  _safe_path failed for child: %r", f)
        except Exception:
            _logger.exception("  folder scan failed: %r", folder)
        self._refresh_list()

    def _clear_files(self) -> None:
        self.selected_files.clear()
        self._refresh_list()

    def _remove_selected(self) -> None:
        indices = list(self.file_listbox.curselection())
        if not indices:
            return
        for i in reversed(indices):
            self.selected_files.pop(i)
        self._refresh_list()

    def _refresh_list(self) -> None:
        self.file_listbox.delete(0, "end")
        for f in self.selected_files:
            ext = f.suffix.lower()
            size = f.stat().st_size
            size_str = f"{size / 1024:.0f}KB" if size >= 1024 else f"{size}B"
            self.file_listbox.insert("end", f"[{ext}]  {f.name}  ({size_str})  -  {f.parent}")
        self.file_count_label.config(text=f"{len(self.selected_files)} files selected")

    # ── Output handlers ──────────────────────────────────────────────

    def _toggle_output(self) -> None:
        if self.use_default.get():
            self.output_entry.config(state="disabled")
            self.browse_btn.config(state="disabled")
            self.output_dir.set(str(DEFAULT_OUTPUT_DIR))
        else:
            self.output_entry.config(state="normal")
            self.browse_btn.config(state="normal")

    def _browse_output(self) -> None:
        folder = filedialog.askdirectory(title="Select output folder")
        if folder:
            self.output_dir.set(folder)

    # ── Conversion ───────────────────────────────────────────────────

    def _start_conversion(self) -> None:
        if self.converting:
            return
        if not self.selected_files:
            messagebox.showwarning("No files", "Please add files to convert first.")
            return

        out_dir = self._safe_path(self.output_dir.get())
        try:
            out_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Error", f"Cannot create output folder:\n{e}")
            return

        self.converting = True
        self.convert_btn.config(state="disabled", text="Converting...")
        self.progress["maximum"] = len(self.selected_files)
        self.progress["value"] = 0

        thread = threading.Thread(
            target=self._convert_worker,
            args=(list(self.selected_files), out_dir, self.extract_images.get()),
            daemon=True,
        )
        thread.start()

    def _convert_worker(self, files: list[Path], out_dir: Path, extract_images: bool) -> None:
        ok = 0
        fail = 0
        failed_names: list[str] = []

        for i, path in enumerate(files):
            self.after(0, self._update_status, f"Converting {i + 1}/{len(files)}: {path.name}")
            _logger.info("Converting [%d/%d]: %s -> %s", i + 1, len(files), path, out_dir)
            try:
                convert_one(path, out_dir, extract_images=extract_images)
                ok += 1
                _logger.info("  OK: %s", path.name)
            except Exception as exc:
                fail += 1
                tb = traceback.format_exc()
                failed_names.append(f"{path.name}: {exc}")
                _logger.error("  FAIL: %s\n%s", path.name, tb)
            self.after(0, self._update_progress, i + 1)

        self.after(0, self._conversion_done, ok, fail, failed_names, out_dir)

    def _update_status(self, text: str) -> None:
        self.status_label.config(text=text)

    def _update_progress(self, value: int) -> None:
        self.progress["value"] = value

    def _conversion_done(self, ok: int, fail: int, failed: list[str], out_dir: Path) -> None:
        self.converting = False
        self.convert_btn.config(state="normal", text="Convert")
        self.status_label.config(text=f"Done - {ok} succeeded, {fail} failed")

        msg = f"Conversion complete!\n\nSuccess: {ok}\nFailed: {fail}\nOutput: {out_dir}"
        if failed:
            msg += "\n\nFailed files:\n" + "\n".join(f"  - {f}" for f in failed[:10])
            if len(failed) > 10:
                msg += f"\n  ... and {len(failed) - 10} more"

        if fail > 0:
            msg += f"\n\nError log: {_log_path}"
            messagebox.showwarning("Conversion Complete", msg)
        else:
            messagebox.showinfo("Conversion Complete", msg)

        # Open output folder
        if ok > 0:
            os.startfile(str(out_dir))


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
