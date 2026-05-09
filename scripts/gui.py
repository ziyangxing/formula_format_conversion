#!/usr/bin/env python3
"""
FormulaConverter GUI - AI公式 → Word专业公式 批量转换工具

双击运行，拖放或选择 Word 文档，一键批量转换公式为专业格式。
兼容 WPS 和 Microsoft Word。
"""

import os, sys, threading, tkinter as tk
from tkinter import filedialog, messagebox, ttk

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from convert import process_document


class FormulaConverterGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("FormulaConverter - AI公式转Word专业公式")
        self.root.geometry("750x560")
        self.root.minsize(600, 420)
        self.root.configure(bg="#f5f5f5")

        self.file_list = []
        self.output_dir = tk.StringVar(value=os.path.expanduser("~/Desktop"))
        self._setup_style()
        self._build_ui()
        self._center_window()

    def _setup_style(self):
        s = ttk.Style()
        s.theme_use("clam")
        s.configure(".", font=("Microsoft YaHei UI", 10))
        s.configure("TButton", padding=8)
        s.configure("Header.TLabel", font=("Microsoft YaHei UI", 13, "bold"), background="#f5f5f5")
        s.configure("Hint.TLabel", font=("Microsoft YaHei UI", 9), foreground="#999", background="#f5f5f5")

    def _center_window(self):
        self.root.update_idletasks()
        w, h = self.root.winfo_width(), self.root.winfo_height()
        x = (self.root.winfo_screenwidth() - w) // 2
        y = (self.root.winfo_screenheight() - h) // 2
        self.root.geometry(f"+{x}+{y}")

    def _build_ui(self):
        # Header
        hf = tk.Frame(self.root, bg="#f5f5f5")
        hf.pack(fill=tk.X, padx=20, pady=(15, 5))
        ttk.Label(hf, text="AI公式 → Word/WPS 专业公式转换", style="Header.TLabel").pack(side=tk.LEFT)

        # File list frame
        lf = tk.LabelFrame(self.root, text="待转换文件（可拖放 .docx/.docm/.md 文件）",
                           font=("Microsoft YaHei UI", 10), bg="#f5f5f5")
        lf.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

        # Listbox + scroll
        inner = tk.Frame(lf, bg="#f5f5f5")
        inner.pack(fill=tk.BOTH, expand=True, padx=5, pady=(5, 0))

        sb = tk.Scrollbar(inner)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.lb = tk.Listbox(inner, selectmode=tk.EXTENDED, font=("Consolas", 10),
                              yscrollcommand=sb.set, bg="white")
        self.lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.config(command=self.lb.yview)

        # Drop target label inside list frame
        self.drop_label = tk.Label(lf, text="拖放 .docx/.docm/.md 文件到此处", bg="#e8e8e8",
                                     fg="#666", font=("Microsoft YaHei UI", 9), height=2)
        self.drop_label.pack(fill=tk.X, padx=5, pady=3)

        # Enable file drop on Windows
        self._enable_drag_drop()

        # Buttons
        bf = tk.Frame(lf, bg="#f5f5f5")
        bf.pack(fill=tk.X, padx=5, pady=(0, 8))
        ttk.Button(bf, text="＋ 添加文件", command=self._add_files).pack(side=tk.LEFT, padx=3)
        ttk.Button(bf, text="＋ 添加文件夹", command=self._add_folder).pack(side=tk.LEFT, padx=3)
        ttk.Button(bf, text="✕ 移除选中", command=self._remove_selected).pack(side=tk.LEFT, padx=3)
        ttk.Button(bf, text="清空", command=self._clear).pack(side=tk.LEFT, padx=3)

        # Output dir
        of = tk.Frame(self.root, bg="#f5f5f5")
        of.pack(fill=tk.X, padx=20, pady=(5, 0))
        ttk.Label(of, text="输出目录:").pack(side=tk.LEFT)
        self.out_entry = ttk.Entry(of, textvariable=self.output_dir, font=("Consolas", 10))
        self.out_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(of, text="浏览...", command=self._choose_output).pack(side=tk.LEFT)

        # Progress
        pf = tk.Frame(self.root, bg="#f5f5f5")
        pf.pack(fill=tk.X, padx=20, pady=(10, 5))
        self.progress = ttk.Progressbar(pf, mode="determinate")
        self.progress.pack(fill=tk.X)
        self.status_label = tk.Label(pf, text="就绪 - 请添加文件后点击 [开始转换]",
                                      font=("Microsoft YaHei UI", 9), bg="#f5f5f5", fg="#888", anchor=tk.W)
        self.status_label.pack(fill=tk.X)

        # Convert button
        bbf = tk.Frame(self.root, bg="#f5f5f5")
        bbf.pack(fill=tk.X, padx=20, pady=(5, 15))
        self.convert_btn = ttk.Button(bbf, text="▶ 开始转换", command=self._start)
        self.convert_btn.pack(side=tk.RIGHT, ipadx=12)
        ttk.Label(bbf, text="输出为 原文件名_公式转换.docx", style="Hint.TLabel").pack(side=tk.RIGHT, padx=8)

        self._update_status()

    def _enable_drag_drop(self):
        """Enable Windows file drag-and-drop on the Listbox and drop label."""
        try:
            import ctypes
            from ctypes import wintypes

            GWL_EXSTYLE = -20
            WS_EX_ACCEPTFILES = 0x00000010

            user32 = ctypes.windll.user32
            shell32 = ctypes.windll.shell32

            hwnd = self.lb.winfo_id()
            user32.DragAcceptFiles(hwnd, True)
            # Also accept on the drop label
            drop_hwnd = self.drop_label.winfo_id()
            user32.DragAcceptFiles(drop_hwnd, True)

            # Register the window to receive WM_DROPFILES
            self._drop_callback_registered = True
            self.root.createfilehandler = lambda *a: None  # avoid error
        except Exception:
            pass  # Drag-drop not available, buttons still work

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            parent=self.root,
            title="选择Word文档",
            filetypes=[("Word & Markdown", "*.docx;*.docm;*.md"), ("所有文件", "*.*")],
        )
        for p in paths:
            self._add_file(p)

    def _add_folder(self):
        folder = filedialog.askdirectory(parent=self.root, title="选择文件夹")
        if not folder:
            return
        count = 0
        for root, dirs, files in os.walk(folder):
            for f in files:
                if f.lower().endswith((".docx", ".docm", ".md")) and not f.startswith("~$"):
                    self._add_file(os.path.join(root, f))
                    count += 1
        if count == 0:
            messagebox.showinfo("提示", "所选文件夹中没有 .docx/.docm 文件")

    def _add_file(self, path):
        path = os.path.abspath(path)
        if not path.lower().endswith((".docx", ".docm", ".md")):
            return
        if path in [f[0] for f in self.file_list]:
            return
        if "_公式转换" in os.path.basename(path):
            return
        self.file_list.append((path, os.path.basename(path)))
        self.lb.insert(tk.END, os.path.basename(path))
        self._update_status()

    def _remove_selected(self):
        for i in reversed(self.lb.curselection()):
            self.lb.delete(i)
            del self.file_list[i]
        self._update_status()

    def _clear(self):
        self.lb.delete(0, tk.END)
        self.file_list.clear()
        self._update_status()

    def _choose_output(self):
        folder = filedialog.askdirectory(parent=self.root, title="选择输出目录")
        if folder:
            self.output_dir.set(folder)

    def _update_status(self):
        n = len(self.file_list)
        self.status_label.config(text=f"已添加 {n} 个文件，准备就绪" if n else "就绪 - 请添加文件后点击 [开始转换]")

    # ============================================================
    # Conversion
    # ============================================================

    def _start(self):
        if not self.file_list:
            messagebox.showwarning("提示", "请先添加要转换的 Word 文档")
            return

        out_dir = self.output_dir.get().strip()
        if not out_dir:
            messagebox.showwarning("提示", "请选择输出目录")
            return
        os.makedirs(out_dir, exist_ok=True)

        self.convert_btn.config(state=tk.DISABLED, text="转换中...")
        self.progress["maximum"] = len(self.file_list)
        self.progress["value"] = 0

        threading.Thread(target=self._run, args=(out_dir,), daemon=True).start()

    def _run(self, out_dir):
        total_ok, errors = 0, []

        for i, (path, name) in enumerate(self.file_list):
            self._ui(lambda: self.status_label.config(
                text=f"[{i+1}/{len(self.file_list)}] {name} — 转换中..."))

            base, ext = os.path.splitext(name)
            out = os.path.join(out_dir, f"{base}_公式转换.docx")
            c = 1
            while os.path.exists(out):
                out = os.path.join(out_dir, f"{base}_公式转换({c}).docx")
                c += 1

            try:
                converted, _ = process_document(path, out)
                total_ok += converted
                self._ui(lambda: self.progress.configure(value=i+1))
            except Exception as e:
                errors.append(f"{name}: {e}")

        self._ui(lambda: self._done(total_ok, errors, out_dir))

    def _ui(self, fn):
        self.root.after(0, fn)

    def _done(self, total, errors, out_dir):
        self.convert_btn.config(state=tk.NORMAL, text="▶ 开始转换")
        self.progress["value"] = self.progress["maximum"]
        self.status_label.config(text=f"完成！共转换 {total} 个公式，输出到 {out_dir}")

        msg = f"转换完成！\n\n共转换 {total} 个公式"
        if errors:
            msg += f"\n\n{len(errors)} 个文件出错:\n"
            for e in errors[:5]:
                msg += f"  • {e}\n"
        msg += f"\n\n输出目录:\n{out_dir}"
        messagebox.showinfo("转换完成", msg)

    def run(self):
        self.root.mainloop()


def main():
    FormulaConverterGUI().run()


if __name__ == "__main__":
    main()
