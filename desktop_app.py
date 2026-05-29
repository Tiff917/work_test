from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from docx import Document

from paper_generator import ParseError, TEMPLATE_OPTIONS, generate_professional_docx


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_OUTPUT_DIR = BASE_DIR / "output"
SAMPLE_INPUT_PATH = BASE_DIR / "sample_input.txt"


class DesktopFormatterApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("專業文件格式生成器 - 桌面版")
        self.root.geometry("1180x820")
        self.root.minsize(960, 700)

        self.template_map = {item["label"]: item for item in TEMPLATE_OPTIONS}
        self.template_var = tk.StringVar(value=TEMPLATE_OPTIONS[0]["label"])
        self.file_name_var = tk.StringVar(value="專業文件初稿")
        self.status_var = tk.StringVar(value="貼上內容後，按右上角生成，系統會先跳出存檔視窗。")

        self._build_layout()
        self.load_sample()

    def _build_layout(self) -> None:
        outer = ttk.Frame(self.root, padding=14)
        outer.pack(fill="both", expand=True)

        top = ttk.LabelFrame(outer, text="使用方式", padding=12)
        top.pack(fill="x")
        ttk.Label(
            top,
            text=(
                "1. 選模板\n"
                "2. 貼上整篇內容\n"
                "3. 按「生成並選擇儲存位置」\n"
                "4. 會直接跳出 Windows 存檔視窗"
            ),
            justify="left",
        ).pack(anchor="w")

        content = ttk.Panedwindow(outer, orient=tk.HORIZONTAL)
        content.pack(fill="both", expand=True, pady=(12, 0))

        left = ttk.Frame(content, padding=(0, 0, 12, 0))
        right = ttk.Frame(content)
        content.add(left, weight=3)
        content.add(right, weight=7)

        self._build_sidebar(left)
        self._build_editor(right)

        footer = ttk.Frame(outer)
        footer.pack(fill="x", pady=(10, 0))
        ttk.Label(footer, textvariable=self.status_var).pack(side="right")

    def _build_sidebar(self, parent: ttk.Frame) -> None:
        panel = ttk.LabelFrame(parent, text="模板與操作", padding=12)
        panel.pack(fill="both", expand=True)

        ttk.Label(panel, text="文件模板").pack(anchor="w")
        self.template_box = ttk.Combobox(
            panel,
            state="readonly",
            textvariable=self.template_var,
            values=[item["label"] for item in TEMPLATE_OPTIONS],
        )
        self.template_box.pack(fill="x", pady=(6, 0))
        self.template_box.bind("<<ComboboxSelected>>", self._on_template_change)

        self.template_hint = tk.Text(panel, height=5, wrap="word", font=("Microsoft JhengHei UI", 10))
        self.template_hint.pack(fill="x", pady=(10, 0))
        self.template_hint.configure(state="disabled")

        ttk.Label(panel, text="建議檔名").pack(anchor="w", pady=(16, 0))
        ttk.Entry(panel, textvariable=self.file_name_var).pack(fill="x", pady=(6, 0))

        button_bar = ttk.Frame(panel)
        button_bar.pack(fill="x", pady=(18, 0))
        ttk.Button(button_bar, text="載入範例", command=self.load_sample).pack(fill="x")
        ttk.Button(button_bar, text="開啟 txt / md / docx", command=self.open_input_file).pack(fill="x", pady=(8, 0))
        ttk.Button(button_bar, text="清空內容", command=self.clear_content).pack(fill="x", pady=(8, 0))
        ttk.Button(button_bar, text="生成並選擇儲存位置", command=self.generate_document).pack(fill="x", pady=(16, 0))

        tips = ttk.LabelFrame(panel, text="小提醒", padding=12)
        tips.pack(fill="x", pady=(18, 0))
        ttk.Label(
            tips,
            text=(
                "桌面版會使用 Windows 原生存檔視窗。\n"
                "也就是你每次按生成，都能自己決定要存到哪個資料夾。"
            ),
            justify="left",
        ).pack(anchor="w")

        self._update_template_hint()

    def _build_editor(self, parent: ttk.Frame) -> None:
        panel = ttk.LabelFrame(parent, text="內容輸入", padding=12)
        panel.pack(fill="both", expand=True)

        ttk.Label(
            panel,
            text="直接貼整篇內容。原文若已有「第一章、第一節、摘要、參考文獻」，系統會優先辨識。",
        ).pack(anchor="w")

        self.editor = tk.Text(panel, wrap="word", font=("Microsoft JhengHei UI", 11))
        self.editor.pack(fill="both", expand=True, pady=(10, 0))

    def _on_template_change(self, _event=None) -> None:
        self._update_template_hint()
        current = self.template_key
        if current == "proposal":
            self.file_name_var.set("商業提案初稿")
        elif current == "report":
            self.file_name_var.set("專題報告初稿")
        else:
            self.file_name_var.set("論文初稿")

    @property
    def template_key(self) -> str:
        label = self.template_var.get()
        return self.template_map[label]["key"]

    def _update_template_hint(self) -> None:
        label = self.template_var.get()
        item = self.template_map[label]
        hint = f"{item['label']}：{item['hint']}"
        self.template_hint.configure(state="normal")
        self.template_hint.delete("1.0", tk.END)
        self.template_hint.insert("1.0", hint)
        self.template_hint.configure(state="disabled")

    def load_sample(self) -> None:
        content = SAMPLE_INPUT_PATH.read_text(encoding="utf-8") if SAMPLE_INPUT_PATH.exists() else ""
        self.editor.delete("1.0", tk.END)
        self.editor.insert("1.0", content)
        self.status_var.set("已載入範例內容。")

    def open_input_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="選擇要載入的內容檔",
            filetypes=[
                ("支援檔案", "*.txt;*.md;*.docx"),
                ("文字檔", "*.txt"),
                ("Markdown", "*.md"),
                ("Word 文件", "*.docx"),
                ("所有檔案", "*.*"),
            ],
        )
        if not file_path:
            return

        path = Path(file_path)
        try:
            if path.suffix.lower() == ".docx":
                content = self._read_docx(path)
            else:
                content = path.read_text(encoding="utf-8")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("載入失敗", str(exc))
            return

        self.editor.delete("1.0", tk.END)
        self.editor.insert("1.0", content)
        self.status_var.set(f"已載入：{path.name}")

    def _read_docx(self, path: Path) -> str:
        doc = Document(str(path))
        paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        return "\n\n".join(paragraphs)

    def clear_content(self) -> None:
        if messagebox.askyesno("確認清空", "要清空目前內容嗎？"):
            self.editor.delete("1.0", tk.END)
            self.status_var.set("內容已清空。")

    def generate_document(self) -> None:
        raw_text = self.editor.get("1.0", tk.END).strip()
        if not raw_text:
            messagebox.showwarning("沒有內容", "請先貼上內容。")
            return

        DEFAULT_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        suggested_name = f"{self.file_name_var.get().strip() or '專業文件初稿'}.docx"
        file_path = filedialog.asksaveasfilename(
            title="選擇要儲存的位置",
            initialdir=str(DEFAULT_OUTPUT_DIR),
            initialfile=suggested_name,
            defaultextension=".docx",
            filetypes=[("Word 文件", "*.docx")],
        )
        if not file_path:
            self.status_var.set("已取消儲存。")
            return

        try:
            output = generate_professional_docx(raw_text, file_path, template_key=self.template_key)
        except ParseError as exc:
            messagebox.showerror("格式解析失敗", str(exc))
            self.status_var.set("生成失敗，請檢查內容格式。")
            return
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("生成失敗", str(exc))
            self.status_var.set("生成失敗。")
            return

        self.status_var.set(f"已儲存：{output}")
        messagebox.showinfo("完成", f"文件已儲存到：\n{output}")


def main() -> None:
    root = tk.Tk()
    DesktopFormatterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
