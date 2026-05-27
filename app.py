from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from docx import Document

from paper_generator import ParseError, generate_professional_docx


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_OUTPUT = BASE_DIR / "output" / "專業文件初稿.docx"
SAMPLE_TEXT = """題目：企劃轉專業文件系統
英文題目：Project to Professional Document Formatter
作者：王小明
單位：資訊管理學系
年份：2026
月份：6

摘要
本系統可將整篇企劃、報告草稿或論文內容，自動整理為正式且可交付的 Word 文件格式。
關鍵詞：文件生成、排版、自動化

第一章 緒論
第一節 背景與動機
近年來，許多內容已經先寫好，但最後整理成正式文件的時間仍然很高。

第二節 研究目的
本工具希望讓使用者把整篇內容直接丟進去，就能得到完整格式的 Word 文件。

第二章 系統說明
第一節 核心流程
系統會盡量辨識章節、摘要、參考文獻、附錄與圖表標記。

參考文獻
林雍智（2020）。教育學門論文寫作格式指引：APA 格式第七版之應用。心理出版社。
"""


class ProfessionalFormatterApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("專業文件格式生成器")
        self.root.geometry("1200x820")
        self.output_var = tk.StringVar(value=str(DEFAULT_OUTPUT))
        self.status_var = tk.StringVar(value="把整篇內容直接貼進來，按右上角生成。")
        self._build_layout()

    def _build_layout(self) -> None:
        outer = ttk.Frame(self.root, padding=14)
        outer.pack(fill="both", expand=True)

        intro = ttk.LabelFrame(outer, text="使用方式", padding=12)
        intro.pack(fill="x")
        ttk.Label(
            intro,
            text=(
                "1. 直接把整篇內容貼到下面的大框\n"
                "2. 也可以直接開啟 txt、md、docx\n"
                "3. 按「生成 Word 文件」後，系統會整理成正式格式的 Word 文件"
            ),
            justify="left",
        ).pack(anchor="w")

        top_bar = ttk.Frame(outer)
        top_bar.pack(fill="x", pady=(10, 10))
        ttk.Label(top_bar, text="輸出位置").pack(side="left")
        ttk.Entry(top_bar, textvariable=self.output_var).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(top_bar, text="選擇位置", command=self.choose_output).pack(side="left")
        ttk.Button(top_bar, text="生成 Word 文件", command=self.generate).pack(side="right")

        actions = ttk.Frame(outer)
        actions.pack(fill="x", pady=(0, 10))
        ttk.Button(actions, text="開啟文字或 Word 檔", command=self.open_input_file).pack(side="left")
        ttk.Button(actions, text="載入範例", command=self.load_sample).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="清空", command=self.clear_text).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="另存目前內容", command=self.save_input_text).pack(side="left", padx=(8, 0))

        ttk.Label(outer, text="整篇內容").pack(anchor="w")
        self.editor = tk.Text(outer, wrap="word", font=("Microsoft JhengHei UI", 11))
        self.editor.pack(fill="both", expand=True)

        note = ttk.LabelFrame(outer, text="可選圖表格式", padding=12)
        note.pack(fill="x", pady=(10, 0))
        ttk.Label(
            note,
            text=(
                "如果你要插圖或表格，可以在內容裡加這種標記：\n"
                "[FIGURE title=\"流程圖\" path=\"C:/圖片/flow.png\" width_cm=\"13\"]\n"
                "註：圖片說明\n"
                "[/FIGURE]\n\n"
                "[TABLE title=\"研究規劃\"]\n"
                "| 階段 | 內容 |\n"
                "| --- | --- |\n"
                "| 第一階段 | 需求分析 |\n"
                "註：表格說明\n"
                "[/TABLE]"
            ),
            justify="left",
        ).pack(anchor="w")

        footer = ttk.Frame(outer)
        footer.pack(fill="x", pady=(10, 0))
        ttk.Label(footer, textvariable=self.status_var).pack(side="right")

        self.load_sample()

    def choose_output(self) -> None:
        file_path = filedialog.asksaveasfilename(
            title="選擇輸出位置",
            defaultextension=".docx",
            filetypes=[("Word 文件", "*.docx")],
            initialfile=DEFAULT_OUTPUT.name,
            initialdir=str(DEFAULT_OUTPUT.parent),
        )
        if file_path:
            self.output_var.set(file_path)

    def open_input_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="開啟內容檔",
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
        if path.suffix.lower() == ".docx":
            content = self._read_docx(path)
        else:
            content = path.read_text(encoding="utf-8")

        self.editor.delete("1.0", tk.END)
        self.editor.insert("1.0", content)
        self.status_var.set(f"已載入：{path.name}")

    def _read_docx(self, path: Path) -> str:
        doc = Document(str(path))
        paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        return "\n\n".join(paragraphs)

    def load_sample(self) -> None:
        self.editor.delete("1.0", tk.END)
        self.editor.insert("1.0", SAMPLE_TEXT)
        self.status_var.set("已載入範例，可以直接試產出。")

    def clear_text(self) -> None:
        if messagebox.askyesno("確認清空", "要清空目前內容嗎？"):
            self.editor.delete("1.0", tk.END)
            self.status_var.set("內容已清空。")

    def save_input_text(self) -> None:
        file_path = filedialog.asksaveasfilename(
            title="另存目前內容",
            defaultextension=".txt",
            filetypes=[("文字檔", "*.txt"), ("Markdown", "*.md")],
        )
        if not file_path:
            return
        Path(file_path).write_text(self.editor.get("1.0", tk.END), encoding="utf-8")
        self.status_var.set("目前內容已另存。")

    def generate(self) -> None:
        raw_text = self.editor.get("1.0", tk.END).strip()
        if not raw_text:
            messagebox.showwarning("沒有內容", "請先貼上整篇內容。")
            return
        try:
            output = generate_professional_docx(raw_text, self.output_var.get())
        except ParseError as exc:
            messagebox.showerror("格式解析失敗", str(exc))
            self.status_var.set("生成失敗，請檢查內容。")
            return
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("生成失敗", str(exc))
            self.status_var.set("生成失敗。")
            return
        self.status_var.set(f"已輸出：{output}")
        messagebox.showinfo("完成", f"已輸出 Word 文件：\n{output}")


def main() -> None:
    root = tk.Tk()
    ProfessionalFormatterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
