import bisect
import difflib
import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText


class TextCompareApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Text Compare Tool")
        self.root.geometry("1200x760")
        self.root.minsize(900, 560)

        self._syncing_scroll = False
        self.status_var = tk.StringVar(
            value="请在左右文本框中粘贴内容，然后点击“开始比对”。"
        )

        self._build_ui()
        self._configure_tags()
        self._bind_shortcuts()

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=12)
        container.pack(fill="both", expand=True)

        toolbar = ttk.Frame(container)
        toolbar.pack(fill="x", pady=(0, 10))

        ttk.Button(toolbar, text="开始比对", command=self.compare_texts).pack(
            side="left"
        )
        ttk.Button(toolbar, text="清除高亮", command=self.clear_highlights).pack(
            side="left", padx=(8, 0)
        )
        ttk.Button(toolbar, text="交换左右文本", command=self.swap_texts).pack(
            side="left", padx=(8, 0)
        )

        ttk.Label(
            toolbar,
            text="快捷键：Ctrl+Enter 开始比对",
            foreground="#555555",
        ).pack(side="right")

        editors_frame = ttk.Frame(container)
        editors_frame.pack(fill="both", expand=True)

        ttk.Label(editors_frame, text="左侧文本").grid(
            row=0, column=0, sticky="w", padx=(0, 6), pady=(0, 6)
        )
        ttk.Label(editors_frame, text="右侧文本").grid(
            row=0, column=1, sticky="w", padx=(6, 0), pady=(0, 6)
        )

        self.left_text = ScrolledText(
            editors_frame,
            wrap="word",
            undo=True,
            font=("Consolas", 11),
        )
        self.right_text = ScrolledText(
            editors_frame,
            wrap="word",
            undo=True,
            font=("Consolas", 11),
        )

        self.left_text.grid(row=1, column=0, sticky="nsew", padx=(0, 6))
        self.right_text.grid(row=1, column=1, sticky="nsew", padx=(6, 0))

        editors_frame.columnconfigure(0, weight=1)
        editors_frame.columnconfigure(1, weight=1)
        editors_frame.rowconfigure(1, weight=1)

        self.left_text.vbar.configure(command=lambda *args: self._on_scrollbar(*args))
        self.right_text.vbar.configure(command=lambda *args: self._on_scrollbar(*args))

        self.left_text.configure(
            yscrollcommand=lambda first, last: self._sync_vertical(
                self.left_text, self.right_text, first, last
            )
        )
        self.right_text.configure(
            yscrollcommand=lambda first, last: self._sync_vertical(
                self.right_text, self.left_text, first, last
            )
        )

        status_bar = ttk.Label(
            container,
            textvariable=self.status_var,
            anchor="w",
            padding=(4, 10, 4, 0),
        )
        status_bar.pack(fill="x")

    def _configure_tags(self) -> None:
        for widget in (self.left_text, self.right_text):
            widget.tag_configure("replace", background="#ffe08a")
            widget.tag_configure("delete", background="#ffb3b3")
            widget.tag_configure("insert", background="#b9f6ca")

    def _bind_shortcuts(self) -> None:
        self.root.bind("<Control-Return>", lambda event: self.compare_texts())

    def _on_scrollbar(self, *args: str) -> None:
        self.left_text.yview(*args)
        self.right_text.yview(*args)

    def _sync_vertical(
        self,
        source_widget: ScrolledText,
        target_widget: ScrolledText,
        first: str,
        last: str,
    ) -> None:
        source_widget.vbar.set(first, last)
        target_widget.vbar.set(first, last)

        if self._syncing_scroll:
            return

        self._syncing_scroll = True
        try:
            target_widget.yview_moveto(first)
        finally:
            self._syncing_scroll = False

    def clear_highlights(self) -> None:
        for widget in (self.left_text, self.right_text):
            for tag_name in ("replace", "delete", "insert"):
                widget.tag_remove(tag_name, "1.0", "end")
        self.status_var.set("高亮已清除。")

    def swap_texts(self) -> None:
        left_content = self.left_text.get("1.0", "end-1c")
        right_content = self.right_text.get("1.0", "end-1c")
        self.left_text.delete("1.0", "end")
        self.right_text.delete("1.0", "end")
        self.left_text.insert("1.0", right_content)
        self.right_text.insert("1.0", left_content)
        self.clear_highlights()
        self.status_var.set("左右文本已交换。")

    def compare_texts(self) -> None:
        left_content = self.left_text.get("1.0", "end-1c")
        right_content = self.right_text.get("1.0", "end-1c")
        left_line_starts = self._build_line_starts(left_content)
        right_line_starts = self._build_line_starts(right_content)

        self.clear_highlights()

        matcher = difflib.SequenceMatcher(a=left_content, b=right_content)
        opcodes = matcher.get_opcodes()

        diff_count = 0
        for tag, i1, i2, j1, j2 in opcodes:
            if tag == "equal":
                continue

            diff_count += 1
            if tag == "replace":
                self._tag_range(
                    self.left_text,
                    left_line_starts,
                    i1,
                    i2,
                    "replace",
                )
                self._tag_range(
                    self.right_text,
                    right_line_starts,
                    j1,
                    j2,
                    "replace",
                )
            elif tag == "delete":
                self._tag_range(
                    self.left_text,
                    left_line_starts,
                    i1,
                    i2,
                    "delete",
                )
            elif tag == "insert":
                self._tag_range(
                    self.right_text,
                    right_line_starts,
                    j1,
                    j2,
                    "insert",
                )

        if diff_count == 0:
            self.status_var.set("两段文字完全一致。")
        else:
            self.status_var.set(
                f"发现 {diff_count} 处差异。黄色表示替换，红色表示左侧多出，绿色表示右侧多出。"
            )

    def _tag_range(
        self,
        widget: ScrolledText,
        line_starts: list[int],
        start_offset: int,
        end_offset: int,
        tag_name: str,
    ) -> None:
        if start_offset == end_offset:
            return

        start_index = self._offset_to_index(line_starts, start_offset)
        end_index = self._offset_to_index(line_starts, end_offset)
        widget.tag_add(tag_name, start_index, end_index)

    @staticmethod
    def _build_line_starts(content: str) -> list[int]:
        line_starts = [0]
        for index, char in enumerate(content):
            if char == "\n":
                line_starts.append(index + 1)
        return line_starts

    @staticmethod
    def _offset_to_index(line_starts: list[int], offset: int) -> str:
        line_number = bisect.bisect_right(line_starts, offset) - 1
        column = offset - line_starts[line_number]
        return f"{line_number + 1}.{column}"


def main() -> None:
    root = tk.Tk()
    style = ttk.Style()
    if "vista" in style.theme_names():
        style.theme_use("vista")
    app = TextCompareApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
