from __future__ import annotations

import copy
import difflib
import os
import re
import tkinter as tk
import zipfile
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
from xml.etree import ElementTree as ET


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
NS = {"w": WORD_NS}

ET.register_namespace("w", WORD_NS)


@dataclass
class ParagraphRecord:
    order_index: int
    source_index: int
    text: str
    normalized_text: str


@dataclass
class DocumentContent:
    path: Path
    file_type: str
    paragraphs: list[ParagraphRecord]


@dataclass
class DifferenceItem:
    item_id: int
    similarity: float
    standard_doc_label: str
    standard_paragraph: ParagraphRecord
    target_doc_label: str
    target_paragraph: ParagraphRecord
    original_target_text: str
    corrected: bool = False

    @property
    def active_target_text(self) -> str:
        if self.corrected:
            return self.standard_paragraph.text
        return self.original_target_text


class TextCompareApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Text Compare Tool Ver2")
        self.root.geometry("1480x860")
        self.root.minsize(1180, 720)

        self.file_a_var = tk.StringVar()
        self.file_b_var = tk.StringVar()
        self.standard_var = tk.StringVar(value="A")
        self.threshold_var = tk.DoubleVar(value=0.68)
        self.status_var = tk.StringVar(
            value="请选择两个文件并点击“开始分析”。支持 docx 和 txt，推荐使用 docx。"
        )
        self.summary_var = tk.StringVar(value="尚未分析。")
        self.selection_var = tk.StringVar(value="当前未选中差异项。")

        self.documents: dict[str, DocumentContent] = {}
        self.differences: list[DifferenceItem] = []
        self.corrected_targets: dict[int, str] = {}
        self.current_index = -1
        self.last_target_label = "B"
        self.matched_pair_count = 0
        self.exact_match_count = 0

        self._build_ui()
        self._configure_preview_tags()

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=12)
        container.pack(fill="both", expand=True)

        self._build_top_panel(container)

        body = ttk.Panedwindow(container, orient="horizontal")
        body.pack(fill="both", expand=True, pady=(10, 0))

        preview_frame = ttk.Frame(body, padding=10)
        sidebar_frame = ttk.Frame(body, padding=10)
        body.add(preview_frame, weight=4)
        body.add(sidebar_frame, weight=2)

        self._build_preview_panel(preview_frame)
        self._build_sidebar(sidebar_frame)

        status_bar = ttk.Label(
            container,
            textvariable=self.status_var,
            anchor="w",
            padding=(4, 10, 4, 0),
        )
        status_bar.pack(fill="x")

    def _build_top_panel(self, parent: ttk.Frame) -> None:
        top = ttk.LabelFrame(parent, text="文件与分析设置", padding=10)
        top.pack(fill="x")

        ttk.Label(top, text="文件 A").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.file_a_var).grid(
            row=0, column=1, sticky="ew", padx=(8, 8)
        )
        ttk.Button(top, text="选择文件", command=lambda: self._choose_file("A")).grid(
            row=0, column=2, padx=(0, 8)
        )
        ttk.Radiobutton(
            top,
            text="以 A 为标准",
            value="A",
            variable=self.standard_var,
            command=self._on_standard_changed,
        ).grid(row=0, column=3, sticky="w")

        ttk.Label(top, text="文件 B").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.file_b_var).grid(
            row=1, column=1, sticky="ew", padx=(8, 8), pady=(8, 0)
        )
        ttk.Button(top, text="选择文件", command=lambda: self._choose_file("B")).grid(
            row=1, column=2, padx=(0, 8), pady=(8, 0)
        )
        ttk.Radiobutton(
            top,
            text="以 B 为标准",
            value="B",
            variable=self.standard_var,
            command=self._on_standard_changed,
        ).grid(row=1, column=3, sticky="w", pady=(8, 0))

        ttk.Label(top, text="高相似阈值").grid(
            row=0, column=4, sticky="e", padx=(16, 6)
        )
        ttk.Spinbox(
            top,
            from_=0.50,
            to=0.95,
            increment=0.01,
            textvariable=self.threshold_var,
            width=8,
            format="%.2f",
        ).grid(row=0, column=5, sticky="w")

        ttk.Button(top, text="开始分析", command=self.analyze_documents).grid(
            row=1, column=4, padx=(16, 8), pady=(8, 0)
        )
        ttk.Button(top, text="重置结果", command=self.reset_results).grid(
            row=1, column=5, sticky="w", pady=(8, 0)
        )

        top.columnconfigure(1, weight=1)

    def _build_preview_panel(self, parent: ttk.Frame) -> None:
        header = ttk.Frame(parent)
        header.pack(fill="x")

        ttk.Label(
            header,
            text="左侧预览",
            font=("Segoe UI", 11, "bold"),
        ).pack(side="left")
        ttk.Label(
            header,
            textvariable=self.selection_var,
            foreground="#555555",
        ).pack(side="right")

        preview_pane = ttk.Panedwindow(parent, orient="vertical")
        preview_pane.pack(fill="both", expand=True, pady=(8, 0))

        top_box = ttk.LabelFrame(preview_pane, text="标准段落", padding=8)
        bottom_box = ttk.LabelFrame(preview_pane, text="待修正段落", padding=8)
        preview_pane.add(top_box, weight=1)
        preview_pane.add(bottom_box, weight=1)

        self.standard_preview = ScrolledText(
            top_box,
            wrap="word",
            font=("Consolas", 11),
            state="disabled",
        )
        self.standard_preview.pack(fill="both", expand=True)

        self.target_preview = ScrolledText(
            bottom_box,
            wrap="word",
            font=("Consolas", 11),
            state="disabled",
        )
        self.target_preview.pack(fill="both", expand=True)

    def _build_sidebar(self, parent: ttk.Frame) -> None:
        ttk.Label(
            parent,
            text="右侧统计与操作",
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w")

        ttk.Label(
            parent,
            textvariable=self.summary_var,
            justify="left",
            foreground="#333333",
            wraplength=360,
        ).pack(fill="x", pady=(8, 10))

        list_frame = ttk.LabelFrame(parent, text="差异项列表", padding=8)
        list_frame.pack(fill="both", expand=True)

        self.diff_listbox = tk.Listbox(
            list_frame,
            activestyle="dotbox",
            font=("Consolas", 10),
        )
        self.diff_listbox.pack(fill="both", expand=True)
        self.diff_listbox.bind("<<ListboxSelect>>", self._on_listbox_select)

        nav_frame = ttk.Frame(parent)
        nav_frame.pack(fill="x", pady=(10, 0))
        ttk.Button(nav_frame, text="上一条", command=self.select_previous).pack(
            side="left"
        )
        ttk.Button(nav_frame, text="下一条", command=self.select_next).pack(
            side="left", padx=(8, 0)
        )

        action_frame = ttk.LabelFrame(parent, text="自动纠正", padding=8)
        action_frame.pack(fill="x", pady=(10, 0))

        ttk.Button(action_frame, text="纠正当前项", command=self.apply_current).pack(
            fill="x"
        )
        ttk.Button(action_frame, text="一键纠正全部", command=self.apply_all).pack(
            fill="x", pady=(8, 0)
        )
        ttk.Button(action_frame, text="导出修正结果", command=self.export_corrected).pack(
            fill="x", pady=(8, 0)
        )

    def _configure_preview_tags(self) -> None:
        for widget in (self.standard_preview, self.target_preview):
            widget.tag_configure("replace", background="#ffe08a")
            widget.tag_configure("delete", background="#ffb3b3")
            widget.tag_configure("insert", background="#b9f6ca")

    def _choose_file(self, label: str) -> None:
        path = filedialog.askopenfilename(
            title=f"选择文件 {label}",
            filetypes=[
                ("支持的文件", "*.docx *.txt"),
                ("Word 文档", "*.docx"),
                ("文本文件", "*.txt"),
                ("所有文件", "*.*"),
            ],
        )
        if not path:
            return

        if label == "A":
            self.file_a_var.set(path)
        else:
            self.file_b_var.set(path)

    def _on_standard_changed(self) -> None:
        if self.documents:
            self.corrected_targets = {}
            self._rebuild_difference_view()

    def analyze_documents(self) -> None:
        file_a = self.file_a_var.get().strip()
        file_b = self.file_b_var.get().strip()
        if not file_a or not file_b:
            messagebox.showwarning("缺少文件", "请先选择文件 A 和文件 B。")
            return

        try:
            threshold = float(self.threshold_var.get())
        except (TypeError, ValueError):
            messagebox.showwarning("阈值无效", "请输入 0.50 到 0.95 之间的阈值。")
            return

        if not 0.50 <= threshold <= 0.95:
            messagebox.showwarning("阈值无效", "请输入 0.50 到 0.95 之间的阈值。")
            return

        try:
            self.documents["A"] = self._load_document(Path(file_a))
            self.documents["B"] = self._load_document(Path(file_b))
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("读取失败", f"文件读取失败：\n{exc}")
            return

        if not self.documents["A"].paragraphs or not self.documents["B"].paragraphs:
            messagebox.showwarning("无可比对段落", "至少有一个文件未读取到有效自然段。")
            return

        self.corrected_targets = {}
        self._build_differences(threshold)
        self._refresh_sidebar()

        if self.differences:
            self._select_item(0)
            self.status_var.set(
                f"分析完成。已识别 {len(self.differences)} 处可纠正的相似段落差异。"
            )
        else:
            self._clear_preview()
            self.status_var.set("分析完成。未找到需要纠正的高相似段落差异。")

    def reset_results(self) -> None:
        self.documents = {}
        self.differences = []
        self.corrected_targets = {}
        self.current_index = -1
        self.matched_pair_count = 0
        self.exact_match_count = 0
        self.diff_listbox.delete(0, "end")
        self._clear_preview()
        self.summary_var.set("尚未分析。")
        self.selection_var.set("当前未选中差异项。")
        self.status_var.set("结果已重置。请选择两个文件后重新分析。")

    def select_previous(self) -> None:
        if not self.differences:
            return
        next_index = len(self.differences) - 1 if self.current_index <= 0 else self.current_index - 1
        self._select_item(next_index)

    def select_next(self) -> None:
        if not self.differences:
            return
        next_index = 0 if self.current_index >= len(self.differences) - 1 else self.current_index + 1
        self._select_item(next_index)

    def apply_current(self) -> None:
        if not self._has_selected_item():
            messagebox.showinfo("无选中项", "请先在右侧列表中选择一个差异项。")
            return

        item = self.differences[self.current_index]
        self._apply_difference(item)
        self._refresh_sidebar()
        self._select_item(self.current_index)
        self.status_var.set(
            f"已按“{item.standard_doc_label}”的内容纠正当前段落，请继续预览或导出结果。"
        )

    def apply_all(self) -> None:
        if not self.differences:
            messagebox.showinfo("无差异项", "当前没有可自动纠正的差异项。")
            return

        for item in self.differences:
            self._apply_difference(item)

        self._refresh_sidebar()
        if self.differences:
            self._select_item(0)
        self.status_var.set("已完成全部自动纠正，请点击“导出修正结果”保存文件。")

    def export_corrected(self) -> None:
        if not self.documents or not self.differences:
            messagebox.showinfo("暂无结果", "请先完成分析后再导出。")
            return

        target_label = "B" if self.standard_var.get() == "A" else "A"
        target_document = self.documents[target_label]
        replacements = {
            item.target_paragraph.source_index: item.standard_paragraph.text
            for item in self.differences
            if item.corrected
        }

        if not replacements:
            messagebox.showinfo("尚未纠正", "请先执行“纠正当前项”或“一键纠正全部”。")
            return

        default_name = (
            f"{target_document.path.stem}_corrected{target_document.path.suffix}"
        )
        output_path = filedialog.asksaveasfilename(
            title="导出修正结果",
            initialfile=default_name,
            defaultextension=target_document.path.suffix,
            filetypes=[
                ("Word 文档", "*.docx"),
                ("文本文件", "*.txt"),
                ("所有文件", "*.*"),
            ],
        )
        if not output_path:
            return

        try:
            output = Path(output_path)
            if target_document.file_type == "docx":
                self._write_docx_with_replacements(target_document.path, output, replacements)
            else:
                self._write_text_with_replacements(target_document, output, replacements)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("导出失败", f"无法导出修正结果：\n{exc}")
            return

        self.status_var.set(f"修正结果已导出到：{output_path}")
        messagebox.showinfo("导出成功", f"修正结果已保存：\n{output_path}")

    def _load_document(self, path: Path) -> DocumentContent:
        if not path.exists():
            raise FileNotFoundError(f"文件不存在：{path}")

        suffix = path.suffix.lower()
        if suffix == ".docx":
            paragraphs = self._load_docx_paragraphs(path)
            return DocumentContent(path=path, file_type="docx", paragraphs=paragraphs)
        if suffix == ".txt":
            paragraphs = self._load_text_paragraphs(path)
            return DocumentContent(path=path, file_type="txt", paragraphs=paragraphs)

        raise ValueError(f"暂不支持该文件类型：{path.suffix}")

    def _load_docx_paragraphs(self, path: Path) -> list[ParagraphRecord]:
        with zipfile.ZipFile(path, "r") as archive:
            xml_bytes = archive.read("word/document.xml")

        root = ET.fromstring(xml_bytes)
        paragraph_elements = root.findall(".//w:p", NS)
        paragraphs: list[ParagraphRecord] = []

        for source_index, paragraph in enumerate(paragraph_elements):
            text = self._extract_docx_paragraph_text(paragraph)
            if not text.strip():
                continue
            paragraphs.append(
                ParagraphRecord(
                    order_index=len(paragraphs),
                    source_index=source_index,
                    text=text,
                    normalized_text=self._normalize_text(text),
                )
            )
        return paragraphs

    def _load_text_paragraphs(self, path: Path) -> list[ParagraphRecord]:
        content = path.read_text(encoding="utf-8")
        chunks = re.split(r"\n\s*\n", content)
        paragraphs: list[ParagraphRecord] = []

        for source_index, chunk in enumerate(chunks):
            text = chunk.strip()
            if not text:
                continue
            paragraphs.append(
                ParagraphRecord(
                    order_index=len(paragraphs),
                    source_index=source_index,
                    text=text,
                    normalized_text=self._normalize_text(text),
                )
            )
        return paragraphs

    def _build_differences(self, threshold: float) -> None:
        doc_a = self.documents["A"]
        doc_b = self.documents["B"]
        matches = self._match_paragraphs(doc_a.paragraphs, doc_b.paragraphs, threshold)
        standard_label = self.standard_var.get()
        standard_doc = self.documents[standard_label]
        target_label = "B" if standard_label == "A" else "A"
        target_doc = self.documents[target_label]
        self.last_target_label = target_label
        self.matched_pair_count = len(matches)
        self.exact_match_count = 0

        differences: list[DifferenceItem] = []
        for item_id, (para_a, para_b, similarity) in enumerate(matches):
            standard_paragraph = para_a if standard_label == "A" else para_b
            target_paragraph = para_b if standard_label == "A" else para_a

            if standard_paragraph.normalized_text == target_paragraph.normalized_text:
                self.exact_match_count += 1
                continue

            differences.append(
                DifferenceItem(
                    item_id=item_id,
                    similarity=similarity,
                    standard_doc_label=f"{standard_label}: {standard_doc.path.name}",
                    standard_paragraph=standard_paragraph,
                    target_doc_label=f"{target_label}: {target_doc.path.name}",
                    target_paragraph=target_paragraph,
                    original_target_text=target_paragraph.text,
                )
            )

        self.differences = differences
        self.current_index = -1

    def _match_paragraphs(
        self,
        paragraphs_a: list[ParagraphRecord],
        paragraphs_b: list[ParagraphRecord],
        threshold: float,
    ) -> list[tuple[ParagraphRecord, ParagraphRecord, float]]:
        candidates: list[tuple[float, int, int, ParagraphRecord, ParagraphRecord]] = []

        for para_a in paragraphs_a:
            for para_b in paragraphs_b:
                similarity = difflib.SequenceMatcher(
                    None,
                    para_a.normalized_text,
                    para_b.normalized_text,
                ).ratio()
                if similarity < threshold:
                    continue

                distance = abs(para_a.order_index - para_b.order_index)
                candidates.append((similarity, -distance, para_a.order_index, para_a, para_b))

        candidates.sort(key=lambda item: (-item[0], -item[1], item[2]))

        used_a: set[int] = set()
        used_b: set[int] = set()
        matches: list[tuple[ParagraphRecord, ParagraphRecord, float]] = []

        for similarity, _distance_key, _order, para_a, para_b in candidates:
            if para_a.source_index in used_a or para_b.source_index in used_b:
                continue
            used_a.add(para_a.source_index)
            used_b.add(para_b.source_index)
            matches.append((para_a, para_b, similarity))

        matches.sort(key=lambda item: item[0].order_index)
        return matches

    def _refresh_sidebar(self) -> None:
        self.diff_listbox.delete(0, "end")

        corrected_count = sum(1 for item in self.differences if item.corrected)
        pending_count = len(self.differences) - corrected_count
        standard_label = self.standard_var.get()
        target_label = "B" if standard_label == "A" else "A"

        if self.documents:
            summary = (
                f"标准文件：{self.documents[standard_label].path.name}\n"
                f"待修正文件：{self.documents[target_label].path.name}\n"
                f"已匹配高相似段落：{self.matched_pair_count}\n"
                f"其中完全一致：{self.exact_match_count}\n"
                f"待纠正差异：{pending_count}\n"
                f"已纠正：{corrected_count}"
            )
        else:
            summary = "尚未分析。"
        self.summary_var.set(summary)

        for item in self.differences:
            status = "已纠正" if item.corrected else "待处理"
            display_text = (
                f"[{status}] 标准段 {item.standard_paragraph.order_index + 1} -> "
                f"目标段 {item.target_paragraph.order_index + 1} | "
                f"相似度 {item.similarity:.0%}"
            )
            self.diff_listbox.insert("end", display_text)

    def _select_item(self, index: int) -> None:
        if not self.differences:
            self.current_index = -1
            self._clear_preview()
            return

        index = max(0, min(index, len(self.differences) - 1))
        self.current_index = index
        self.diff_listbox.selection_clear(0, "end")
        self.diff_listbox.selection_set(index)
        self.diff_listbox.activate(index)
        self.diff_listbox.see(index)
        self._update_preview(self.differences[index])

    def _on_listbox_select(self, _event: tk.Event) -> None:
        if not self.diff_listbox.curselection():
            return
        self._select_item(int(self.diff_listbox.curselection()[0]))

    def _update_preview(self, item: DifferenceItem) -> None:
        status = "已纠正" if item.corrected else "待处理"
        self.selection_var.set(
            f"第 {self.current_index + 1}/{len(self.differences)} 项 | {status} | 相似度 {item.similarity:.1%}"
        )
        self._fill_preview(
            self.standard_preview,
            item.standard_paragraph.text,
            item.active_target_text,
        )
        self._fill_preview(
            self.target_preview,
            item.active_target_text,
            item.standard_paragraph.text,
        )

    def _fill_preview(
        self,
        widget: ScrolledText,
        text: str,
        other_text: str,
    ) -> None:
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", text)
        for tag_name in ("replace", "delete", "insert"):
            widget.tag_remove(tag_name, "1.0", "end")

        self._highlight_text_diff(widget, text, other_text)
        widget.configure(state="disabled")

    def _highlight_text_diff(
        self,
        widget: ScrolledText,
        text: str,
        other_text: str,
    ) -> None:
        matcher = difflib.SequenceMatcher(a=text, b=other_text)
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == "equal":
                continue
            if tag == "replace":
                widget.tag_add("replace", f"1.0+{i1}c", f"1.0+{i2}c")
            elif tag == "delete":
                widget.tag_add("delete", f"1.0+{i1}c", f"1.0+{i2}c")

    def _apply_difference(self, item: DifferenceItem) -> None:
        item.corrected = True
        self.corrected_targets[item.target_paragraph.source_index] = item.standard_paragraph.text

    def _has_selected_item(self) -> bool:
        return 0 <= self.current_index < len(self.differences)

    def _clear_preview(self) -> None:
        self.selection_var.set("当前未选中差异项。")
        for widget in (self.standard_preview, self.target_preview):
            widget.configure(state="normal")
            widget.delete("1.0", "end")
            widget.configure(state="disabled")

    def _rebuild_difference_view(self) -> None:
        if not self.documents:
            return

        self._build_differences(float(self.threshold_var.get()))
        self._refresh_sidebar()
        if self.differences:
            self._select_item(0)
        else:
            self._clear_preview()

    @staticmethod
    def _normalize_text(text: str) -> str:
        return re.sub(r"\s+", " ", text).strip()

    @staticmethod
    def _extract_docx_paragraph_text(paragraph: ET.Element) -> str:
        texts = []
        for node in paragraph.findall(".//w:t", NS):
            texts.append(node.text or "")
        return "".join(texts)

    def _write_docx_with_replacements(
        self,
        source_path: Path,
        output_path: Path,
        replacements: dict[int, str],
    ) -> None:
        with zipfile.ZipFile(source_path, "r") as archive:
            files = {name: archive.read(name) for name in archive.namelist()}

        root = ET.fromstring(files["word/document.xml"])
        paragraph_elements = root.findall(".//w:p", NS)

        for source_index, new_text in replacements.items():
            paragraph = paragraph_elements[source_index]
            self._set_docx_paragraph_text(paragraph, new_text)

        files["word/document.xml"] = ET.tostring(
            root,
            encoding="utf-8",
            xml_declaration=True,
        )

        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for name, content in files.items():
                archive.writestr(name, content)

    def _write_text_with_replacements(
        self,
        document: DocumentContent,
        output_path: Path,
        replacements: dict[int, str],
    ) -> None:
        paragraphs_by_index: dict[int, str] = {}
        max_index = -1
        for paragraph in document.paragraphs:
            paragraphs_by_index[paragraph.source_index] = paragraph.text
            max_index = max(max_index, paragraph.source_index)

        ordered_paragraphs = []
        for index in range(max_index + 1):
            paragraph_text = replacements.get(index, paragraphs_by_index.get(index, ""))
            ordered_paragraphs.append(paragraph_text)

        output_text = "\n\n".join(text for text in ordered_paragraphs if text)
        output_path.write_text(output_text, encoding="utf-8")

    @staticmethod
    def _set_docx_paragraph_text(paragraph: ET.Element, new_text: str) -> None:
        paragraph_properties = paragraph.find(f"{{{WORD_NS}}}pPr")

        run_properties = None
        first_run = paragraph.find(f"{{{WORD_NS}}}r")
        if first_run is not None:
            original_run_properties = first_run.find(f"{{{WORD_NS}}}rPr")
            if original_run_properties is not None:
                run_properties = copy.deepcopy(original_run_properties)

        for child in list(paragraph):
            if child is not paragraph_properties:
                paragraph.remove(child)

        if not new_text:
            return

        run = ET.Element(f"{{{WORD_NS}}}r")
        if run_properties is not None:
            run.append(run_properties)

        text_node = ET.SubElement(run, f"{{{WORD_NS}}}t")
        if new_text != new_text.strip() or "  " in new_text:
            text_node.set(f"{{{XML_NS}}}space", "preserve")
        text_node.text = new_text
        paragraph.append(run)


def main() -> None:
    root = tk.Tk()
    style = ttk.Style()
    if "vista" in style.theme_names():
        style.theme_use("vista")
    TextCompareApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
