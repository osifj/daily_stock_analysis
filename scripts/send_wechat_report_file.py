# -*- coding: utf-8 -*-
"""Convert generated Markdown reports to Word and send them to WeChat Work."""
from __future__ import annotations

import argparse
import logging
import re
import sys
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Iterable, List, Optional, Sequence

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from docx import Document
from docx.document import Document as DocumentObject
from docx.enum.text import WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

from src.config import get_config
from src.notification_sender.wechat_sender import WechatSender


logger = logging.getLogger(__name__)
CN_TZ = timezone(timedelta(hours=8))


def _set_cell_shading(cell, fill: str = "D9EAF7") -> None:
    tc_pr = cell._tc.get_or_add_tcPr()  # noqa: SLF001 - python-docx table formatting hook.
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def _set_cell_border(cell) -> None:
    tc = cell._tc  # noqa: SLF001 - python-docx table formatting hook.
    tc_pr = tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = f"w:{edge}"
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "4")
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), "B7C0CC")


def _configure_document(document: DocumentObject) -> None:
    section = document.sections[0]
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    normal_style = document.styles["Normal"]
    normal_style.font.name = "Microsoft YaHei"
    normal_style._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")  # noqa: SLF001
    normal_style.font.size = Pt(10.5)

    for style_name, size in (("Heading 1", 18), ("Heading 2", 15), ("Heading 3", 13)):
        style = document.styles[style_name]
        style.font.name = "Microsoft YaHei"
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")  # noqa: SLF001
        style.font.size = Pt(size)
        style.font.bold = True


def _add_inline_markdown(paragraph, text: str) -> None:
    """Render a small subset of inline Markdown: bold and inline code."""
    pattern = re.compile(r"(\*\*[^*]+\*\*|`[^`]+`)")
    position = 0
    for match in pattern.finditer(text):
        if match.start() > position:
            paragraph.add_run(text[position:match.start()])
        token = match.group(0)
        if token.startswith("**"):
            run = paragraph.add_run(token[2:-2])
            run.bold = True
        else:
            run = paragraph.add_run(token[1:-1])
            run.font.name = "Consolas"
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "Consolas")  # noqa: SLF001
        position = match.end()
    if position < len(text):
        paragraph.add_run(text[position:])


def _is_table_separator(line: str) -> bool:
    cells = [cell.strip() for cell in line.strip().strip("|").split("|")]
    return bool(cells) and all(re.fullmatch(r":?-{3,}:?", cell or "") for cell in cells)


def _parse_table(lines: Sequence[str], start: int) -> tuple[Optional[List[List[str]]], int]:
    if start + 1 >= len(lines):
        return None, start
    if "|" not in lines[start] or not _is_table_separator(lines[start + 1]):
        return None, start

    rows: List[List[str]] = []
    index = start
    while index < len(lines) and "|" in lines[index] and lines[index].strip():
        if not _is_table_separator(lines[index]):
            rows.append([cell.strip() for cell in lines[index].strip().strip("|").split("|")])
        index += 1
    return rows, index


def _add_table(document: DocumentObject, rows: Sequence[Sequence[str]]) -> None:
    if not rows:
        return
    column_count = max(len(row) for row in rows)
    table = document.add_table(rows=len(rows), cols=column_count)
    table.style = "Table Grid"
    for row_index, row in enumerate(rows):
        for col_index in range(column_count):
            cell = table.cell(row_index, col_index)
            cell.text = row[col_index] if col_index < len(row) else ""
            _set_cell_border(cell)
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(9.5)
            if row_index == 0:
                _set_cell_shading(cell)
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.bold = True


def markdown_to_docx(markdown: str, output_path: Path, *, title: Optional[str] = None) -> Path:
    document = Document()
    _configure_document(document)

    report_title = title or "股票智能分析报告"
    document.add_heading(report_title, level=0)
    document.add_paragraph(f"生成时间：{datetime.now(CN_TZ).strftime('%Y-%m-%d %H:%M:%S')}（北京时间）")

    lines = markdown.splitlines()
    index = 0
    in_code_block = False
    code_lines: List[str] = []

    while index < len(lines):
        raw_line = lines[index].rstrip()
        line = raw_line.strip()

        if line.startswith("```"):
            if in_code_block:
                paragraph = document.add_paragraph(style="No Spacing")
                run = paragraph.add_run("\n".join(code_lines))
                run.font.name = "Consolas"
                run._element.rPr.rFonts.set(qn("w:eastAsia"), "Consolas")  # noqa: SLF001
                run.font.size = Pt(9)
                code_lines = []
                in_code_block = False
            else:
                in_code_block = True
            index += 1
            continue

        if in_code_block:
            code_lines.append(raw_line)
            index += 1
            continue

        table_rows, next_index = _parse_table(lines, index)
        if table_rows is not None:
            _add_table(document, table_rows)
            index = next_index
            continue

        if not line:
            index += 1
            continue
        if re.fullmatch(r"[-*_]{3,}", line):
            paragraph = document.add_paragraph()
            paragraph.add_run().add_break(WD_BREAK.LINE)
            index += 1
            continue

        heading_match = re.match(r"^(#{1,6})\s+(.+)$", line)
        if heading_match:
            level = min(len(heading_match.group(1)), 3)
            document.add_heading(heading_match.group(2).strip(), level=level)
            index += 1
            continue

        bullet_match = re.match(r"^[-*+]\s+(.+)$", line)
        if bullet_match:
            paragraph = document.add_paragraph(style="List Bullet")
            _add_inline_markdown(paragraph, bullet_match.group(1).strip())
            index += 1
            continue

        numbered_match = re.match(r"^\d+[.)]\s+(.+)$", line)
        if numbered_match:
            paragraph = document.add_paragraph(style="List Number")
            _add_inline_markdown(paragraph, numbered_match.group(1).strip())
            index += 1
            continue

        paragraph = document.add_paragraph()
        _add_inline_markdown(paragraph, line)
        index += 1

    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)
    return output_path


def find_report_files(reports_dir: Path) -> List[Path]:
    patterns = ("report_*.md", "market_review_*.md")
    files: List[Path] = []
    for pattern in patterns:
        files.extend(reports_dir.glob(pattern))
    return sorted(files, key=lambda path: path.stat().st_mtime)


def build_combined_markdown(report_files: Iterable[Path]) -> str:
    parts: List[str] = []
    for report_file in report_files:
        title = "大盘复盘" if report_file.name.startswith("market_review_") else "个股分析"
        parts.append(f"# {title}\n\n" + report_file.read_text(encoding="utf-8"))
    return "\n\n---\n\n".join(parts)


def send_wechat_report_file(reports_dir: Path, output_path: Optional[Path] = None) -> Path:
    report_files = find_report_files(reports_dir)
    if not report_files:
        raise FileNotFoundError(f"未找到可转换的 Markdown 报告: {reports_dir}")

    output = output_path or reports_dir / f"stock_analysis_report_{datetime.now(CN_TZ).strftime('%Y%m%d_%H%M%S')}.docx"
    markdown = build_combined_markdown(report_files)
    markdown_to_docx(markdown, output, title="股票智能分析报告")

    sender = WechatSender(get_config())
    summary = (
        "股票智能分析报告已生成，完整内容请查看 Word 附件。\n"
        f"包含报告：{', '.join(path.name for path in report_files)}"
    )
    if not sender.send_to_wechat(summary):
        raise RuntimeError("企业微信报告摘要发送失败")
    if not sender.send_file_to_wechat(output):
        raise RuntimeError("企业微信 Word 报告文件发送失败")
    return output


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Send generated reports as a WeChat Work Word file")
    parser.add_argument("--reports-dir", default="reports", help="Report directory containing Markdown reports")
    parser.add_argument("--output", default="", help="Optional .docx output path")
    return parser.parse_args()


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    args = parse_args()
    output = Path(args.output) if args.output else None
    try:
        docx_path = send_wechat_report_file(Path(args.reports_dir), output)
        logger.info("企业微信 Word 报告已发送: %s", docx_path)
        return 0
    except Exception as exc:  # noqa: BLE001 - CLI should report a clean failure.
        logger.exception("企业微信 Word 报告发送失败: %s", exc)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
