#!/usr/bin/env python3
"""把 DOCX 的批注/修订/高亮线性化为 Markdown 格式文本。"""
from __future__ import annotations

import argparse
import datetime as _dt
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set, Tuple
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
W = f"{{{W_NS}}}"
XML = f"{{{XML_NS}}}"


def _qn(local: str) -> str:
    return f"{W}{local}"


def _parse_iso_date(value: Optional[str]) -> Optional[str]:
    if not value:
        return None
    value = value.strip()
    if not value:
        return None
    try:
        dt = _dt.datetime.fromisoformat(value.replace("Z", "+00:00"))
        return dt.date().isoformat()
    except ValueError:
        return value


def _read_xml_from_docx(docx_path: Path, internal_path: str) -> Optional[ET.Element]:
    try:
        with zipfile.ZipFile(docx_path, "r") as zf:
            try:
                data = zf.read(internal_path)
            except KeyError:
                return None
    except zipfile.BadZipFile as exc:
        raise RuntimeError(f"Not a valid .docx (zip) file: {docx_path}") from exc

    try:
        return ET.fromstring(data)
    except ET.ParseError as exc:
        raise RuntimeError(f"Failed to parse XML: {internal_path} in {docx_path}") from exc


def _iter_plain_text_nodes(node: ET.Element) -> Iterable[str]:
    tag = node.tag
    if tag in (_qn("t"), _qn("delText")):
        if node.text:
            yield node.text
        return
    if tag == _qn("tab"):
        yield "\t"
        return
    if tag in (_qn("br"), _qn("cr")):
        yield "\n"
        return
    if tag == _qn("noBreakHyphen"):
        yield "\u2011"
        return
    if tag == _qn("softHyphen"):
        yield "\u00ad"
        return

    for child in list(node):
        yield from _iter_plain_text_nodes(child)


def _text_of(node: ET.Element) -> str:
    return "".join(_iter_plain_text_nodes(node))


# =============================================================================
# Comment 数据结构
# =============================================================================

@dataclass(frozen=True)
class Comment:
    comment_id: str
    author: Optional[str]
    date: Optional[str]
    text: str


def _parse_comments(comments_root: Optional[ET.Element]) -> Dict[str, Comment]:
    if comments_root is None:
        return {}

    comments: Dict[str, Comment] = {}
    for c in comments_root.findall(f".//{_qn('comment')}"):
        cid = c.get(_qn("id"))
        if not cid:
            continue
        author = c.get(_qn("author"))
        date = _parse_iso_date(c.get(_qn("date")))
        text_raw = _text_of(c).strip()
        text = " ".join(text_raw.split())
        comments[cid] = Comment(comment_id=cid, author=author, date=date, text=text)
    return comments


# =============================================================================
# 编号/列表解析 (numbering.xml)
# =============================================================================

@dataclass
class NumberingLevel:
    """一个编号定义中某一级的格式。"""
    num_fmt: str  # decimal, bullet, upperLetter, lowerLetter, upperRoman, lowerRoman, etc.
    lvl_text: str  # 例如 "%1.", "%1)", "•"
    start: int = 1


@dataclass
class NumberingDefinition:
    """一个抽象编号定义 (abstractNum)。"""
    levels: Dict[int, NumberingLevel] = field(default_factory=dict)


@dataclass
class NumberingInfo:
    """整个文档的编号信息。"""
    # abstractNumId -> NumberingDefinition
    abstract_nums: Dict[str, NumberingDefinition] = field(default_factory=dict)
    # numId -> abstractNumId
    num_to_abstract: Dict[str, str] = field(default_factory=dict)

    def get_level(self, num_id: str, ilvl: int) -> Optional[NumberingLevel]:
        abstract_id = self.num_to_abstract.get(num_id)
        if not abstract_id:
            return None
        abstract = self.abstract_nums.get(abstract_id)
        if not abstract:
            return None
        return abstract.levels.get(ilvl)


def _parse_numbering(numbering_root: Optional[ET.Element]) -> NumberingInfo:
    info = NumberingInfo()
    if numbering_root is None:
        return info

    # 解析 abstractNum
    for abstract in numbering_root.findall(f".//{_qn('abstractNum')}"):
        abstract_id = abstract.get(_qn("abstractNumId"))
        if not abstract_id:
            continue
        defn = NumberingDefinition()
        for lvl in abstract.findall(f"./{_qn('lvl')}"):
            ilvl_str = lvl.get(_qn("ilvl"))
            if ilvl_str is None:
                continue
            ilvl = int(ilvl_str)

            num_fmt_el = lvl.find(f"./{_qn('numFmt')}")
            num_fmt = num_fmt_el.get(_qn("val")) if num_fmt_el is not None else "decimal"

            lvl_text_el = lvl.find(f"./{_qn('lvlText')}")
            lvl_text = lvl_text_el.get(_qn("val")) if lvl_text_el is not None else ""

            start_el = lvl.find(f"./{_qn('start')}")
            start = int(start_el.get(_qn("val"))) if start_el is not None else 1

            defn.levels[ilvl] = NumberingLevel(num_fmt=num_fmt, lvl_text=lvl_text, start=start)
        info.abstract_nums[abstract_id] = defn

    # 解析 num -> abstractNumId 映射
    for num in numbering_root.findall(f".//{_qn('num')}"):
        num_id = num.get(_qn("numId"))
        if not num_id:
            continue
        abstract_ref = num.find(f"./{_qn('abstractNumId')}")
        if abstract_ref is not None:
            info.num_to_abstract[num_id] = abstract_ref.get(_qn("val")) or ""

    return info


# =============================================================================
# 样式解析 (styles.xml)
# =============================================================================

@dataclass
class StyleInfo:
    """段落样式信息。"""
    # styleId -> heading level (1-6) 或 None
    heading_levels: Dict[str, int] = field(default_factory=dict)


def _parse_styles(styles_root: Optional[ET.Element]) -> StyleInfo:
    info = StyleInfo()
    if styles_root is None:
        return info

    for style in styles_root.findall(f".//{_qn('style')}"):
        style_id = style.get(_qn("styleId"))
        if not style_id:
            continue

        # 检查 outlineLvl（大纲级别）
        ppr = style.find(f"./{_qn('pPr')}")
        if ppr is not None:
            outline = ppr.find(f"./{_qn('outlineLvl')}")
            if outline is not None:
                val = outline.get(_qn("val"))
                if val is not None:
                    lvl = int(val)
                    if 0 <= lvl <= 5:  # Word 大纲级别 0-5 对应 H1-H6
                        info.heading_levels[style_id] = lvl + 1

        # 也检查样式名称，常见的 Heading 样式
        name_el = style.find(f"./{_qn('name')}")
        if name_el is not None:
            name = name_el.get(_qn("val")) or ""
            name_lower = name.lower()
            # "Heading 1", "heading 2", "标题 1" 等
            match = re.match(r"(?:heading|标题)\s*(\d+)", name_lower)
            if match:
                lvl = int(match.group(1))
                if 1 <= lvl <= 6:
                    info.heading_levels[style_id] = lvl

    return info


# =============================================================================
# 渲染状态与输出
# =============================================================================

@dataclass
class CommentRange:
    """跟踪一个批注范围收集的原文。"""
    comment_id: str
    text_parts: List[str] = field(default_factory=list)


@dataclass
class CollectedComment:
    """收集完成的批注，包含原文和批注内容。"""
    comment_id: str
    author: Optional[str]
    date: Optional[str]
    original_text: str
    comment_text: str


@dataclass
class RenderState:
    comments: Dict[str, Comment]
    numbering: NumberingInfo
    styles: StyleInfo
    # 正在收集原文的批注范围（可能嵌套）
    active_comment_ranges: Dict[str, CommentRange] = field(default_factory=dict)
    # 已完成收集的批注（当前段落内）
    collected_comments: List[CollectedComment] = field(default_factory=list)
    # 已经输出过的批注 ID
    emitted_comment_ids: Set[str] = field(default_factory=set)
    # 列表编号计数器：(numId, ilvl) -> current_count
    list_counters: Dict[Tuple[str, int], int] = field(default_factory=dict)

    def start_comment_range(self, cid: str) -> None:
        """开始收集批注范围的原文。"""
        if cid not in self.active_comment_ranges:
            self.active_comment_ranges[cid] = CommentRange(comment_id=cid)

    def append_to_active_ranges(self, text: str) -> None:
        """向所有活跃的批注范围追加文本。"""
        for cr in self.active_comment_ranges.values():
            cr.text_parts.append(text)

    def end_comment_range(self, cid: str) -> None:
        """结束批注范围，收集完成。"""
        if cid in self.active_comment_ranges:
            cr = self.active_comment_ranges.pop(cid)
            original = "".join(cr.text_parts).strip()
            original = " ".join(original.split())  # 合并空白

            comment = self.comments.get(cid)
            if comment and cid not in self.emitted_comment_ids:
                self.collected_comments.append(CollectedComment(
                    comment_id=cid,
                    author=comment.author,
                    date=comment.date,
                    original_text=original,
                    comment_text=comment.text,
                ))
                self.emitted_comment_ids.add(cid)

    def handle_comment_reference(self, cid: str) -> None:
        """处理 commentReference（批注引用点）。"""
        # 有时批注只有 reference 没有 range，这里兜底
        if cid in self.emitted_comment_ids:
            return
        if cid in self.active_comment_ranges:
            # 范围还没结束，等 RangeEnd
            return
        # 没有范围，直接收集
        comment = self.comments.get(cid)
        if comment:
            self.collected_comments.append(CollectedComment(
                comment_id=cid,
                author=comment.author,
                date=comment.date,
                original_text="",
                comment_text=comment.text,
            ))
            self.emitted_comment_ids.add(cid)

    def pop_collected_comments(self) -> List[CollectedComment]:
        """取出并清空当前收集的批注。"""
        result = self.collected_comments
        self.collected_comments = []
        return result

    def get_list_marker(self, num_id: str, ilvl: int) -> str:
        """获取列表标记。"""
        level = self.numbering.get_level(num_id, ilvl)
        if not level:
            return "- "

        indent = "  " * ilvl

        if level.num_fmt == "bullet":
            return f"{indent}- "
        elif level.num_fmt in ("decimal", "upperLetter", "lowerLetter", "upperRoman", "lowerRoman"):
            key = (num_id, ilvl)
            count = self.list_counters.get(key, level.start)
            self.list_counters[key] = count + 1
            return f"{indent}{count}. "
        else:
            return f"{indent}- "

    def reset_list_counter(self, num_id: str, ilvl: int) -> None:
        """重置列表计数器。"""
        key = (num_id, ilvl)
        level = self.numbering.get_level(num_id, ilvl)
        if level:
            self.list_counters[key] = level.start

    def get_heading_level(self, style_id: Optional[str]) -> Optional[int]:
        """获取标题级别。"""
        if not style_id:
            return None
        return self.styles.heading_levels.get(style_id)


@dataclass(frozen=True)
class RunContext:
    """Run 级别的上下文。"""
    highlight: bool = False
    bold: bool = False
    italic: bool = False


@dataclass
class RenderOut:
    """渲染输出缓冲。"""
    parts: List[str] = field(default_factory=list)
    _highlight_open: bool = False
    _bold_open: bool = False
    _italic_open: bool = False

    def _close_all_formats(self) -> None:
        """关闭所有打开的格式标记。"""
        if self._italic_open:
            self.parts.append("*")
            self._italic_open = False
        if self._bold_open:
            self.parts.append("**")
            self._bold_open = False
        if self._highlight_open:
            self.parts.append("==")
            self._highlight_open = False

    def append_text(self, text: str, ctx: RunContext, state: RenderState) -> None:
        """添加文本，处理格式标记。"""
        if not text:
            return

        # 向活跃的批注范围追加原文（不带格式标记）
        state.append_to_active_ranges(text)

        # 处理高亮
        if ctx.highlight and not self._highlight_open:
            self.parts.append("==")
            self._highlight_open = True
        elif not ctx.highlight and self._highlight_open:
            self.parts.append("==")
            self._highlight_open = False

        # 处理粗体
        if ctx.bold and not self._bold_open:
            self.parts.append("**")
            self._bold_open = True
        elif not ctx.bold and self._bold_open:
            self.parts.append("**")
            self._bold_open = False

        # 处理斜体
        if ctx.italic and not self._italic_open:
            self.parts.append("*")
            self._italic_open = True
        elif not ctx.italic and self._italic_open:
            self.parts.append("*")
            self._italic_open = False

        self.parts.append(text)

    def append_literal(self, literal: str, state: Optional[RenderState] = None) -> None:
        """添加字面量（如修订标记），先关闭格式。"""
        self._close_all_formats()
        if literal:
            self.parts.append(literal)
            # 字面量也追加到批注范围
            if state:
                state.append_to_active_ranges(literal)

    def finish(self) -> str:
        """完成渲染，关闭所有格式。"""
        self._close_all_formats()
        return "".join(self.parts)


# =============================================================================
# Run 属性解析
# =============================================================================

def _get_run_context(run: ET.Element) -> RunContext:
    """从 run 的 rPr 解析格式属性。"""
    rpr = run.find(f"./{_qn('rPr')}")
    if rpr is None:
        return RunContext()

    # 高亮
    highlight = False
    hl = rpr.find(f"./{_qn('highlight')}")
    if hl is not None:
        val = (hl.get(_qn("val")) or "").strip().lower()
        highlight = val not in ("", "none")

    # 粗体
    bold = False
    b = rpr.find(f"./{_qn('b')}")
    if b is not None:
        val = b.get(_qn("val"))
        bold = val is None or val.lower() not in ("false", "0")

    # 斜体
    italic = False
    i = rpr.find(f"./{_qn('i')}")
    if i is not None:
        val = i.get(_qn("val"))
        italic = val is None or val.lower() not in ("false", "0")

    return RunContext(highlight=highlight, bold=bold, italic=italic)


# =============================================================================
# 段落属性解析
# =============================================================================

@dataclass
class ParagraphProps:
    """段落属性。"""
    style_id: Optional[str] = None
    num_id: Optional[str] = None
    ilvl: int = 0


def _get_paragraph_props(p: ET.Element) -> ParagraphProps:
    """解析段落属性。"""
    props = ParagraphProps()
    ppr = p.find(f"./{_qn('pPr')}")
    if ppr is None:
        return props

    # 样式
    pstyle = ppr.find(f"./{_qn('pStyle')}")
    if pstyle is not None:
        props.style_id = pstyle.get(_qn("val"))

    # 编号/列表
    numpr = ppr.find(f"./{_qn('numPr')}")
    if numpr is not None:
        ilvl_el = numpr.find(f"./{_qn('ilvl')}")
        if ilvl_el is not None:
            props.ilvl = int(ilvl_el.get(_qn("val")) or "0")
        numid_el = numpr.find(f"./{_qn('numId')}")
        if numid_el is not None:
            props.num_id = numid_el.get(_qn("val"))

    return props


# =============================================================================
# 渲染逻辑
# =============================================================================

def _render_node(node: ET.Element, ctx: RunContext, out: RenderOut, state: RenderState) -> None:
    """递归渲染节点。"""
    tag = node.tag

    # 批注范围开始
    if tag == _qn("commentRangeStart"):
        cid = node.get(_qn("id"))
        if cid:
            state.start_comment_range(cid)
        return

    # 批注范围结束
    if tag == _qn("commentRangeEnd"):
        cid = node.get(_qn("id"))
        if cid:
            state.end_comment_range(cid)
        return

    # 插入修订
    if tag == _qn("ins"):
        inner = RenderOut()
        for child in list(node):
            _render_node(child, ctx, inner, state)
        txt = inner.finish()
        if txt:
            out.append_literal("{+" + txt + "+}", state)
        return

    # 删除修订
    if tag == _qn("del"):
        inner = RenderOut()
        for child in list(node):
            _render_node(child, ctx, inner, state)
        txt = inner.finish()
        if txt:
            out.append_literal("[-" + txt + "-]", state)
        return

    # Run
    if tag == _qn("r"):
        run_ctx = _get_run_context(node)
        # 合并上下文
        merged_ctx = RunContext(
            highlight=ctx.highlight or run_ctx.highlight,
            bold=ctx.bold or run_ctx.bold,
            italic=ctx.italic or run_ctx.italic,
        )
        for child in list(node):
            ctag = child.tag
            if ctag == _qn("rPr"):
                continue
            if ctag == _qn("commentReference"):
                cid = child.get(_qn("id"))
                if cid:
                    state.handle_comment_reference(cid)
                continue

            if ctag in (_qn("t"), _qn("delText")):
                out.append_text(child.text or "", merged_ctx, state)
                continue
            if ctag == _qn("tab"):
                out.append_literal("\t", state)
                continue
            if ctag in (_qn("br"), _qn("cr")):
                out.append_literal("\n", state)
                continue
            if ctag == _qn("noBreakHyphen"):
                out.append_text("\u2011", merged_ctx, state)
                continue
            if ctag == _qn("softHyphen"):
                out.append_text("\u00ad", merged_ctx, state)
                continue

            _render_node(child, merged_ctx, out, state)
        return

    # 段落（在顶层调用时跳过 pPr）
    if tag == _qn("p"):
        for child in list(node):
            if child.tag == _qn("pPr"):
                continue
            _render_node(child, ctx, out, state)
        return

    # 表格
    if tag == _qn("tbl"):
        _render_table(node, ctx, out, state)
        return

    # 其它：递归
    for child in list(node):
        _render_node(child, ctx, out, state)


def _render_paragraph_content(p: ET.Element, state: RenderState) -> str:
    """渲染段落内容（不含前缀）。"""
    out = RenderOut()
    _render_node(p, RunContext(), out, state)
    return out.finish().rstrip()


def _format_comment_block(c: CollectedComment) -> str:
    """格式化批注为 Markdown 引用块。"""
    lines = []

    # 元信息行
    meta_parts = [f"**[批注 #{c.comment_id}]**"]
    if c.author:
        meta_parts.append(c.author)
    if c.date:
        meta_parts.append(f"({c.date})")
    lines.append("> " + " ".join(meta_parts))

    # 原文
    if c.original_text:
        lines.append(f"> **原文**：{c.original_text}")

    # 批注内容
    if c.comment_text:
        lines.append(f"> **批注**：{c.comment_text}")

    return "\n".join(lines)


def _render_paragraph(p: ET.Element, state: RenderState) -> List[str]:
    """渲染段落，返回输出行列表。"""
    props = _get_paragraph_props(p)
    content = _render_paragraph_content(p, state)

    # 收集该段落的批注
    comments = state.pop_collected_comments()

    lines: List[str] = []

    # 空段落跳过
    if not content and not comments:
        return lines

    # 标题
    heading_level = state.get_heading_level(props.style_id)
    if heading_level and content:
        prefix = "#" * heading_level + " "
        lines.append(prefix + content)
    # 列表
    elif props.num_id and content:
        marker = state.get_list_marker(props.num_id, props.ilvl)
        lines.append(marker + content)
    # 普通段落
    elif content:
        lines.append(content)

    # 附加批注块
    for c in comments:
        lines.append("")  # 空行分隔
        lines.append(_format_comment_block(c))

    return lines


def _render_table(tbl: ET.Element, _ctx: RunContext, out: RenderOut, state: RenderState) -> None:
    """渲染表格为简单的 Markdown 表格或分隔行。"""
    rows: List[List[str]] = []

    for tr in tbl.findall(f"./{_qn('tr')}"):
        row_cells: List[str] = []
        for tc in tr.findall(f"./{_qn('tc')}"):
            cell_parts: List[str] = []
            for p in tc.findall(f".//{_qn('p')}"):
                content = _render_paragraph_content(p, state)
                if content:
                    cell_parts.append(content)
            row_cells.append(" ".join(cell_parts).strip())
        rows.append(row_cells)

    if not rows:
        return

    # 确定列数
    max_cols = max(len(r) for r in rows) if rows else 0
    if max_cols == 0:
        return

    # 输出 Markdown 表格
    for i, row in enumerate(rows):
        # 补齐列数
        while len(row) < max_cols:
            row.append("")
        line = "| " + " | ".join(row) + " |"
        out.append_literal(line + "\n", state)
        # 第一行后加分隔行
        if i == 0:
            sep = "| " + " | ".join(["---"] * max_cols) + " |"
            out.append_literal(sep + "\n", state)


def linearize_docx(docx_path: Path) -> str:
    """将 DOCX 文档线性化为 Markdown 文本。"""
    document_root = _read_xml_from_docx(docx_path, "word/document.xml")
    if document_root is None:
        raise RuntimeError(f"Missing word/document.xml in {docx_path}")

    comments_root = _read_xml_from_docx(docx_path, "word/comments.xml")
    comments = _parse_comments(comments_root)

    numbering_root = _read_xml_from_docx(docx_path, "word/numbering.xml")
    numbering = _parse_numbering(numbering_root)

    styles_root = _read_xml_from_docx(docx_path, "word/styles.xml")
    styles = _parse_styles(styles_root)

    state = RenderState(comments=comments, numbering=numbering, styles=styles)

    body = document_root.find(f".//{_qn('body')}")
    if body is None:
        return ""

    output_lines: List[str] = []

    for child in list(body):
        if child.tag == _qn("p"):
            para_lines = _render_paragraph(child, state)
            output_lines.extend(para_lines)
        elif child.tag == _qn("tbl"):
            out = RenderOut()
            _render_table(child, RunContext(), out, state)
            chunk = out.finish().rstrip()
            if chunk:
                output_lines.append(chunk)
        # sectPr 等跳过

    # 处理跨段落的未闭合批注
    remaining = state.pop_collected_comments()
    for c in remaining:
        output_lines.append("")
        output_lines.append(_format_comment_block(c))

    text = "\n".join(output_lines).rstrip()
    return (text + "\n") if text else ""


def _build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="把 DOCX 的批注/修订/高亮线性化为 Markdown 格式文本。",
    )
    p.add_argument("input", type=str, help="输入 .docx 路径")
    p.add_argument(
        "output",
        type=str,
        nargs="?",
        default="-",
        help='输出路径；默认 "-" 表示 stdout',
    )
    return p


def main(argv: Optional[List[str]] = None) -> int:
    args = _build_arg_parser().parse_args(argv)
    input_path = Path(args.input)
    if not input_path.exists():
        raise SystemExit(f"Input not found: {input_path}")

    text = linearize_docx(input_path)

    if args.output == "-" or args.output.strip() == "":
        print(text, end="")
        return 0

    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(text, encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
