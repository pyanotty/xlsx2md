"""⑤ 生成: Block 列を「設計書として自然な Markdown」に描く。

表は内部結合が無ければ GFM、あれば HTML(colspan/rowspan)で描く。
HTML 化のロジックは converter.render_faithful_table と同じ方針をブロック単位に縮小したもの。
"""

from __future__ import annotations

from html import escape

from .extract import SheetModel
from .interpret import Block, Diagram, Figure, Heading, KeyValue, Paragraph, Table
from .segment import Region


def render_document(sheet: SheetModel, blocks: list[Block], image_dirname: str,
                    image_map: dict[tuple[int, int], list[str]] | None = None) -> str:
    image_map = image_map or {}
    parts: list[str] = [f"# {sheet.name}"]
    for b in blocks:
        parts.append(_render_block(sheet, b, image_dirname, image_map))
    return "\n\n".join(parts) + "\n"


def _render_block(sheet: SheetModel, b: Block, image_dirname: str,
                  image_map: dict[tuple[int, int], list[str]]) -> str:
    if isinstance(b, Heading):
        return f"{'#' * b.level} {b.text}"
    if isinstance(b, Paragraph):
        # セル内改行は段落内の改行(行末2スペース)として保つ
        return b.text.replace("\n", "  \n")
    if isinstance(b, KeyValue):
        return "\n".join(
            f"- **{_inline(k)}**: {_inline(v)}" if k else f"- {_inline(v)}"
            for k, v in b.pairs
        )
    if isinstance(b, Figure):
        return "\n\n".join(
            f"![{_stem(name)}]({image_dirname}/{name})" for name in b.images
        )
    if isinstance(b, Diagram):
        return f"```mermaid\n{b.mermaid}\n```"
    if isinstance(b, Table):
        return (_render_html_table(sheet, b.region, image_dirname, image_map)
                if b.has_merges
                else _render_gfm_table(sheet, b.region, image_dirname, image_map))
    return ""


def _inline(text: str) -> str:
    """インライン文脈用: 改行を除去。"""
    return text.replace("\n", " ").strip()


def _stem(name: str) -> str:
    return name.rsplit(".", 1)[0]


# --- GFM テーブル(結合なし) ---
def _render_gfm_table(sheet: SheetModel, reg: Region, image_dirname: str,
                      image_map: dict[tuple[int, int], list[str]]) -> str:
    # 全空の行・列(スペーサー)は落として自然な表にする
    rows_idx = [r for r in range(reg.r0, reg.r1 + 1)
                if not _row_empty(sheet, image_map, r, reg.c0, reg.c1)]
    cols_idx = [c for c in range(reg.c0, reg.c1 + 1)
                if not _col_empty(sheet, image_map, c, reg.r0, reg.r1)]
    if not rows_idx or not cols_idx:
        return ""

    rows: list[list[str]] = []
    for r in rows_idx:
        row = []
        for c in cols_idx:
            cell = _gfm_cell(sheet.text_at(r, c))
            for name in image_map.get((r, c), []):
                cell += f"<br>![]({image_dirname}/{name})"
            row.append(cell)
        rows.append(row)

    n = len(cols_idx)
    header = "| " + " | ".join(rows[0]) + " |"
    sep = "| " + " | ".join(["---"] * n) + " |"
    body = "\n".join("| " + " | ".join(r) + " |" for r in rows[1:])
    return "\n".join([header, sep] + ([body] if body else []))


def _row_empty(sheet, image_map, r, c0, c1) -> bool:
    return all(sheet.text_at(r, c) == "" and (r, c) not in image_map
               for c in range(c0, c1 + 1))


def _col_empty(sheet, image_map, c, r0, r1) -> bool:
    return all(sheet.text_at(r, c) == "" and (r, c) not in image_map
               for r in range(r0, r1 + 1))


def _gfm_cell(text: str) -> str:
    return text.replace("|", "\\|").replace("\n", "<br>")


# --- HTML テーブル(結合あり) ---
def _render_html_table(sheet: SheetModel, reg: Region, image_dirname: str,
                       image_map: dict[tuple[int, int], list[str]]) -> str:
    lines = ["<table>"]
    for r in range(reg.r0, reg.r1 + 1):
        cells: list[str] = []
        for c in range(reg.c0, reg.c1 + 1):
            if (r, c) in sheet.covered_to_origin:
                continue
            rs, cs = sheet.span_of.get((r, c), (1, 1))
            # 領域からはみ出す span はクランプ
            rs = min(rs, reg.r1 - r + 1)
            cs = min(cs, reg.c1 - c + 1)
            attrs = ""
            if rs > 1:
                attrs += f' rowspan="{rs}"'
            if cs > 1:
                attrs += f' colspan="{cs}"'
            content = escape(sheet.text_at(r, c)).replace("\n", "<br>")
            for name in image_map.get((r, c), []):
                img = f'<img src="{image_dirname}/{name}" alt="{_stem(name)}">'
                content = f"{content}<br>{img}" if content else img
            cells.append(f"<td{attrs}>{content}</td>")
        lines.append("<tr>" + "".join(cells) + "</tr>")
    lines.append("</table>")
    return "\n".join(lines)
