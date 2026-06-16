"""① 抽出層: Worksheet を解釈しやすい SheetModel に変換する。

v1(converter.py)の値整形・アンカー検出を再利用しつつ、
② 以降の段が必要とする「結合・書式・画像セル」を構造化して保持する。
出力(Markdown)には出さない書式も、③ の役割判定の手がかりとして読む。
"""

from __future__ import annotations

from dataclasses import dataclass, field

from openpyxl.utils import get_column_letter

from .converter import _anchor_cell, _format_value


@dataclass
class CellInfo:
    """セル(結合の左上=origin)の書式サマリ。③ の役割判定に使う。"""
    bold: bool = False
    font_size: float | None = None
    has_fill: bool = False
    has_border: bool = False
    h_align: str | None = None


@dataclass
class SheetModel:
    """1 シートの構造化表現。"""
    name: str
    min_row: int
    min_col: int
    max_row: int
    max_col: int
    # origin(結合左上 or 通常セル)の非空テキストのみ保持
    origin_text: dict[tuple[int, int], str] = field(default_factory=dict)
    span_of: dict[tuple[int, int], tuple[int, int]] = field(default_factory=dict)
    covered_to_origin: dict[tuple[int, int], tuple[int, int]] = field(default_factory=dict)
    style: dict[tuple[int, int], CellInfo] = field(default_factory=dict)
    image_cells: dict[tuple[int, int], int] = field(default_factory=dict)

    # --- 問い合わせ用ヘルパ ---
    def origin_of(self, r: int, c: int) -> tuple[int, int]:
        return self.covered_to_origin.get((r, c), (r, c))

    def text_at(self, r: int, c: int) -> str:
        return self.origin_text.get(self.origin_of(r, c), "")

    def has_image(self, r: int, c: int) -> bool:
        return self.image_cells.get((r, c), 0) > 0

    def is_blank(self, r: int, c: int) -> bool:
        return self.text_at(r, c) == "" and not self.has_image(r, c)

    def addr(self, r: int, c: int) -> str:
        return f"{get_column_letter(c)}{r}"


def extract_sheet(ws, wsf, options) -> SheetModel | None:
    """Worksheet から SheetModel を構築する。内容が無ければ None。"""
    # 結合情報
    span_of: dict[tuple[int, int], tuple[int, int]] = {}
    covered_to_origin: dict[tuple[int, int], tuple[int, int]] = {}
    for mr in ws.merged_cells.ranges:
        top = (mr.min_row, mr.min_col)
        span_of[top] = (mr.max_row - mr.min_row + 1, mr.max_col - mr.min_col + 1)
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                if (r, c) != top:
                    covered_to_origin[(r, c)] = top

    origin_text: dict[tuple[int, int], str] = {}
    style: dict[tuple[int, int], CellInfo] = {}

    for row in ws.iter_rows():
        for cell in row:
            pos = (cell.row, cell.column)
            if pos in covered_to_origin:
                continue  # 被覆セルは origin 側で扱う
            raw = cell.value
            if raw is None and options.formula_fallback and wsf is not None:
                vf = wsf.cell(row=cell.row, column=cell.column).value
                if isinstance(vf, str) and vf.startswith("="):
                    raw = vf
            text = _format_value(raw)
            if text.strip() != "":
                origin_text[pos] = text
                style[pos] = _cell_style(cell)

    # 画像セル(ファイルは書かず、アンカー位置のみ収集)
    image_cells: dict[tuple[int, int], int] = {}
    for img in (getattr(ws, "_images", None) or []):
        anchor = _anchor_cell(img)
        if anchor is not None:
            image_cells[anchor] = image_cells.get(anchor, 0) + 1

    # 外接矩形(外周空白をトリム)
    bbox = _bbox(origin_text, span_of, image_cells)
    if bbox is None:
        return None
    min_row, min_col, max_row, max_col = bbox

    return SheetModel(
        name=ws.title,
        min_row=min_row, min_col=min_col, max_row=max_row, max_col=max_col,
        origin_text=origin_text, span_of=span_of,
        covered_to_origin=covered_to_origin, style=style, image_cells=image_cells,
    )


def _cell_style(cell) -> CellInfo:
    font = cell.font
    fill = cell.fill
    has_fill = bool(getattr(fill, "patternType", None)) and fill.patternType != "none"
    border = cell.border
    has_border = any(
        getattr(getattr(border, side), "style", None)
        for side in ("left", "right", "top", "bottom")
    )
    align = getattr(cell.alignment, "horizontal", None)
    return CellInfo(
        bold=bool(getattr(font, "bold", False)),
        font_size=getattr(font, "sz", None),
        has_fill=has_fill,
        has_border=has_border,
        h_align=align,
    )


def _bbox(origin_text, span_of, image_cells):
    min_r = min_c = max_r = max_c = None

    def expand(r1, c1, r2, c2):
        nonlocal min_r, min_c, max_r, max_c
        min_r = r1 if min_r is None else min(min_r, r1)
        min_c = c1 if min_c is None else min(min_c, c1)
        max_r = r2 if max_r is None else max(max_r, r2)
        max_c = c2 if max_c is None else max(max_c, c2)

    for (r, c) in origin_text:
        rs, cs = span_of.get((r, c), (1, 1))
        expand(r, c, r + rs - 1, c + cs - 1)
    for (r, c) in image_cells:
        expand(r, c, r, c)

    if min_r is None:
        return None
    return (min_r, min_c, max_r, max_c)
