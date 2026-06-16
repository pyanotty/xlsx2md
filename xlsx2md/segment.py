"""② ブロック分割(再帰 XY-cut)＋ ④ 読み順。

シートを意味要素ごとの矩形(Region)に分ける。手がかり:
  1. 完全な空行で横バンドに分割(上→下)
  2. 全幅の結合バナー行を境界に分割(見出し ↔ 本文の隣接を切る)
  3. 完全な空列で縦に分割(左→右)
再帰木の走査順がそのまま読み順になる。

役割の確定は ③(将来の interpret)で行う。ここでは可視化用の
暫定ロール guess_role() のみ提供する。
"""

from __future__ import annotations

from dataclasses import dataclass

from .extract import SheetModel


@dataclass
class Region:
    r0: int
    c0: int
    r1: int
    c1: int
    depth: int = 0

    @property
    def n_rows(self) -> int:
        return self.r1 - self.r0 + 1

    @property
    def n_cols(self) -> int:
        return self.c1 - self.c0 + 1


def segment(sheet: SheetModel) -> list[Region]:
    """シートを Region 列(読み順)に分割する。"""
    root = Region(sheet.min_row, sheet.min_col, sheet.max_row, sheet.max_col)
    leaves: list[Region] = []
    _cut(sheet, root, leaves, depth=0)
    return leaves


def _cut(sheet: SheetModel, reg: Region, out: list[Region], depth: int) -> None:
    # 1) 空行で横分割
    bands = _split_rows_by_blank(sheet, reg)
    if len(bands) > 1:
        for b in bands:
            _cut(sheet, b, out, depth + 1)
        return
    # 2) 全幅バナー行で横分割
    fw = _split_full_width(sheet, reg)
    if len(fw) > 1:
        for b in fw:
            _cut(sheet, b, out, depth + 1)
        return
    # 3) 空列で縦分割
    cols = _split_cols_by_blank(sheet, reg)
    if len(cols) > 1:
        for b in cols:
            _cut(sheet, b, out, depth + 1)
        return
    # 葉
    reg.depth = depth
    out.append(reg)


# --------------------------------------------------------------------------- #
# 分割ヘルパ
# --------------------------------------------------------------------------- #
def _row_blank(sheet: SheetModel, r: int, c0: int, c1: int) -> bool:
    return all(sheet.is_blank(r, c) for c in range(c0, c1 + 1))


def _col_blank(sheet: SheetModel, c: int, r0: int, r1: int) -> bool:
    return all(sheet.is_blank(r, c) for r in range(r0, r1 + 1))


def _split_rows_by_blank(sheet: SheetModel, reg: Region) -> list[Region]:
    """非空行の連続群をバンドにまとめ、空行を区切りとして落とす。"""
    groups: list[tuple[int, int]] = []
    start = None
    for r in range(reg.r0, reg.r1 + 1):
        if _row_blank(sheet, r, reg.c0, reg.c1):
            if start is not None:
                groups.append((start, r - 1))
                start = None
        else:
            if start is None:
                start = r
    if start is not None:
        groups.append((start, reg.r1))

    if len(groups) <= 1:
        # 区切りなし(先頭/末尾の空白だけトリム)
        if groups and (groups[0] != (reg.r0, reg.r1)):
            g = groups[0]
            return [Region(g[0], reg.c0, g[1], reg.c1)]
        return [reg]
    return [Region(s, reg.c0, e, reg.c1) for (s, e) in groups]


def _split_cols_by_blank(sheet: SheetModel, reg: Region) -> list[Region]:
    groups: list[tuple[int, int]] = []
    start = None
    for c in range(reg.c0, reg.c1 + 1):
        if _col_blank(sheet, c, reg.r0, reg.r1):
            if start is not None:
                groups.append((start, c - 1))
                start = None
        else:
            if start is None:
                start = c
    if start is not None:
        groups.append((start, reg.c1))

    if len(groups) <= 1:
        if groups and (groups[0] != (reg.c0, reg.c1)):
            g = groups[0]
            return [Region(reg.r0, g[0], reg.r1, g[1])]
        return [reg]
    return [Region(reg.r0, s, reg.r1, e) for (s, e) in groups]


def _is_full_width_banner(sheet: SheetModel, r: int, c0: int, c1: int) -> bool:
    """行 r が「領域幅いっぱいに広がる単一結合セル」のみで構成されるか。"""
    origin = sheet.origin_of(r, c0)
    if origin not in sheet.origin_text:
        return False
    rs, cs = sheet.span_of.get(origin, (1, 1))
    # origin が行頭から領域全幅を覆い、その行に他の内容が無い
    if origin[1] != c0 or cs < (c1 - c0 + 1):
        return False
    return True


def _split_full_width(sheet: SheetModel, reg: Region) -> list[Region]:
    """全幅バナー行を独立領域として切り出す。"""
    banner_rows = [
        r for r in range(reg.r0, reg.r1 + 1)
        if _is_full_width_banner(sheet, r, reg.c0, reg.c1)
    ]
    if not banner_rows:
        return [reg]

    parts: list[Region] = []
    cursor = reg.r0
    for br in banner_rows:
        if br > cursor:
            parts.append(Region(cursor, reg.c0, br - 1, reg.c1))
        parts.append(Region(br, reg.c0, br, reg.c1))  # バナー単独
        cursor = br + 1
    if cursor <= reg.r1:
        parts.append(Region(cursor, reg.c0, reg.r1, reg.c1))

    return parts if len(parts) > 1 else [reg]


# --------------------------------------------------------------------------- #
# 暫定ロール(可視化用。確定は ③ で行う)
# --------------------------------------------------------------------------- #
def guess_role(sheet: SheetModel, reg: Region) -> str:
    n_img = sum(
        sheet.image_cells.get((r, c), 0)
        for r in range(reg.r0, reg.r1 + 1)
        for c in range(reg.c0, reg.c1 + 1)
    )
    origins = {
        sheet.origin_of(r, c)
        for r in range(reg.r0, reg.r1 + 1) for c in range(reg.c0, reg.c1 + 1)
        if sheet.text_at(r, c) != ""
    }
    n_text = len(origins)
    if n_img and n_text <= 1:
        return "figure?"

    # 単一の論理セル(全幅バナー含む): 短文=見出し / 長文=段落
    if n_text == 1:
        only = sheet.origin_text.get(next(iter(origins)), "")
        return "heading?" if len(only) <= 20 else "paragraph?"

    # 2 列で左がラベルらしい → KeyValue
    if reg.n_cols == 2 and reg.n_rows >= 2:
        return "kv?"

    # 2x2 以上 → 表
    if reg.n_rows >= 2 and reg.n_cols >= 2:
        return "table?"

    return "?"
