"""③ 役割付与: Region 列を意味ブロック(Block)列に変換する(ルールベース)。

将来 LLM に差し替えられるよう、interpret() という関数 1 個を差込口とする。
RuleBasedInterpreter = この関数。LLMInterpreter は同じ (sheet, regions, image_map)
を受けて同じ Block 列を返すよう実装すればよい。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .extract import SheetModel
from .segment import Region

# 見出し番号: "1." "1.1" "1.2.3" "(1)" など先頭の章番号
_NUM_RE = re.compile(r"^\s*\(?(\d+(?:\.\d+)*)")
# 段落とみなす最大文字数(これ以下の単一セルは見出し寄り)
_HEADING_MAXLEN = 25


# --- Block 定義 ---
@dataclass
class Heading:
    level: int
    text: str


@dataclass
class Paragraph:
    text: str


@dataclass
class Table:
    region: Region          # 描画は render 側で sheet を参照して行う
    has_merges: bool


@dataclass
class KeyValue:
    pairs: list[tuple[str, str]] = field(default_factory=list)


@dataclass
class Figure:
    images: list[str] = field(default_factory=list)


@dataclass
class Diagram:
    mermaid: str            # 描画オブジェクトから復元した Mermaid flowchart


Block = Heading | Paragraph | Table | KeyValue | Figure | Diagram


def interpret(sheet: SheetModel, regions: list[Region],
              image_map: dict[tuple[int, int], list[str]]) -> list[Block]:
    blocks: list[Block] = []
    for reg in regions:
        blocks.append(_classify(sheet, reg, image_map))
    return blocks


def _classify(sheet: SheetModel, reg: Region,
              image_map: dict[tuple[int, int], list[str]]) -> Block:
    origins = _text_origins(sheet, reg)
    images = _images_in(reg, image_map)

    # 図: 画像があり、テキストがほぼ無い
    if images and len(origins) <= 1:
        return Figure(images=images)

    region_ruled = _region_has_ruled(sheet, reg)

    # 表を最優先で保護: 罫線で囲まれた複数セル領域は必ず表として扱う
    if region_ruled and len(origins) >= 2 and (reg.n_rows >= 2 or reg.n_cols >= 2):
        return Table(region=reg, has_merges=_has_merges(sheet, reg))

    # 単一の論理セル: 積極証拠があれば見出し / それ以外は段落
    if len(origins) == 1:
        o = origins[0]
        text = sheet.origin_text[o]
        if _is_heading(sheet, o, text, _is_full_width(sheet, reg, o)):
            return Heading(level=_heading_level(text), text=text)
        return Paragraph(text=text)

    # 罫線なしの 2 列ラベル/値 → KeyValue(罫線ありなら表として扱う)
    if (not region_ruled and reg.n_cols == 2 and reg.n_rows >= 2
            and _looks_like_kv(sheet, reg)):
        return KeyValue(pairs=_kv_pairs(sheet, reg))

    # 2x2 以上 → 表
    if reg.n_rows >= 2 and reg.n_cols >= 2:
        return Table(region=reg, has_merges=_has_merges(sheet, reg))

    # フォールバック: テキストを段落として連結
    joined = " ".join(sheet.origin_text[o] for o in origins)
    return Paragraph(text=joined)


def _region_has_ruled(sheet: SheetModel, reg: Region) -> bool:
    if not sheet.borders_are_structural:
        return False
    return any(
        sheet.is_ruled(r, c)
        for r in range(reg.r0, reg.r1 + 1) for c in range(reg.c0, reg.c1 + 1)
    )


def _is_heading(sheet: SheetModel, origin: tuple[int, int], text: str,
                full_width: bool) -> bool:
    """見出しは積極証拠を要求する(短いだけでは見出しにしない)。

    証拠: 章番号 / 太字・塗り・大フォント / 全幅バナー(短文に限る)。
    長文は全幅でも段落とみなす。
    """
    if _NUM_RE.match(text):
        return True                      # 章番号は強い証拠
    if len(text) > _HEADING_MAXLEN:
        return False                     # 長文は段落(全幅でも)
    # 表セル(罫線つき)は見出しにしない
    if sheet.is_ruled(*origin):
        return False
    if full_width:
        return True                      # 全幅バナーの短文 = 見出し
    st = sheet.style.get(origin)
    if st is None:
        return False
    larger = bool(st.font_size and sheet.body_font_size
                  and st.font_size > sheet.body_font_size)
    return bool(st.bold or st.has_fill or larger)


def _is_full_width(sheet: SheetModel, reg: Region, origin: tuple[int, int]) -> bool:
    """origin が領域の全幅を覆う単一セル(=バナー)か。"""
    if origin[1] != reg.c0:
        return False
    _, cs = sheet.span_of.get(origin, (1, 1))
    return cs >= reg.n_cols


# --- 補助 ---
def _text_origins(sheet: SheetModel, reg: Region) -> list[tuple[int, int]]:
    seen: list[tuple[int, int]] = []
    found = set()
    for r in range(reg.r0, reg.r1 + 1):
        for c in range(reg.c0, reg.c1 + 1):
            if sheet.text_at(r, c) == "":
                continue
            o = sheet.origin_of(r, c)
            if o not in found:
                found.add(o)
                seen.append(o)
    return seen


def _images_in(reg: Region, image_map: dict[tuple[int, int], list[str]]) -> list[str]:
    out: list[str] = []
    for (r, c), names in image_map.items():
        if reg.r0 <= r <= reg.r1 and reg.c0 <= c <= reg.c1:
            out.extend(names)
    return out


def _has_merges(sheet: SheetModel, reg: Region) -> bool:
    for r in range(reg.r0, reg.r1 + 1):
        for c in range(reg.c0, reg.c1 + 1):
            if (r, c) in sheet.covered_to_origin:
                return True
            rs, cs = sheet.span_of.get((r, c), (1, 1))
            if rs > 1 or cs > 1:
                return True
    return False


def _looks_like_kv(sheet: SheetModel, reg: Region) -> bool:
    """左列が短いラベル(主に <=15 文字)中心なら KeyValue とみなす。"""
    left = reg.c0
    labels = [sheet.text_at(r, left) for r in range(reg.r0, reg.r1 + 1)]
    labels = [t for t in labels if t]
    if not labels:
        return False
    return all(len(t) <= 15 and "\n" not in t for t in labels)


def _kv_pairs(sheet: SheetModel, reg: Region) -> list[tuple[str, str]]:
    pairs: list[tuple[str, str]] = []
    for r in range(reg.r0, reg.r1 + 1):
        label = sheet.text_at(r, reg.c0)
        value = sheet.text_at(r, reg.c1)
        if label or value:
            pairs.append((label, value))
    return pairs


def _heading_level(text: str) -> int:
    m = _NUM_RE.match(text)
    if m:
        # "1"->2(##), "1.1"->3(###), ...
        return min(2 + m.group(1).count("."), 6)
    return 2
