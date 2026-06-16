"""分割結果の可視化。実ファイルで「②分割が正しいか」を目視確認するためのダンプ。"""

from __future__ import annotations

from .extract import SheetModel
from .segment import Region, guess_role


def dump_segments(sheet: SheetModel, regions: list[Region]) -> str:
    lines = [
        f"# シート: {sheet.name}",
        f"  使用領域: {sheet.addr(sheet.min_row, sheet.min_col)}"
        f":{sheet.addr(sheet.max_row, sheet.max_col)}"
        f"  ブロック数: {len(regions)}",
        "",
    ]
    for i, reg in enumerate(regions, start=1):
        rng = f"{sheet.addr(reg.r0, reg.c0)}:{sheet.addr(reg.r1, reg.c1)}"
        role = guess_role(sheet, reg)
        preview = _preview(sheet, reg)
        lines.append(
            f"  [{i:02d}] {rng:<12} {reg.n_rows}x{reg.n_cols} "
            f"depth={reg.depth} role={role:<11} | {preview}"
        )
    lines.append("")
    return "\n".join(lines)


def _preview(sheet: SheetModel, reg: Region, limit: int = 70) -> str:
    texts = []
    for r in range(reg.r0, reg.r1 + 1):
        for c in range(reg.c0, reg.c1 + 1):
            t = sheet.text_at(r, c)
            if t and (not texts or texts[-1] != t):
                texts.append(t.replace("\n", "↵"))
    joined = " / ".join(texts)
    if len(joined) > limit:
        joined = joined[:limit] + "…"
    return joined or "(画像/空)"
