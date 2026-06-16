"""逆バティム検証: 元セルの非空テキストが出力 Markdown に全て現れるか照合する。

ルールベース(ブロック取りこぼし)でも、将来の LLM(欠落/改変)でも同じ網で検出する。
比較は空白・記号・HTML エスケープの差を吸収するため、両者を正規化して部分一致を見る。
"""

from __future__ import annotations

import re
from html import unescape

from .extract import SheetModel

_NORM_RE = re.compile(r"\s+")


def verify_content(sheet: SheetModel, markdown: str) -> list[str]:
    """出力に現れない元テキストの一覧(欠落)を返す。空なら完全保全。"""
    haystack = _normalize(markdown)
    missing: list[str] = []
    for (r, c), text in sheet.origin_text.items():
        needle = _normalize(text)
        if needle and needle not in haystack:
            missing.append(f"{sheet.addr(r, c)}: {text[:40]}")
    return missing


def _normalize(s: str) -> str:
    s = unescape(s)
    # markdown/HTML の構造記号を除去してから空白を畳む
    s = re.sub(r"</?td[^>]*>|</?tr>|</?table>|<br\s*/?>", "", s)
    s = s.replace("|", "").replace("#", "").replace("*", "").replace("\\", "")
    return _NORM_RE.sub("", s)
