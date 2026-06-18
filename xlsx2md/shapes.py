"""画面遷移図などの描画オブジェクト抽出(openpyxl 外の経路)。

openpyxl は図形(sp)・コネクタ(cxnSp)をほぼ破棄するため、xlsx(=zip)内の
`xl/drawings/drawingN.xml` を直接パースする。

フェーズ1: コネクタに明示接続(stCxn/endCxn)がある矢印のみをエッジとして採用する。
手置き(座標のみ)の矢印は幾何推定が要るため、ここでは対象外(将来フェーズ2)。
"""

from __future__ import annotations

import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass, field

_NS = {
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
}


@dataclass
class Node:
    id: str                  # 図形の cNvPr id(コネクタ接続のキー)
    text: str
    geom: str                # prstGeom の prst(rect/roundRect/diamond...)
    row: int                 # アンカー先セル(1-indexed)
    col: int


@dataclass
class Edge:
    src: str                 # 始点ノード id
    dst: str                 # 終点ノード id
    label: str = ""


@dataclass
class SheetDiagram:
    anchor_row: int
    anchor_col: int
    nodes: list[Node] = field(default_factory=list)
    edges: list[Edge] = field(default_factory=list)
    skipped_connectors: int = 0   # 明示接続が無く落とした矢印の数


def extract_diagrams(xlsx_path) -> dict[str, list[SheetDiagram]]:
    """シート名 -> その上の図(SheetDiagram のリスト)。"""
    out: dict[str, list[SheetDiagram]] = {}
    with zipfile.ZipFile(xlsx_path) as z:
        names = set(z.namelist())
        for sheet_name, sheet_part in _sheet_parts(z, names):
            drawing_part = _drawing_for_sheet(z, names, sheet_part)
            if drawing_part is None or drawing_part not in names:
                continue
            diagram = _parse_drawing(z.read(drawing_part))
            if diagram and diagram.nodes:
                out.setdefault(sheet_name, []).append(diagram)
    return out


# --------------------------------------------------------------------------- #
# リレーション解決(workbook → sheet → drawing)
# --------------------------------------------------------------------------- #
def _sheet_parts(z, names):
    """[(シート名, 'xl/worksheets/sheetN.xml'), ...] を出現順で返す。"""
    if "xl/workbook.xml" not in names:
        return []
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = _rels(z, names, "xl/_rels/workbook.xml.rels")
    result = []
    for sheet in wb.findall("ct:sheets/ct:sheet", _NS):
        name = sheet.get("name")
        rid = sheet.get(f"{{{_NS['r']}}}id")
        target = rels.get(rid)
        if target:
            result.append((name, _resolve("xl/workbook.xml", target)))
    return result


def _drawing_for_sheet(z, names, sheet_part) -> str | None:
    base = sheet_part.rsplit("/", 1)[-1]              # sheetN.xml
    rels_part = sheet_part.rsplit("/", 1)[0] + f"/_rels/{base}.rels"
    if rels_part not in names:
        return None
    for target in _rels(z, names, rels_part).values():
        if "drawing" in target:
            return _resolve(sheet_part, target)
    return None


def _rels(z, names, part) -> dict[str, str]:
    if part not in names:
        return {}
    root = ET.fromstring(z.read(part))
    return {rel.get("Id"): rel.get("Target")
            for rel in root.findall("pr:Relationship", _NS)}


def _resolve(owner_part: str, target: str) -> str:
    """rels の Target を所有パス基準で zip 内パスに解決する。"""
    if target.startswith("/"):
        return target.lstrip("/")            # 絶対(パッケージルート基準)
    base_dir = owner_part.rsplit("/", 1)[0] if "/" in owner_part else ""
    return _norm(f"{base_dir}/{target}" if base_dir else target)


def _norm(path: str) -> str:
    # "xl/worksheets/../drawings/x" → "xl/drawings/x"
    parts: list[str] = []
    for seg in path.split("/"):
        if seg == "..":
            if parts:
                parts.pop()
        else:
            parts.append(seg)
    return "/".join(parts)


# --------------------------------------------------------------------------- #
# drawing XML のパース
# --------------------------------------------------------------------------- #
def _parse_drawing(data: bytes) -> SheetDiagram | None:
    root = ET.fromstring(data)
    nodes: list[Node] = []
    edges: list[Edge] = []
    skipped = 0

    for anchor in root:
        frm = anchor.find("xdr:from", _NS)
        row = int(frm.findtext("xdr:row", "0", _NS)) + 1 if frm is not None else 1
        col = int(frm.findtext("xdr:col", "0", _NS)) + 1 if frm is not None else 1

        sp = anchor.find("xdr:sp", _NS)
        if sp is not None:
            nodes.append(Node(
                id=_shape_id(sp),
                text=_shape_text(sp),
                geom=_geom(sp),
                row=row, col=col,
            ))
            continue

        cxn = anchor.find("xdr:cxnSp", _NS)
        if cxn is not None:
            cxpr = cxn.find("xdr:nvCxnSpPr/xdr:cNvCxnSpPr", _NS)
            st = cxpr.find("a:stCxn", _NS) if cxpr is not None else None
            en = cxpr.find("a:endCxn", _NS) if cxpr is not None else None
            if st is not None and en is not None:
                edges.append(Edge(
                    src=st.get("id"), dst=en.get("id"),
                    label=_shape_text(cxn),
                ))
            else:
                skipped += 1   # 明示接続なし → フェーズ1では落とす

    if not nodes:
        return None
    anchor_row = min(n.row for n in nodes)
    anchor_col = min(n.col for n in nodes)
    return SheetDiagram(anchor_row, anchor_col, nodes, edges, skipped)


def _shape_id(sp) -> str:
    cnv = sp.find("xdr:nvSpPr/xdr:cNvPr", _NS)
    return cnv.get("id") if cnv is not None else ""


def _shape_text(sp) -> str:
    texts = [t.text or "" for t in sp.findall("xdr:txBody/a:p/a:r/a:t", _NS)]
    # 段落区切り(改行)も拾う
    if not texts:
        texts = [t.text or "" for t in sp.findall(".//a:t", _NS)]
    return "".join(texts).strip()


def _geom(sp) -> str:
    g = sp.find("xdr:spPr/a:prstGeom", _NS)
    return g.get("prst") if g is not None else "rect"
