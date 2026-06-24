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


# セル内オフセット(EMU)を分数セルに直すための公称サイズ。
# 厳密な列幅・行高の解決は避け、近似で十分(最近傍マッチングに使うだけ)。
_EMU_PER_COL = 609600    # 既定列幅 ≈ 64px
_EMU_PER_ROW = 190500    # 既定行高 ≈ 20px
# 端点を図形に対応づける許容距離(セル単位)
_MATCH_TOL = 3.0
# sp でも「線/コネクタ/矢印」の preset はエッジ候補として扱う
_CONNECTOR_GEOMS = {
    "line", "straightConnector1",
    "bentConnector2", "bentConnector3", "bentConnector4", "bentConnector5",
    "curvedConnector2", "curvedConnector3", "curvedConnector4", "curvedConnector5",
    "rightArrow", "leftArrow", "upArrow", "downArrow",
    "leftRightArrow", "upDownArrow", "bentArrow", "bentUpArrow", "curvedRightArrow",
}


# --------------------------------------------------------------------------- #
# drawing XML のパース
# --------------------------------------------------------------------------- #
def _parse_drawing(data: bytes) -> SheetDiagram | None:
    root = ET.fromstring(data)
    nodes: list[Node] = []
    node_boxes: dict[str, tuple[float, float, float, float]] = {}
    # エッジ候補: ("explicit", src, dst, label) or ("geom", p_from, p_to, label)
    candidates: list[tuple] = []

    for anchor in root.iter(f"{{{_NS['xdr']}}}twoCellAnchor"):
        _parse_anchor(anchor, nodes, node_boxes, candidates)
    for anchor in root.iter(f"{{{_NS['xdr']}}}oneCellAnchor"):
        _parse_anchor(anchor, nodes, node_boxes, candidates)

    if not nodes:
        return None

    edges: list[Edge] = []
    skipped = 0
    for cand in candidates:
        if cand[0] == "explicit":
            _, src, dst, label = cand
            if src in node_boxes and dst in node_boxes and src != dst:
                edges.append(Edge(src=src, dst=dst, label=label))
            else:
                skipped += 1
        else:  # 幾何推定: 端点に最も近い図形へ対応づけ
            _, p_from, p_to, label = cand
            src = _nearest(p_from, node_boxes)
            dst = _nearest(p_to, node_boxes)
            if src and dst and src != dst:
                edges.append(Edge(src=src, dst=dst, label=label))
            else:
                skipped += 1

    anchor_row = min(n.row for n in nodes)
    anchor_col = min(n.col for n in nodes)
    return SheetDiagram(anchor_row, anchor_col, nodes, edges, skipped)


def _parse_anchor(anchor, nodes, node_boxes, candidates) -> None:
    p_from = _marker(anchor, "from")
    p_to = _marker(anchor, "to") or p_from
    if p_from is None:
        return

    sp = anchor.find("xdr:sp", _NS)
    cxn = anchor.find("xdr:cxnSp", _NS)

    # コネクタ(cxnSp) or 線/矢印形状(sp) → エッジ候補
    if cxn is not None:
        candidates.append(_edge_candidate(cxn, p_from, p_to))
        return
    if sp is not None and _geom(sp) in _CONNECTOR_GEOMS:
        candidates.append(_edge_candidate(sp, p_from, p_to))
        return

    # それ以外の sp → ノード(図形)
    if sp is not None:
        nid = _shape_id(sp)
        nodes.append(Node(id=nid, text=_shape_text(sp), geom=_geom(sp),
                          row=int(p_from[0]) + 1, col=int(p_from[1]) + 1))
        node_boxes[nid] = (p_from[0], p_from[1], p_to[0], p_to[1])


def _edge_candidate(el, p_from, p_to) -> tuple:
    """明示接続があれば ('explicit',...)、無ければ ('geom',...) を返す。"""
    cxpr = el.find("xdr:nvCxnSpPr/xdr:cNvCxnSpPr", _NS)
    st = cxpr.find("a:stCxn", _NS) if cxpr is not None else None
    en = cxpr.find("a:endCxn", _NS) if cxpr is not None else None
    label = _shape_text(el)
    if st is not None and en is not None:
        return ("explicit", st.get("id"), en.get("id"), label)
    return ("geom", p_from, p_to, label)


def _marker(anchor, tag) -> tuple[float, float] | None:
    """xdr:from / xdr:to を分数セル座標 (row, col) で返す。"""
    m = anchor.find(f"xdr:{tag}", _NS)
    if m is None:
        return None
    col = int(m.findtext("xdr:col", "0", _NS))
    coff = int(m.findtext("xdr:colOff", "0", _NS))
    row = int(m.findtext("xdr:row", "0", _NS))
    roff = int(m.findtext("xdr:rowOff", "0", _NS))
    return (row + roff / _EMU_PER_ROW, col + coff / _EMU_PER_COL)


def _nearest(point, boxes) -> str | None:
    """点に最も近い図形 id を返す(許容距離 _MATCH_TOL を超えたら None)。"""
    best, best_d = None, float("inf")
    for nid, box in boxes.items():
        d = _point_box_dist(point, box)
        if d < best_d:
            best, best_d = nid, d
    return best if best_d <= _MATCH_TOL else None


def _point_box_dist(point, box) -> float:
    pr, pc = point
    r0, c0, r1, c1 = box
    if r1 < r0:
        r0, r1 = r1, r0
    if c1 < c0:
        c0, c1 = c1, c0
    dr = max(r0 - pr, 0.0, pr - r1)
    dc = max(c0 - pc, 0.0, pc - c1)
    return (dr * dr + dc * dc) ** 0.5


def _shape_id(sp) -> str:
    cnv = sp.find("xdr:nvSpPr/xdr:cNvPr", _NS)
    if cnv is None:
        cnv = sp.find("xdr:nvCxnSpPr/xdr:cNvPr", _NS)
    return cnv.get("id") if cnv is not None else ""


def _shape_text(sp) -> str:
    texts = [t.text or "" for t in sp.findall("xdr:txBody/a:p/a:r/a:t", _NS)]
    if not texts:
        texts = [t.text or "" for t in sp.findall(".//a:t", _NS)]
    return "".join(texts).strip()


def _geom(sp) -> str:
    g = sp.find("xdr:spPr/a:prstGeom", _NS)
    return g.get("prst") if g is not None else "rect"
