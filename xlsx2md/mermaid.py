"""SheetDiagram(ノード＋エッジ) → Mermaid flowchart 文字列。"""

from __future__ import annotations

from .shapes import Node, SheetDiagram

# prstGeom(preset geometry) → Mermaid ノード形状(前後の囲み記号)
_GEOM_WRAP = {
    "roundRect": ("(", ")"),               # 角丸四角 → ( )
    "rect": ("[", "]"),                    # 四角 → [ ]
    "flowChartProcess": ("[", "]"),
    "diamond": ("{", "}"),                 # 菱形(判断) → { }
    "flowChartDecision": ("{", "}"),
    "ellipse": ("((", "))"),               # 楕円 → 円
    "flowChartTerminator": ("([", "])"),   # 端子 → スタジアム
}
_DEFAULT_WRAP = ("[", "]")


def build_mermaid(diagram: SheetDiagram) -> str:
    # ノード id → Mermaid 用の安全な識別子(n1, n2, ...)
    name = {n.id: f"n{i}" for i, n in enumerate(diagram.nodes, start=1)}

    lines = [f"flowchart {_direction(diagram.nodes)}"]
    for n in diagram.nodes:
        lo, hi = _GEOM_WRAP.get(n.geom, _DEFAULT_WRAP)
        lines.append(f"  {name[n.id]}{lo}{_label(n.text)}{hi}")
    for e in diagram.edges:
        if e.src not in name or e.dst not in name:
            continue
        arrow = f"  {name[e.src]} -->"
        if e.label:
            arrow += f"|{_edge_label(e.label)}|"
        lines.append(f"{arrow} {name[e.dst]}")
    return "\n".join(lines)


def _direction(nodes: list[Node]) -> str:
    """図形の広がりが横長なら左→右(LR)、縦長なら上→下(TD)。"""
    rows = [n.row for n in nodes]
    cols = [n.col for n in nodes]
    spread_r = max(rows) - min(rows)
    spread_c = max(cols) - min(cols)
    return "LR" if spread_c > spread_r else "TD"


def _label(text: str) -> str:
    """ノードラベル: 改行は <br/>、" はエスケープ、全体を "..." で囲む。"""
    t = text.replace('"', "'").replace("\n", "<br/>")
    return f'"{t or " "}"'


def _edge_label(text: str) -> str:
    return text.replace("\n", " ").replace("|", "/").strip()
