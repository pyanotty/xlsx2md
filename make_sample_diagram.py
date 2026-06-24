"""動作確認用: 明示コネクタ付きの画面遷移図を持つ xlsx を生成する。

openpyxl は図形・コネクタを書き出せないため、保存後に zip へ drawing XML を注入する。
図: (角丸)ログイン画面 → (菱形)認証成功? --OK--> (四角)トップ画面
                                      └--NG--> ログイン画面 へ戻る
"""
import shutil
import zipfile
from pathlib import Path

from openpyxl import Workbook

SRC = Path("sample_diagram.xlsx")

wb = Workbook()
ws = wb.active
ws.title = "画面遷移"
ws["A1"] = "3. 画面遷移フロー"            # 見出し(図の前に出る → 併合順の確認)
wb.save(SRC)

# --- 注入する drawing XML(図形3 + コネクタ3、明示接続つき) ---
DRAWING = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp macro="" textlink="">
      <xdr:nvSpPr><xdr:cNvPr id="2" name="角丸四角形 1"/><xdr:cNvSpPr/></xdr:nvSpPr>
      <xdr:spPr><a:xfrm><a:off x="900000" y="900000"/><a:ext cx="1500000" cy="600000"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom></xdr:spPr>
      <xdr:txBody><a:bodyPr/><a:p><a:r><a:t>ログイン画面</a:t></a:r></a:p></xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>7</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp macro="" textlink="">
      <xdr:nvSpPr><xdr:cNvPr id="3" name="菱形 2"/><xdr:cNvSpPr/></xdr:nvSpPr>
      <xdr:spPr><a:xfrm><a:off x="900000" y="2700000"/><a:ext cx="1500000" cy="600000"/></a:xfrm><a:prstGeom prst="flowChartDecision"><a:avLst/></a:prstGeom></xdr:spPr>
      <xdr:txBody><a:bodyPr/><a:p><a:r><a:t>認証成功?</a:t></a:r></a:p></xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp macro="" textlink="">
      <xdr:nvSpPr><xdr:cNvPr id="4" name="四角形 3"/><xdr:cNvSpPr/></xdr:nvSpPr>
      <xdr:spPr><a:xfrm><a:off x="900000" y="4500000"/><a:ext cx="1500000" cy="600000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
      <xdr:txBody><a:bodyPr/><a:p><a:r><a:t>トップ画面</a:t></a:r></a:p></xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:cxnSp macro="">
      <xdr:nvCxnSpPr><xdr:cNvPr id="5" name="コネクタ 4"/><xdr:cNvCxnSpPr><a:stCxn id="2" idx="2"/><a:endCxn id="3" idx="0"/></xdr:cNvCxnSpPr></xdr:nvCxnSpPr>
      <xdr:spPr><a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom><a:ln><a:tailEnd type="triangle"/></a:ln></xdr:spPr>
    </xdr:cxnSp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>7</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:cxnSp macro="">
      <xdr:nvCxnSpPr><xdr:cNvPr id="6" name="コネクタ 5"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr>
      <xdr:spPr><a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom><a:ln><a:tailEnd type="triangle"/></a:ln></xdr:spPr>
      <xdr:txBody><a:bodyPr/><a:p><a:r><a:t>OK</a:t></a:r></a:p></xdr:txBody>
    </xdr:cxnSp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:cxnSp macro="">
      <xdr:nvCxnSpPr><xdr:cNvPr id="7" name="コネクタ 6"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr>
      <xdr:spPr><a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom><a:ln><a:tailEnd type="triangle"/></a:ln></xdr:spPr>
      <xdr:txBody><a:bodyPr/><a:p><a:r><a:t>NG</a:t></a:r></a:p></xdr:txBody>
    </xdr:cxnSp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"""

SHEET_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>"""

DRAWING_CT = ('<Override PartName="/xl/drawings/drawing1.xml" '
              'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>')

# --- zip を読み直し、drawing 関連を注入して書き戻す ---
tmp = SRC.with_suffix(".tmp.xlsx")
with zipfile.ZipFile(SRC) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.namelist():
        data = zin.read(item)
        if item == "[Content_Types].xml":
            data = data.replace(b"</Types>", DRAWING_CT.encode() + b"</Types>")
        zout.writestr(item, data)
    zout.writestr("xl/drawings/drawing1.xml", DRAWING)
    zout.writestr("xl/worksheets/_rels/sheet1.xml.rels", SHEET_RELS)

shutil.move(tmp, SRC)
print("generated", SRC)
