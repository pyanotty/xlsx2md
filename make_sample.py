"""動作確認用のサンプル xlsx を生成する。"""
import base64
import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage

# 1x1 透明 PNG
_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M8AAAMBAQDJ"
    "/pLvAAAAAElFTkSuQmCC"
)

out = Path("sample.xlsx")
img_path = Path("_tmp_sample.png")
img_path.write_bytes(_PNG)

wb = Workbook()

# --- シート1: 要件定義書ふう(結合フォーム + データ表 + 画像) ---
ws1 = wb.active
ws1.title = "機能要件"
ws1["A1"] = "機能要件一覧"
ws1.merge_cells("A1:D1")               # 見出しの横結合
ws1["A2"] = "項目"
ws1["B2"] = "内容"
ws1.merge_cells("B2:D2")
ws1["A3"] = "機能名"
ws1["B3"] = "ログイン\n(認証)"          # セル内改行
ws1.merge_cells("B3:D3")
ws1["A4"] = "優先度"
ws1["B4"] = "高"
ws1["A5"] = "作成日"
ws1["B5"] = datetime.datetime(2026, 6, 15, 10, 30, 0)
ws1["A6"] = "件数"
ws1["B6"] = "=3+4"                      # 数式(未計算 -> 式文字列フォールバック検証)
ws1["A7"] = "参照"
ws1["B7"] = "公式サイト"
ws1["B7"].hyperlink = "https://example.com/spec"
# 縦結合
ws1["A8"] = "区分"
ws1.merge_cells("A8:A9")
ws1["B8"] = "画面"
ws1["B9"] = "帳票"
ws1.add_image(XLImage(str(img_path)), "C8")  # C8 にアンカーした画像

# --- シート2: 画面設計書ふう(複数セクション/空行区切り) ---
ws2 = wb.create_sheet("画面設計 ログイン")  # シート名に空白を含める
# セクション1: 見出しバナー
ws2["B2"] = "1. ログイン画面"
ws2.merge_cells("B2:E2")
# 説明段落(全幅結合の長文)
ws2["B3"] = "本画面はユーザ認証を行う。ID とパスワードを入力し、ログインボタンを押下する。"
ws2.merge_cells("B3:E3")
# (4行目は空 → セクション区切り)
# セクション2: 小見出し + 罫線つきデータ表
ws2["B5"] = "1.1 入力項目"
ws2.merge_cells("B5:E5")
ws2["B6"] = "項目"
ws2["C6"] = "種別"
ws2["D6"] = "必須"
ws2["E6"] = "備考"
ws2["B7"] = "ID"
ws2["C7"] = "テキスト"
ws2["D7"] = "○"
ws2["E7"] = "半角英数"
ws2["B8"] = "PW"
ws2["C8"] = "パスワード"
ws2["D8"] = "○"
ws2["E8"] = "8文字以上"
# (9行目は空 → セクション区切り)
# セクション3: 図
ws2["B10"] = "1.2 画面イメージ"
ws2.merge_cells("B10:E10")
ws2.add_image(XLImage(str(img_path)), "B11")

wb.save(out)
print("generated", out)
