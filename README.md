# xlsx2md

Excel 文書(アプリケーションの要件定義書・画面設計書を想定)を Markdown に変換する Python モジュール。
要件定義書は [xlsx2md_requirements.md](xlsx2md_requirements.md) を参照。

## 特徴

- **シートごとに独立した .md** を出力(各シート = 独立した文書)
- **1 シート = 1 つの HTML `<table>`** として忠実再現(LLM 入力向け)
- **セル結合**を `colspan` / `rowspan` に変換
- **画像(PNG/JPEG)** を抽出し、アンカー先セルにプレースホルダを差し込む
- 数式はキャッシュ値、無ければ式文字列にフォールバック
- 使用領域の外周空白のみトリミング(内部レイアウトは保持)

## セットアップ

```bash
python3 -m venv .venv
./.venv/bin/pip install -r requirements.txt   # openpyxl (画像生成テストには Pillow も)
```

## 使い方(CLI)

```bash
python -m xlsx2md path/to/spec.xlsx -o output
```

出力構成:

```
output/
  spec/
    01_シート名.md
    02_シート名.md
    images/
      シート名_img1.png
```

主なオプション:

| オプション | 説明 |
|---|---|
| `-o, --out-dir` | 出力先ディレクトリ(既定: `output`) |
| `--image-dirname` | 画像サブディレクトリ名(既定: `images`) |
| `--no-hidden` | 隠し行/列を出力に含めない |
| `--no-formula-fallback` | 数式キャッシュが無い時に式文字列へフォールバックしない |
| `-v, --verbose` | 詳細ログ |

## 使い方(ライブラリ)

```python
from xlsx2md import convert, Options

paths = convert("spec.xlsx", "output", Options(include_hidden=True))
print(paths)  # 生成した .md のパス一覧
```

## 動作確認

```bash
./.venv/bin/python make_sample.py          # sample.xlsx を生成
./.venv/bin/python -m xlsx2md sample.xlsx -o output -v
```

## スコープ外(初版)

グラフ、図形・テキストボックス・矢印などの描画オブジェクト、セルのコメント、
見た目書式(色・太字・罫線)。詳細は要件定義書の「4. スコープ外」を参照。
