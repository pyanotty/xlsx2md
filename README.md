# xlsx2md

Excel で作成された**要件定義書・画面設計書**を、**設計書として自然な Markdown** に変換する Python モジュール。

Excel は「方眼紙」としてセル結合・罫線・配置を駆使し、見た目で文書構造（見出し・説明文・表・図）を表現している。本ツールはその**視覚レイアウトを手がかりに文書構造を復元**し、`#` 見出し → 説明段落 → 表 → 図 からなる自然な Markdown に組み直す。出力の主用途は **LLM への入力**。

要件定義は [xlsx2md_requirements.md](xlsx2md_requirements.md) を参照。

---

## 1. 設計思想

このタスクの本質は「セルの**転写**」ではなく「視覚レイアウトからの**文書構造の復元**」である。そのため次の原則を置く。

- **シートごとに独立した文書**として扱い、シートごとに 1 つの `.md` を出力する。
- **抽出（決定的）と解釈（推定的）を分離**し、段ごとに責務を切る。
- **表は「切ってはいけない原子単位」**。罫線を最重要シグナルとして表領域を保護する。
- **見出しは積極証拠を要求**する（短いだけでは見出しにしない）。
- **本文は逆バティム保全**。出力後に「元セルのテキストが全て出力に現れるか」を機械照合する。
- 解釈エンジン（③）は **`interpret()` 1 関数**を差込口とし、将来 LLM 版に差し替え可能にする。

---

## 2. パイプライン全体

```
xlsx
 │  ① extract   openpyxl で SheetModel を構築（値・結合・罫線・書式・画像）
 ▼
SheetModel
 │  ② segment   再帰 XY-cut で Region 列に分割（＝読み順）
 ▼
[Region, …]
 │  ③ interpret 各 Region に役割を付与 → Block 列（見出し/段落/表/KV/図）
 ▼
[Block, …]
 │  ⑤ render    Block を自然な Markdown に描画
 ▼
Markdown ──► verify  元セルテキストの保全を照合（欠落を警告）
```

| 段 | モジュール | 性質 |
|---|---|---|
| ① 抽出 | [extract.py](xlsx2md/extract.py) | 決定的 |
| ② 分割 | [segment.py](xlsx2md/segment.py) | ほぼ決定的 |
| ③ 役割付与 | [interpret.py](xlsx2md/interpret.py) | 推定的（ルールベース、将来 LLM 化可） |
| ⑤ 生成 | [render.py](xlsx2md/render.py) | 決定的 |
| 検証 | [verify.py](xlsx2md/verify.py) | 決定的 |
| 統合 | [pipeline.py](xlsx2md/pipeline.py) | — |

> v1 の「1 シート = 1 つの忠実な HTML テーブル」方式は `converter.render_faithful_table` として温存しており、`--faithful` で利用できる（突き合わせ・将来の LLM 入力用）。

---

## 3. 各段の詳細

### ① 抽出 — `SheetModel`

Worksheet を解釈しやすい構造に落とす。**出力に出さない罫線・書式も、③の判定材料として読む**のがポイント。

- `origin_text` … 結合の左上（origin）セルの非空テキスト。数式はキャッシュ値、無ければ式文字列。ハイパーリンクは `テキスト (URL)` で併記。
- `span_of` / `covered_to_origin` … セル結合の rowspan/colspan と被覆セル→origin の対応。
- `bordered` … **罫線のある全物理セル**（空セルも含む）。表領域の保護に使う。
- `borders_are_structural` … 罫線が「表の構造」か「方眼紙の全面装飾」かの判定。
- `style` / `body_font_size` … 太字・塗り・フォントサイズ（見出し判定の証拠）。

**最重要メソッド**：罫線つきの空セルを「非空」とみなすことで、後段の XY-cut が表を切り刻むのを防ぐ。

```python
def is_ruled(self, r, c):
    """罫線で囲まれた表セルか(空セルでも True)。装飾罫線は無視。"""
    if not self.borders_are_structural:
        return False
    return (r, c) in self.bordered or self.origin_of(r, c) in self.bordered

def is_blank(self, r, c):
    # 罫線つきの空セルも「表の一部」として非空扱い → XY-cut で粉砕させない
    return (self.text_at(r, c) == ""
            and not self.has_image(r, c)
            and not self.is_ruled(r, c))
```

罫線が表か装飾かは被覆率で判定する。表が主体のシートは被覆率が高くなるため、**装飾とみなすのは「ほぼ全セルが罫線」の極端な場合のみ**（保護を効かせる方向に倒す）。

```python
# 使用領域に対する罫線セルの割合が 0.9 未満なら「構造的な罫線(=表)」とみなす
borders_are_structural = bool(bordered) and area > 0 and (in_box / area) < 0.9
```

### ② 分割 — 再帰 XY-cut

#### なぜ XY-cut なのか

文書レイアウト解析の古典 **XY-cut**（再帰射影分割）を採用している。「内容のかたまりは余白で隔てられている」という文書の普遍的性質を使い、**領域を空白の行/列で再帰的に二分割していく**手法。設計書の Excel は「上から下へ、左から右へ」読む文書なので、この分割の再帰木の**走査順（pre-order DFS）がそのまま読み順**になるという利点がある。

> 代替案との比較：単純な「セルの連結成分」では、スペーサー列で隔てられた表の列がバラバラの成分になってしまう。v1 の「シート丸ごと 1 テーブル」はレイアウト座標は保つが文書構造を失う。XY-cut は**余白による意味の区切り**を捉えつつ読み順も得られる中庸点。

#### 3 種類の切り方とその順序

`_cut` は 1 つの領域（Region）に対し、次の順で切れるか試し、最初に成功した切り方で再帰する。

```python
def _cut(sheet, reg, out, depth):
    # 1) 空行で横分割（上→下）
    bands = _split_rows_by_blank(sheet, reg)
    if len(bands) > 1:
        for b in bands: _cut(sheet, b, out, depth + 1)
        return
    # 2) 全幅バナー行で横分割（見出し ↔ 本文の隣接を切る）
    fw = _split_full_width(sheet, reg)
    if len(fw) > 1:
        for b in fw: _cut(sheet, b, out, depth + 1)
        return
    # 3) 空列で縦分割（左→右）
    cols = _split_cols_by_blank(sheet, reg)
    if len(cols) > 1:
        for b in cols: _cut(sheet, b, out, depth + 1)
        return
    # それ以上切れない → 葉（1 ブロック）
    reg.depth = depth
    out.append(reg)
```

**この順序には理由がある。**

| 順 | 切り方 | 狙い | なぜこの順位か |
|---|---|---|---|
| 1 | 空行で横分割 | セクション・段落・表を縦方向に切り出す | 文書は基本「上から下」。最も粗い区切りである空行を最初に当てる |
| 2 | 全幅バナー行 | 見出しと、直下の本文/表を分離 | 見出しは空行を挟まず本文に隣接しがち。空行で切れなかった塊に対し見出しだけ剥がす |
| 3 | 空列で縦分割 | 左右に並んだ別要素（2 段組み等）を分離 | 縦割りは最後。先に縦で切ると表の列を割りやすいため、横方向を出し切ってから当てる |

`_split_rows_by_blank` は連続する非空行をバンドにまとめ、空行を区切りとして落とす。**先頭・末尾の空白だけのトリミングも同じ関数が担う**（区切りが 1 つでも、外周に空白があれば内側へ詰める）。

#### 罫線による表の保護（最重要）

XY-cut は「余白があれば切る」ため、放っておくと**表を列ごと・セルごとに粉砕**してしまう（特にスペーサー列や空の備考列で縦に割れる）。これを防ぐのが ① の `is_blank` で、**罫線つきの空セルを「非空」とみなす**。

```python
def _row_blank(sheet, r, c0, c1):
    return all(sheet.is_blank(r, c) for c in range(c0, c1 + 1))
# is_blank は text が空 かつ 画像なし かつ 罫線なし のときだけ True
```

結果、**罫線で囲まれた格子は「非空の塊」として一体化**し、内部の空行・空列で切られなくなる。これが「表が崩れる」問題への中心的な対処。

> セル結合も自然に保護される。被覆セルの `text_at` は結合の左上（origin）のテキストを返すため、内容のある結合は全行・全列が「非空」となり、結合をまたいで切られない。

#### 全幅バナー行の検出

見出しが表や本文と**空行なしで隣接**していても分離するための手当て。「行 = 領域幅いっぱいの単一結合セル」をバナーとみなし、その行だけを独立領域に切り出す。

```python
def _is_full_width_banner(sheet, r, c0, c1):
    """行 r が「領域幅いっぱいに広がる単一結合セル」のみで構成されるか。"""
    origin = sheet.origin_of(r, c0)
    if origin not in sheet.origin_text:
        return False
    rs, cs = sheet.span_of.get(origin, (1, 1))
    return origin[1] == c0 and cs >= (c1 - c0 + 1)
```

#### 分割トレース例（シート「詳細設計」）

```
使用領域 B2:E7
└─ ① 空行で横分割（行3が空）
   ├─ B2:E2  … バナー「2. 詳細設計」          → 葉（depth 1）
   └─ B4:E7  … 罫線つき表（D列は空だが罫線あり）
        ① 空行なし（罫線で全行が非空）
        ② バナーなし（B4「項目」は全幅ではない）
        ③ 空列なし（D列も罫線で非空）          → 葉（depth 1, 1 ブロック=表）
```

D 列が**罫線で保護**されているため ③ の縦分割が発火せず、表が 1 ブロックのまま残る点が肝。`--segment` で各葉の範囲・暫定ロールを確認できる。

#### 既知の弱点

- 罫線が**全面に引かれた**シートは ① で「装飾」と判定され保護が外れ得る（§7-1）。
- **罫線のない（整列だけの）表**は保護されず、内部の空列で割れ得る（§7-2）。
- 暫定ロール `guess_role()` は可視化専用で、確定は ③ が行う。

### ③ 役割付与 — `interpret()`

#### 役割（Block）と差込口

各 Region を 5 種の Block のいずれかに分類する。

| Block | 意味 | 出力先（⑤） |
|---|---|---|
| `Heading(level, text)` | 見出し | `##` 等 |
| `Paragraph(text)` | 説明文 | 本文 |
| `Table(region, has_merges)` | 表 | GFM / HTML |
| `KeyValue(pairs)` | ラベル＝値の組 | 箇条書き |
| `Figure(images)` | 図 | `![]()` |

解釈は **`interpret()` 関数 1 つが差込口**。現在は決定的なルールベース（`RuleBasedInterpreter` 相当）だが、同じ `(sheet, regions, image_map) → list[Block]` のシグネチャで**将来 LLM 版に差し替え可能**な設計にしている。

```python
def interpret(sheet, regions, image_map):
    return [_classify(sheet, reg, image_map) for reg in regions]
```

#### 判定の優先順序（`_classify`）

**順序そのものが誤分類を防ぐ防御線**になっている。上から評価し、最初に当たった役割を採用する。

```python
def _classify(sheet, reg, image_map):
    origins = _text_origins(sheet, reg)        # 領域内の非空 origin セル群
    images = _images_in(reg, image_map)

    # (1) 図: 画像があり、テキストがほぼ無い
    if images and len(origins) <= 1:
        return Figure(images=images)

    region_ruled = _region_has_ruled(sheet, reg)

    # (2) 表を最優先で保護: 罫線で囲まれた複数セル領域は必ず表
    if region_ruled and len(origins) >= 2 and (reg.n_rows >= 2 or reg.n_cols >= 2):
        return Table(region=reg, has_merges=_has_merges(sheet, reg))

    # (3) 単一の論理セル: 積極証拠があれば見出し / それ以外は段落
    if len(origins) == 1:
        o = origins[0]; text = sheet.origin_text[o]
        if _is_heading(sheet, o, text, _is_full_width(sheet, reg, o)):
            return Heading(level=_heading_level(text), text=text)
        return Paragraph(text=text)

    # (4) 罫線なしの 2 列ラベル/値 → KeyValue（罫線ありなら表）
    if (not region_ruled and reg.n_cols == 2 and reg.n_rows >= 2
            and _looks_like_kv(sheet, reg)):
        return KeyValue(pairs=_kv_pairs(sheet, reg))

    # (5) 2x2 以上 → 表
    if reg.n_rows >= 2 and reg.n_cols >= 2:
        return Table(region=reg, has_merges=_has_merges(sheet, reg))

    # (6) フォールバック: テキストを段落として連結
    return Paragraph(text=" ".join(sheet.origin_text[o] for o in origins))
```

各分岐の意図：

- **(1) 図を最初に** … 画像セルを表やフォールバックに巻き込まれる前に確定。
- **(2) 表保護を見出し判定より前に** … これが「表のヘッダ/カラムが見出しに化ける」事故の主因への対処。罫線つきで 2 セル以上ある領域は、内訳を問わず先に表として確定させ、後段の見出し判定に渡さない。
- **(3) 単一セル** … 中身が 1 つの論理セルだけの領域。ここで**初めて見出しを検討**する（後述の積極証拠つき）。
- **(4) KeyValue は罫線なしに限定** … 罫線つき 2 列は (2) で既に表になっている。罫線のない「項目｜値」の素の 2 列だけを箇条書きにする。
- **(5) 一般の表** … 罫線がなくても複数行×複数列なら表とみなす。
- **(6) フォールバック** … いずれにも当たらない雑多な領域はテキスト連結で段落に。

`has_merges` が表の描画先（GFM か HTML か）を決める。領域内に結合が 1 つでもあれば HTML、なければ GFM。

#### 見出しは「積極証拠」を要求する

旧実装の「短いセルは見出し」は、孤立した表セルを軒並み `##` 化していた。現在は**短いだけでは見出しにせず、いずれかの積極証拠を要求**する。

```python
def _is_heading(sheet, origin, text, full_width):
    if _NUM_RE.match(text):          # 証拠A: "1." "1.1" "(1)" など章番号（最強）
        return True
    if len(text) > _HEADING_MAXLEN:  # 長文は段落（全幅でも）。MAXLEN=25
        return False
    if sheet.is_ruled(*origin):      # 表セル（罫線つき）は見出しにしない
        return False
    if full_width:                   # 証拠B: 全幅バナーの短文 = 見出し
        return True
    st = sheet.style.get(origin)     # 証拠C: 太字 / 塗り / 本文より大きいフォント
    larger = bool(st and st.font_size and sheet.body_font_size
                  and st.font_size > sheet.body_font_size)
    return bool(st and (st.bold or st.has_fill or larger))
```

判定の要点：

- **章番号（A）は単独で見出し確定**。`1.` `1.1` `1.2.3` `(1)` 等の先頭パターン（`_NUM_RE`）。
- **長文は除外**。全幅であっても 25 文字超なら段落（全幅結合の長い説明文を見出しにしない）。
- **罫線つきセルは除外**。表の一部が漏れて単一セル扱いになっても見出しにしない（二重の安全弁）。
- **全幅バナー（B）** … 領域幅いっぱいの短文。`_is_full_width` で「origin が行頭から領域全幅を覆う」かを見る。
- **装飾（C）** … 太字・背景塗り・本文（`body_font_size` = 最頻フォントサイズ）より大きい字。

| 例 | 判定 | 根拠 |
|---|---|---|
| `1. ログイン画面` | 見出し `##` | 章番号A |
| `1.1 入力項目` | 見出し `###` | 章番号A（レベルはドット数） |
| `機能要件一覧`（全幅結合・無装飾） | 見出し `##` | 全幅バナーB |
| `本画面はユーザ認証を行う。…` | 段落 | 長文（MAXLEN 超） |
| `項目`（表のヘッダ・罫線つき） | （表の一部） | (2) で表に吸収／is_ruled で除外 |

#### 見出しレベルの決定

章番号のドット数から決める。番号が無ければ `##`（レベル 2）。

```python
def _heading_level(text):
    m = _NUM_RE.match(text)
    if m:
        return min(2 + m.group(1).count("."), 6)  # "1"→2(##), "1.1"→3(###)
    return 2
```

#### KeyValue の判定

罫線のない 2 列で、**左列が一様に短いラベル**（15 文字以下・改行なし）なら、表ではなく「項目＝値」の組とみなして箇条書きにする。

```python
def _looks_like_kv(sheet, reg):
    labels = [sheet.text_at(r, reg.c0) for r in range(reg.r0, reg.r1 + 1)]
    labels = [t for t in labels if t]
    return bool(labels) and all(len(t) <= 15 and "\n" not in t for t in labels)
```

### ⑤ 生成 — `render`

Block を Markdown に描く。**表は内部結合が無ければ GFM、あれば HTML**（colspan/rowspan）。GFM では**全空の行・列（スペーサー）を除去**して自然な表にする。

```python
# 結合あり → HTML（span を領域内にクランプ、アンカー画像は <img> で差し込む）
content = escape(sheet.text_at(r, c)).replace("\n", "<br>")
for name in image_map.get((r, c), []):
    content += f'<br><img src="{image_dirname}/{name}" alt="{_stem(name)}">'
cells.append(f"<td{attrs}>{content}</td>")
```

| Block | 出力 |
|---|---|
| `Heading(level, text)` | `## text` |
| `Paragraph(text)` | 本文（セル内改行は行末2スペース改行で保持） |
| `Table`（結合なし） | GFM テーブル（空行・空列は除去） |
| `Table`（結合あり） | HTML テーブル（colspan/rowspan、画像は `<img>`） |
| `KeyValue` | `- **ラベル**: 値` の箇条書き |
| `Figure` | `![](images/...)` |

### 検証 — `verify`

元セルの非空テキストが出力 Markdown に全て現れるかを照合する。空白・記号・HTML エスケープの差を正規化して部分一致を見る。欠落があれば `-v` で警告表示する（ルールベースの取りこぼしも、将来の LLM の欠落・改変も同じ網で検出）。

---

## 4. 入力 Excel → 出力 Markdown の具体例

`make_sample.py` が生成する 3 シートで、3 つの代表パターンを示す。

### 例1: 自然な設計書（見出し・段落・表・図）

セクション間が空行で区切られた、素直な画面設計書。

**Excel（シート「画面設計 ログイン」）**

```
     B               C            D      E
2  [ 1. ログイン画面                        ]   ← B2:E2 結合
3  [ 本画面はユーザ認証を行う。ID と…         ]   ← B3:E3 結合（長文）
4  (空行)
5  [ 1.1 入力項目                            ]   ← B5:E5 結合
6    項目            種別         必須    備考    ← 罫線つき表 B6:E8
7    ID              テキスト     ○      半角英数
8    PW              パスワード   ○      8文字以上
9  (空行)
10 [ 1.2 画面イメージ                        ]
11   〔画像〕                                    ← B11 にアンカー
```

**出力 Markdown**

```markdown
# 画面設計 ログイン

## 1. ログイン画面

本画面はユーザ認証を行う。ID とパスワードを入力し、ログインボタンを押下する。

### 1.1 入力項目

| 項目 | 種別 | 必須 | 備考 |
| --- | --- | --- | --- |
| ID | テキスト | ○ | 半角英数 |
| PW | パスワード | ○ | 8文字以上 |

### 1.2 画面イメージ

![画面設計_ログイン_img1](images/画面設計_ログイン_img1.png)
```

- 全幅結合の短文「1. ログイン画面」→ 章番号で `##`、「1.1」→ `###`。
- 全幅結合の長文 → 段落。
- 罫線つき格子 → GFM テーブル。
- アンカーされた画像 → `![]()`。

### 例2: 罫線つきの表（スペーサー列の保護と除去）

「方眼紙」で表の列間に**空のスペーサー列（D列、罫線あり）**が入っているケース。罫線で表全体が保護されて粉砕されず、出力時に空列が除去される。

**Excel（シート「詳細設計」）**

```
     B              C      D(空・罫線) E
2  [ 2. 詳細設計                          ]   ← 結合・太字
3  (空行)
4    項目           値      ░          単位   ← B4:E7 全体に罫線
5    最大同時接続数  100     ░          件
6    タイムアウト    30      ░          秒
7    リトライ回数    3       ░          回
```

**出力 Markdown**

```markdown
# 詳細設計

## 2. 詳細設計

| 項目 | 値 | 単位 |
| --- | --- | --- |
| 最大同時接続数 | 100 | 件 |
| タイムアウト | 30 | 秒 |
| リトライ回数 | 3 | 回 |
```

- D 列は空だが罫線があるため `is_ruled` で「非空」扱いになり、**XY-cut が表を縦に切り刻まない**。
- 描画時に**全空の列として除去**され、3 列のクリーンな表になる。

### 例3: 密なフォーム（結合主体 → HTML テーブル）

内部に空行・空列・全幅バナーが無い密なフォーム。1 ブロックのままとなり、結合・改行・リンク・画像をすべて保持した HTML テーブルになる。

**Excel（シート「機能要件」）**

```
     A          B            C        D
1  [ 機能要件一覧                          ]   ← A1:D1 結合
2    項目       [ 内容                     ]   ← B2:D2 結合
3    機能名     [ ログイン                 ]   ← B3:D3 結合・セル内改行
              [ (認証)                   ]
4    優先度     高
…
7    参照       公式サイト（ハイパーリンク）
8  [区分]       画面         〔画像〕         ← A8:A9 縦結合 / C8 に画像
9              帳票
```

**出力 Markdown**

```markdown
# 機能要件

## 機能要件一覧

<table>
<tr><td>項目</td><td colspan="3">内容</td></tr>
<tr><td>機能名</td><td colspan="3">ログイン<br>(認証)</td></tr>
<tr><td>優先度</td><td>高</td><td></td><td></td></tr>
...
<tr><td>参照</td><td>公式サイト (https://example.com/spec)</td><td></td><td></td></tr>
<tr><td rowspan="2">区分</td><td>画面</td><td><img src="images/機能要件_img1.png" alt="機能要件_img1"></td><td></td></tr>
<tr><td>帳票</td><td></td><td></td></tr>
</table>
```

- 全幅結合の見出し「機能要件一覧」→ `##`。
- 残りは結合主体のため HTML テーブル：`colspan`/`rowspan`、セル内改行 `<br>`、ハイパーリンク `(URL)` 併記、アンカー画像 `<img>` をすべて保持。

---

## 5. セットアップ

```bash
python3 -m venv .venv
./.venv/bin/pip install -r requirements.txt   # openpyxl（サンプル生成には Pillow も）
```

## 6. 使い方

### CLI

```bash
# 自然な設計書 Markdown（既定 = v2 構造復元）
python -m xlsx2md path/to/spec.xlsx -o output

# ②分割の可視化のみ（各ブロックの範囲・暫定ロール・テキスト先頭をダンプ）
python -m xlsx2md spec.xlsx --segment

# v1 忠実モード（1 シート = 1 HTML テーブル）
python -m xlsx2md spec.xlsx -o out_faithful --faithful
```

| オプション | 説明 |
|---|---|
| `-o, --out-dir` | 出力先（既定: `output`） |
| `--segment` | ②分割の可視化のみ |
| `--faithful` | v1 忠実 HTML テーブル出力 |
| `--no-hidden` | 隠し行/列を出力しない |
| `--no-formula-fallback` | 数式キャッシュが無い時に式文字列へフォールバックしない |
| `-v, --verbose` | 詳細ログ（逆バティム検証の欠落警告を含む） |

出力構成：

```
output/
  spec/
    01_シート名.md
    02_シート名.md
    images/
      シート名_img1.png
```

### ライブラリ

```python
from xlsx2md import Options
from xlsx2md.pipeline import convert_structured

paths = convert_structured("spec.xlsx", "output", Options())
print(paths)  # 生成した .md のパス一覧
```

### 動作確認

```bash
./.venv/bin/python make_sample.py          # sample.xlsx を生成（3 シート）
./.venv/bin/python -m xlsx2md sample.xlsx -o output -v
```

---

## 7. 既知の限界

1. **全面が罫線の「表のみ」シート**（罫線被覆率 ≈ 1.0）は装飾と誤判定され、表保護が外れる余地がある。恒久対策は「表矩形の明示検出」への移行。
2. **罫線なし（整列のみ）の表**は XY-cut 依存のため、内部に空列があると崩れうる。
3. 見出しは保守的判定のため、**無装飾・非全幅・番号なしの見出し**は段落に降格しうる。
4. 左右2段組みなど複雑なレイアウトの分割精度は実データ依存。

## 8. スコープ外（初版）

グラフ、図形・テキストボックス・矢印などの描画オブジェクト、セルのコメント、見た目書式（色・太字・罫線）の出力。詳細は要件定義書「4. スコープ外」を参照。
