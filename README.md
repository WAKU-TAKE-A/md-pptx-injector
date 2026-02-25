# md-pptx-injector

Markdown ファイルを PowerPoint プレゼンテーション (.pptx) に変換するPythonスクリプトです。  
テンプレートのレイアウト・プレースホルダーを活用しながら、Markdownで書いたコンテンツをスライドに流し込みます。

---

## 必要環境

- Python 3.10 以上
- python-pptx
- Pillow（画像挿入を使用する場合）

```bash
pip install python-pptx Pillow
```

---

## 使い方

```bash
python md-pptx-injector.py src.md dst.pptx --template ref.pptx
python md-pptx-injector.py src.md dst.pptx --template ref.pptx -v
```

| 引数 | 説明 |
|------|------|
| `src.md` | 入力Markdownファイル |
| `dst.pptx` | 出力PowerPointファイル |
| `--template ref.pptx` | テンプレートPPTXファイル（デフォルト: `template.pptx`） |
| `-v` | 詳細ログを出力（デバッグ用） |

### パス解決ルール

- 絶対パス → そのまま使用
- `./` または `.\` で始まる → カレントディレクトリ基準
- それ以外 → スクリプト（または exe）のディレクトリ基準

---

## Markdown書式

### ページ分割

以下のいずれかでページが分割されます。

| 書式 | 動作 |
|------|------|
| `# 見出し` / `## 見出し` / `### 見出し` | ページ分割 |
| `<!-- new_page -->` | 無条件ページ分割 |
| `<!-- new_page="レイアウト名" -->` | 指定レイアウトで無条件ページ分割 |
| `<!-- layout="レイアウト名" -->` + 直後の `#/##/###` | ページ分割（layoutタグは見出しの直前に必要） |

> **注意:**
> - `---`（水平線）はYAML front matter以外では無視されます。
> - YAML front matterが存在するページでは `#` 見出しは同一ページに含まれます（空のタイトルスライドは生成されません）。

### YAML Front Matter

ドキュメント先頭に記述します。

```markdown
---
title: タイトル
subtitle: サブタイトル
author: 著者名
toc: true
toc_title: 目次
indent: 2
font_size_l0: 18
font_size_l1: 16
font_size_l2: 14
font_size_l3: 12
font_size_l4: 10
---
```

| キー | 説明 |
|------|------|
| `title` | タイトルスライドのタイトル |
| `subtitle` | タイトルスライドのサブタイトル |
| `author` | 著者名（サブタイトルに追記） |
| `toc: true` | 目次スライドを末尾に生成 |
| `toc_title` | 目次スライドのタイトル（デフォルト: `Table of Contents`） |
| `indent` | リストのインデント幅（スペース数、デフォルト: 2） |
| `font_size_l0` 〜 `font_size_l4` | 各レベルの文字サイズ（pt）、文書全体に適用 |

### レイアウト指定

```markdown
<!-- layout="sample02" -->
### ページタイトル
```

- `<!-- layout="..." -->` の**直後の行**が `#/##/###` でなければ無視されます（`-v` で警告表示）
- レイアウト名はテンプレートPPTX内のスライドレイアウト名と一致させてください

レイアウトに文字サイズを指定することもできます（ページ単位で有効）：

```markdown
<!-- layout="sample02" font_size_l0=16 font_size_l1=12 -->
### ページタイトル
```

### プレースホルダー指定

テンプレートのシェイプ名を指定してコンテンツを流し込みます。

```markdown
<!-- placeholder="holder01" -->
#### ■ 内容1
* リスト1
* リスト2
  * リスト2-1

<!-- placeholder="holder02" -->
#### ■ 内容2
本文テキスト
```

プレースホルダーに文字サイズを指定することもできます（そのホルダー内のみ有効）：

```markdown
<!-- placeholder="holder01" font_size_l0=14 font_size_l2=10 -->
#### ■ 内容
* リスト
```

プレースホルダーが見つからない場合は本文領域に配置されます（レスキュー）。

### 文字サイズの優先順位

```
YAML（文書全体） < layout（ページ） < placeholder（ホルダー内）
```

指定のないレベルはテンプレートのマスター設定に従います。

### インライン書式

```markdown
<b>太字</b>
<i>斜体</i>
<u>下線</u>
<b><i>太字+斜体</i></b>
```

タグのミスマッチや未閉じタグはプレーンテキストにフォールバックします（`-v` で警告表示）。

### 見出し（スライド内）

`####` 以降はページ分割を起こさず、スライド内の段落として扱われます。

```markdown
#### ■ セクション見出し    → level=0 の段落
##### □ サブセクション     → level=1 の段落
###### □ サブサブセクション → level=2 の段落
```

### 箇条書き

```markdown
* リスト1
* リスト2
  * リスト2-1（indent 2スペース）
    * リスト2-1-1（indent 4スペース）
- 別記法
+ 別記法2
1. 番号付きリスト
```

最大5レベル（level 0〜4）まで対応。インデント幅は `indent` で設定。

### コードブロック

````markdown
```
コードをここに書く
```
````

灰色背景のテキストボックスとして挿入され、最背面（z-order の底）に配置されます。  
プレースホルダーのコンテンツが前面に描画されるため、コードブロックは背景として機能します。

### 画像

```markdown
<!-- placeholder="photo01" -->
![キャプション](image.png)
```

- `photo01` シェイプの位置・サイズにcontain-fitで挿入
- キャプションテキストが空でない場合、`photo01_caption` という名前のシェイプに書き込み

画像の検索順：
1. Markdownファイルと同じディレクトリ
2. スクリプト（exe）のディレクトリ
3. カレントディレクトリ

### テーブル

```markdown
<!-- placeholder="table01" -->
| 列1   | 列2   | 列3   |
|:------|:-----:|------:|
| A     | B     | C     |
```

- `|:---|` 左揃え / `|:---:|` 中央揃え / `|---:|` 右揃え
- セパレータ行のダッシュ数が列幅の比率になります（例: `|--|----:|` → 1:2 の幅で右揃え）

### 目次（TOC）

YAML に `toc: true` を指定すると、`##` / `###` 見出しを収集して文書末尾に目次スライドを生成します。  
各項目はスライドショーモードでクリックすると対象スライドにジャンプするハイパーリンク付きです。

```markdown
---
toc: true
toc_title: 目次
---
```

---

## サンプル

```markdown
---
title: サンプルプレゼン
subtitle: テスト用
author: 山田太郎
toc: true
toc_title: 目次
font_size_l0: 18
font_size_l1: 14
---

# タイトル
subtitle: サブタイトルをここに

## 第1章

章の説明文です。

<!-- layout="TwoContent" font_size_l0=16 -->
### ページ1-1

<!-- placeholder="holder01" font_size_l1=12 -->
#### ■ 左コンテンツ
* 項目A
* 項目B
  * 項目B-1

<!-- placeholder="holder02" -->
#### ■ 右コンテンツ
本文テキストです。

<!-- new_page -->

空白ページ
```

---

## レイアウト自動判定

`<!-- layout -->` や `<!-- new_page="..." -->` で明示しない場合、見出しレベルで自動判定されます。

| 見出しレベル | 適用レイアウト |
|-------------|---------------|
| `#` またはYAML front matterあり | `Title Slide` |
| `##` | `Section Header` |
| `###` またはその他 | `Title and Content` |

指定レイアウトが見つからない場合は `Title and Content` → テンプレート最初のレイアウトの順でフォールバックします。

---

## バージョン

現在のバージョン: **0.9.4.0**
