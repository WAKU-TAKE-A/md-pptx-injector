# md-pptx-injector.py

Markdownファイルをテンプレートベースで PowerPoint プレゼンテーション（.pptx）に変換するツールです。カスタムレイアウト・プレースホルダー・書式対応・目次自動生成などを備えています。

## 目次

- [クイックスタート](#クイックスタート)
- [機能一覧](#機能一覧)
- [インストール](#インストール)
- [使い方](#使い方)
- [パス解決ルール](#パス解決ルール)
- [Markdown 記法](#markdown-記法)
- [レイアウトとプレースホルダー](#レイアウトとプレースホルダー)
- [高度な機能](#高度な機能)
- [トラブルシューティング](#トラブルシューティング)

---

## クイックスタート

```bash
# 依存パッケージのインストール
pip install python-pptx Pillow

# 基本的な使い方
python md-pptx-injector.py input.md output.pptx --template template.pptx

# デバッグ用詳細ログあり
python md-pptx-injector.py input.md output.pptx --template template.pptx -v
```

---

## 機能一覧

- Markdown → PowerPoint 変換（テンプレート使用）
- HTMLコメントによるカスタムレイアウト指定
- プレースホルダー名を指定したコンテンツ配置
- `<!-- new_page -->` による明示的なページ区切り
- `<b>`, `<i>`, `<u>` タグによるインライン書式（太字・斜体・下線）
- 箇条書き（`-` `*` `+`）および番号付きリスト（最大3レベル）
- コードブロック（``` ``` ``` 囲み → グレー背景テキストボックス）
- Markdown テーブル → PowerPoint テーブル（列幅・アラインメント制御付き）
- 画像挿入（アスペクト比保持のコンテインフィット、キャプション対応）
- YAML front matter によるタイトルスライド生成
- `toc: true` による目次スライド自動生成（スライド間ハイパーリンク付き）
- PyInstaller（exe 化）対応

---

## インストール

**必要環境**

- Python 3.9 以降

**依存パッケージ**

```bash
pip install python-pptx Pillow
```

> `Pillow` は画像挿入に使用します。不要な場合はインストールしなくても動作しますが、画像は無視されます。

---

## 使い方

### コマンドライン

```bash
python md-pptx-injector.py <src> <dst> [オプション]
```

| 引数・オプション | 説明 |
|---|---|
| `src` | 入力 Markdown ファイル |
| `dst` | 出力 PowerPoint ファイル |
| `--template PATH` | テンプレート .pptx ファイル（デフォルト: `template.pptx`） |
| `-v, --verbose` | デバッグログを標準出力に表示 |

**例**

```bash
# テンプレートをデフォルト（template.pptx）から使用
python md-pptx-injector.py slides.md presentation.pptx

# テンプレートを明示指定
python md-pptx-injector.py slides.md presentation.pptx --template custom.pptx

# デバッグモード
python md-pptx-injector.py slides.md presentation.pptx -v
```

---

## パス解決ルール

### アプリケーション基準ディレクトリ

| 実行方法 | 基準ディレクトリ |
|---|---|
| スクリプト実行 | `.py` ファイルが置かれているディレクトリ |
| exe 実行（PyInstaller） | `.exe` ファイルが置かれているディレクトリ |

### src・dst・--template のパス解決順序

1. **絶対パス** → そのまま使用
2. **カレントディレクトリに存在するパス** → カレントディレクトリ基準で解決（優先）
3. **それ以外の相対パス** → アプリケーション基準ディレクトリから解決

### 画像ファイルのパス解決順序

1. Markdown ファイルのあるディレクトリ（最優先）
2. アプリケーション基準ディレクトリ
3. カレントディレクトリ

---

## Markdown 記法

### ページ区切り（スライド境界）

以下のいずれかがスライドの区切りになります。

**見出し（# / ## / ###）**

```markdown
## スライドA

コンテンツA

## スライドB

コンテンツB
```

**明示的なページ区切りコメント**

```markdown
<!-- new_page -->
```

レイアウトを同時に指定することも可能です（後述）。

```markdown
<!-- new_page="Two Content" -->
```

> **注意**: ドキュメント先頭の YAML front matter 内の `---` はページ区切りとして扱われません。

---

### テキスト書式

インライン書式は `<b>`, `<i>`, `<u>` のHTMLタグで指定します。ネスト可能です。

```markdown
<b>太字</b>
<i>斜体</i>
<u>下線</u>
<b><i>太字＋斜体</i></b>
この行の<b>一部だけ</b>太字にすることも<i>できます</i>。
```

> タグの開き忘れや順序の不一致がある場合は書式なしのテキストとして出力されます（`-v` オプションで警告表示）。

---

### 見出し

#### スライドレイアウトを決定する見出し（`#` `##` `###`）

| Markdown | 選択されるレイアウト |
|---|---|
| `# タイトル` | `Title Slide` |
| `## セクション` | `Section Header` |
| `### コンテンツ` | `Title and Content` |

これらの見出しはスライドのタイトルプレースホルダーに書き込まれます。

#### スライド内見出し（`####` `#####`）

スライド内のコンテンツ領域に配置され、後続テキストの **ベースレベル** を設定します。

```markdown
#### Level 0 の見出し
ここから level 0 のテキスト

##### Level 1 の見出し
ここから level 1 のテキスト
- ここは level 1 の箇条書き
  - ここは level 2 の箇条書き
```

| 見出しレベル | ベースレベル |
|---|---|
| `####` | 0 |
| `#####` | 1 |

---

### 箇条書き・番号付きリスト

```markdown
- レベル0の箇条書き
  - レベル1（スペース2個インデント）
    - レベル2（スペース4個インデント）

* マーカーは * でも可
+ マーカーは + でも可

1. 番号付きリスト（番号はそのままテキストとして出力）
2. 2番目
```

- インデント: 2スペース = 1レベル（`indent:` で変更可能、後述）
- 最大レベル: 2（超過分はレベル2にクリップ）
- 番号付きリストの番号は自動採番ではなく、テキストとして出力されます

---

### コードブロック

````markdown
```
def hello():
    print("Hello, World!")
```
````

スライド上にグレー背景・Courier New フォントのテキストボックスとして配置されます。

---

### 空行

空行は PowerPoint 内で空の段落（縦の余白）として挿入されます。

---

## レイアウトとプレースホルダー

### レイアウト選択の優先順位

1. `<!-- layout="..." -->` コメント（最優先）
2. `<!-- new_page="..." -->` コメント
3. 見出しレベル（`#` / `##` / `###`）
4. YAML front matter の有無（見出しなし → `Title Slide`）

指定されたレイアウトがテンプレートに存在しない場合は自動的に `Title and Content` → テンプレートの最初のレイアウト、の順でフォールバックします（エラーにはなりません）。

---

### レイアウト指定コメント

```markdown
<!-- layout="Two Content" -->
## スライドタイトル
```

> **注意**: `<!-- layout="..." -->` の直後の行は `#`/`##`/`###` の見出しである必要があります。見出しが続かない場合、このコメントは無視されます。

---

### タイトルスライド

**パターンA: YAML front matter**

```markdown
---
title: プレゼンテーションタイトル
subtitle: サブタイトル
author: 著者名
toc: true
indent: 2
---
```

**パターンB: Markdown 見出し**（パターンAより優先）

```markdown
# プレゼンテーションタイトル
subtitle: サブタイトル
author: 著者名
```

`subtitle` と `author` は改行区切りでサブタイトルプレースホルダーに挿入されます。

**front matter のキー**

| キー | 内容 |
|---|---|
| `title` | タイトル |
| `subtitle` | サブタイトル |
| `author` | 著者名 |
| `toc` | `true` にすると末尾に目次スライドを自動生成 |
| `indent` | 箇条書きのインデント幅（スペース数、デフォルト: `2`） |

---

### カスタムプレースホルダー

テンプレート内のシェイプ名を指定して、コンテンツを直接流し込みます。

```markdown
<!-- placeholder="LeftBox" -->
ここの内容は "LeftBox" という名前のシェイプに書き込まれます。
（空行が来るまでキャプチャ）

<!-- placeholder="RightBox" -->
ここは "RightBox" に書き込まれます。

<!-- placeholder="LeftBox" -->
同じプレースホルダーに再度書くと、空行区切りで追記されます。
```

**ルール**
- 空行までの内容をキャプチャします
- 同一プレースホルダーへの複数ブロックは、空行区切りで**追記**されます
- 指定したプレースホルダーが見つからない場合、内容はレスキューコンテンツとして body プレースホルダーに流れます

---

### レスキューコンテンツ

プレースホルダー指定のないテキストは、自動的に body プレースホルダーに**追記**されます。

```markdown
<!-- layout="Two Content" -->
## スライドタイトル

<!-- placeholder="LeftBox" -->
左ボックスの内容

このテキストはプレースホルダー指定なし → body プレースホルダーに追記されます。
```

**レスキューの条件**
- `Title Slide` レイアウト以外のスライドのみ
- 空行のみの場合はスキップ
- 明示的に使用されたプレースホルダーは対象外

---

## 高度な機能

### テーブル

```markdown
<!-- placeholder="TableArea" -->
| 左揃え | 中央揃え | 右揃え |
|:-------|:--------:|-------:|
| Alpha  | Bravo    | 1      |
| Charlie| Delta    | 2      |
```

- プレースホルダーの位置・サイズでテーブルを生成します
- 列幅はセパレータ行のダッシュ数の比率で決まります
- アラインメント: `:---` 左揃え、`:--:` 中央揃え、`---:` 右揃え
- プレースホルダー指定なしのテーブルは無視されます

---

### 画像

```markdown
<!-- placeholder="ImageArea" -->
![キャプションテキスト](image.jpg)
```

- プレースホルダーの枠内にコンテインフィット（アスペクト比保持）で挿入
- `![キャプション]` に文字がある場合、`ImageArea_caption` という名前のシェイプがあればそこにキャプションが書き込まれます
- プレースホルダー指定なしの画像は無視されます

---

### 目次スライド

YAML front matter に `toc: true` を記述すると、最終スライドとして目次スライドを自動生成します。

```markdown
---
title: プレゼンテーション
toc: true
---
```

- `##`（Section Header）と `###`（Title and Content）の見出しが目次エントリになります
- 各エントリはスライドへのハイパーリンク付きで出力されます

---

## トラブルシューティング

### ログレベル

| レベル | 通常時 | `-v` 時 | 内容 |
|---|---|---|---|
| `DEBUG` | ❌ | ✅ | シェイプ情報・プレースホルダー解決の詳細 |
| `INFO` | ❌ | ✅ | 進捗メッセージ |
| `WARNING` | ✅ | ✅ | 非致命的な問題（レイアウト未検出・Pillowなしなど） |
| `ERROR` | ✅ | ✅ | 致命的エラー（ファイル未検出・保存失敗など） |

---

### よくある問題と対処

**テンプレートが見つからない**

```
File not found.
```

`--template` で正しいパスを指定するか、スクリプトと同じディレクトリに `template.pptx` を置いてください。

**レイアウトが見つからない**

```
WARNING: Slide 2: layout 'MyLayout' not found. Falling back to auto.
```

PowerPoint テンプレートのスライドレイアウト名と Markdown の指定が完全一致（大文字小文字を含む）しているか確認してください。

**プレースホルダーが見つからない**

```
[page 2] placeholder "Content" NOT FOUND -> rescuing content
```

PowerPoint の「選択ウィンドウ」でシェイプ名を確認し、Markdown の指定と一致させてください。

**画像が挿入されない**

```
WARNING: Image not found: logo.png
```

Markdown ファイルと同じディレクトリに画像を置くか、絶対パスで指定してください。`-v` オプションで検索パスの詳細が表示されます。

**書式タグが効かない**

```
WARNING: Slide 3: Unclosed tag in 'テキスト<b>'. Skipping formatting.
```

`<b>`, `<i>`, `<u>` の開きタグと閉じタグが対応しているか確認してください。

---

### デバッグ手順

```bash
# 1. 詳細ログで実行
python md-pptx-injector.py input.md output.pptx -v

# 2. スライドのシェイプ一覧を確認
# [shapes on slide]
#   - #0: name='Title 1', is_placeholder=True, has_text_frame=True
#   - #1: name='Content 2', is_placeholder=True, has_text_frame=True

# 3. プレースホルダー解決を確認
# [page 2] placeholder "Content" found: actual_name='Content 2'

# 4. 画像パスを確認
# [image] inserted '/full/path/to/image.jpg' alt='キャプション'
```
