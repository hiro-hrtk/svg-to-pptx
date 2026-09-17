---
name: svg-to-pptx
description: >
  SVGファイルをPowerPointのネイティブオブジェクト(図形・コネクタ・テキストボックス)に変換する。
  「SVGをPPTXに変換して」「このSVGをスライドにして」「変換パイプラインを回して」
  などの依頼で起動する。input/またはinput/test/配下のSVGを対象とし、
  build_pptx.py → qa_capture.py のフルパイプラインを実行してQA PNGで結果を確認する。
allowed-tools:
  - Bash
  - Read
  - Write
  - Edit
  - Glob
  - AskUserQuestion
---

# SVG → PPTX 変換スキル

このスキルを読んだら、以下のフローを**そのまま実行**する。推測で判断せず、不明な点は `AskUserQuestion` で確認すること。

---

## Step 0: プロジェクトルートを確定する

```
PROJECT_ROOT = "C:\Users\mfutu\OneDrive\ドキュメント\claude-pjt\svg-to-pptx"
```

以下のパスはすべてこのルートからの相対パスとして解釈する。
Python実行環境は **uv** で管理されている(pip/venvではない)。

**重要**: このマシンではアプリケーション制御ポリシーによりuv venv内の`python.exe`ランチャーの
実行がブロックされる。そのため `uv run python ...` は使わず、必ず `code\run.ps1` 経由で実行する
(詳細: CLAUDE.md「実行環境」、docs/DECISIONS.md ADR-020)。

---

## Step 1: 入力SVGを特定する

### 1-1. ユーザーの依頼からSVGを特定する

ユーザーがSVGパスを明示した場合はそれを使う。明示されていない場合は `input/` 配下のSVGを列挙してユーザーに選ばせる。

```powershell
Get-ChildItem -Recurse "<PROJECT_ROOT>\input" -Filter *.svg
```

### 1-2. 単一SVGかペア形式かを判断する

| ケース | 判断基準 | no-text引数 | text-only引数 |
|--------|---------|-------------|---------------|
| **単一SVG** | ファイルが1つ、またはユーザーが「このSVG」と1ファイルを指定 | 同じSVGファイル | 同じSVGファイル |
| **ペア形式** | `_no_text` と `_text_only` の組み合わせが存在する | `*_no_text.svg` | `*_text_only.svg` |

ペア形式の自動検出ロジック:
- 指定SVGのステム名から `_text_only` / `_no_text` を除いたベース名を取得
- 同じディレクトリに `<base>_no_text.svg` と `<base>_text_only.svg` が両方あればペアと判断

不明な場合は AskUserQuestion で確認する。

### 1-3. 未対応のSVG機能が無いか事前チェックする

変換前に対象SVGを軽くgrepし、以下が含まれていないか確認する(`docs/CONCEPT.md`の「既知の制約」参照)。
含まれている場合は**変換を止めずに実行した上で**、QA時に該当箇所を重点的に確認し、報告に含める
(事前に気づけるほど手戻りが減る)。

| 属性・要素 | 対応状況 |
|-----------|---------|
| `transform="translate(x y)"` | **対応済み**(`<g>`の累積オフセットとして座標に反映) |
| `transform`のscale/rotate/matrix | 非対応。座標がズレる可能性が高い。見つけたらユーザーへ警告する |
| 祖先`<g>`からのfill/stroke/opacity等の継承 | **対応済み**(`<g fill="#fff">`のように子に共通スタイルを持たせるパターン) |
| `<use>` / `<symbol>` | 非対応。該当要素は描画されない |
| `<clipPath>` / `<mask>` | 非対応。無視される |
| `<filter>`(drop-shadow等) | 非対応。影などは再現されない |
| `<style>`内のクラス/型セレクタ(`.name{...}`/`text{...}`) | **対応済み**(単純なセレクタのみ。結合子・疑似クラス・idセレクタは非対応) |
| `marker-end`(SVG標準の矢じり) | **対応済み** |
| 閉じた塗りつぶし`<path>`(角丸矩形等をpathで描画したヘッダー等) | **対応済み**(コンテナ候補として認識、ADR-014) |
| `stroke-linecap`/`stroke-linejoin` | 非対応(PPTX側デフォルトのまま) |
| `currentColor` | 非対応(色が付かない) |
| base64埋め込み`<image>` | 対応済み。ただし画像の中身(ピクセル)は解釈しないため、画像に既にラベル文字が焼き込まれていて、かつ同じ文言の`<text>`も別途重ねられている場合は表示が二重になる |

`transform`にscale/rotate/matrixが使われている場合は致命的(座標が崩れる)なので、AskUserQuestionで
「transform(scale/rotate/matrix)が検出されました。現状非対応のため座標がズレる可能性があります。続行しますか」
と確認してから進める(`translate`のみなら対応済みなので確認不要)。

---

## Step 2: セッション名を決める

セッション名 = 変換対象SVGのステム名(`_no_text` / `_text_only` サフィックスは除去)

例: `jira_agent_cicd_architecture_text_only.svg` → セッション名: `jira_agent_cicd_architecture`

ユーザーが別の名前を希望する場合は受け付ける。

---

## Step 3: build_pptx.py を実行する

```powershell
Set-Location "<PROJECT_ROOT>"
.\code\run.ps1 code\build_pptx.py `
  --session <session_name> `
  --no-text <no_text_svg_path> `
  --text-only <text_only_svg_path>
```

### 実行後の確認

出力先: `output/<session_name>/`

以下のファイルが生成されていることを確認する:
- `<session_name>_native.pptx` — 生成されたPPTX
- `<session_name>_mapping.json` — テキスト→図形マッピング表(デバッグ用)

エラーが発生した場合はスタックトレースを読み、原因を特定して修正を試みる。
修正できない場合はユーザーに報告して停止する。

---

## Step 4: qa_capture.py でPNG出力する

```powershell
Set-Location "<PROJECT_ROOT>"
.\code\run.ps1 code\qa_capture.py "output/<session_name>/<session_name>_native.pptx"
```

出力先: `output/<session_name>/<session_name>_native_qa_slide1.png`

**このPNGファイルを Read ツールで読み込み、Claude が目視確認を行う。**
問題箇所の拡大確認が必要な場合は、PILで該当領域をcrop・拡大してから再度Readで確認するとよい。

---

## Step 5: QA確認を行う

生成されたPNGを Read で読み込み、`docs/TEST_CASES.md` のチェックリストに沿って確認する。特に:

- [ ] テキストがボックスから大きくはみ出していないか
- [ ] 太字ヘッダーや英単語混じりのラベルが単語の途中で改行されていないか(ADR-008/009/013)
- [ ] アイコン画像・アイコン形状下のラベルが幅に引きずられて不必要に折り返されていないか(ADR-012/014)
- [ ] 図形がPowerPointのデフォルト色(青)で塗りつぶされたまま残っていないか(ADR-011)
- [ ] 浮遊テキスト(矢印ラベル等)に白背景が敷かれているか、孤立テキストに不要な白背景が付いていないか(ADR-010)
- [ ] 同じボックス内でフォントサイズが異なるテキスト同士が誤って結合され、片方が拡大/縮小されていないか(ADR-016)
- [ ] `<g fill="...">`のように親でまとめて指定されたアイコン等が、色が付かず消えていないか(ADR-018)
- [ ] Step 1-3で検出した未対応機能がある場合、その周辺の見た目が大きく崩れていないか

### 問題を発見した場合
1. `output/<session_name>/<session_name>_mapping.json` を Read して、テキストの所属マッピングを確認する
2. 問題箇所のテキストIDやシェイプIDを特定する
3. 原因を `docs/CONCEPT.md` の設計原則・`docs/DECISIONS.md` のADRと照らし合わせて分析する
4. 該当モジュール(`code/svg_elements.py` / `code/mapping.py` / `code/concat_text.py` / `code/fit_text.py` / `code/build_pptx.py`)を修正する
5. Step 3 に戻る

---

## Step 6: 結果を報告し、ドキュメントを更新する

### 成功した場合

ユーザーに以下を報告する:
- 生成PPTXのパス
- QA PNG のパス(目視確認結果のサマリを含む)
- 特筆すべき問題点(軽微な課題があれば)

### ドキュメント更新の提案

**docs/CHANGELOG.md の更新**(コードに変更を加えた場合):
- `## [Unreleased]` セクションの `### Fixed` または `### Changed` に変更内容を追記

**docs/TEST_CASES.md の更新**(新しいSVGをテストした場合):
- 「テストSVG一覧」のステータスを `✅ 検証済み` に変更
- 「過去のQA結果ログ」に行を追記:
  ```
  | YYYY-MM-DD | <session_name> | <svg_name> | ✅ Pass | <備考> |
  ```

**docs/DECISIONS.md の更新**(新たな設計判断・トレードオフが生まれた場合):
- 連番でADRを追記する

ユーザーに更新内容を確認してから書き込む。

---

## エラーハンドリング早見表

| エラー | 考えられる原因 | 対処 |
|--------|---------------|------|
| `アプリケーション制御ポリシーによってこのファイルがブロックされました` / `uv run`が失敗する | venv内`python.exe`ランチャーがブロックされる(ADR-020) | `uv run`を使わず`code\run.ps1`経由で実行する |
| `ModuleNotFoundError: pptx`/`PIL`/`yaml`/`win32com` | 依存未インストール、またはvenvの場所を見失っている | `$env:UV_PROJECT_ENVIRONMENT`が設定されているか確認し`uv sync`を再実行 |
| `FileNotFoundError: meiryo.ttc` | フォントパスが違う | `code/config.yaml` の `font.font_file` を確認 |
| SVGがパースできない | SVG形式の問題 | `code/svg_elements.py` の対応要素を確認。非対応要素はスキップされる |
| `ValueError: could not convert string to float` | CSSの`38px`のような単位付き数値が想定外だった | `svg_elements._num()` を使っているか確認(単位を剥がしてfloat化するヘルパー) |
| PPTXは生成されるが見た目がおかしい | テキスト配置・マッピングの問題 | Step 5 のQA手順でデバッグ |
| `qa_capture.py`が失敗する | PowerPointが未インストール/COM権限 | PowerPointのインストール状況を確認。それでも失敗する場合はユーザーに報告 |
| `PermissionError`でPPTX保存に失敗する | 前回実行時のPowerPointプロセスが残ってファイルをロックしている | `Get-Process POWERPNT` で確認し、`qa_capture.py`が自分で起動したインスタンスなら`Stop-Process`で終了してから再実行(通常は`qa_capture.py`側でQuit()される。ユーザーが別途PowerPointを開いている場合は誤って閉じないよう注意) |
| 図形がPowerPointのデフォルト色(青)で塗りつぶされる | fill/lineの明示的クリア漏れ(ADR-011参照) | `apply_shape_style()`で`has_fill_attr=True`の場合は必ず`fill.solid()`か`fill.background()`のどちらかを呼んでいるか確認 |
| 英単語混じりのラベルが単語の途中で改行される | 折り返しが文字単位になっている(ADR-013) | `fit_text._wrap_paragraph()`のトークン化ロジックを確認 |
| アイコン画像下のラベルが不必要に折り返される | 小さい`<image>`がマッピングのコンテナ候補になっている(ADR-012) | `config.yaml`の`mapping.image_container_min_px`を確認・調整 |
