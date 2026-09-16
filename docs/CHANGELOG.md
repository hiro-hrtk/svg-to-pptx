# Changelog — svg-to-pptx

## [Unreleased]

### Added
- `code/svg_elements.py`: `<style>`内の単純なCSSクラスセレクタ(`.name{font-size:..;fill:..;}`等)の解決に対応
  (`get_prop`の優先順位: inline style > CSSクラス > 直接属性 > default)。
- `code/svg_elements.py`: SVG標準の`marker-end="url(#id)"`(`<defs><marker>`)に対応。線・pathの終点に
  向きを合わせた矢じりFreeformを追加描画する。
- `code/svg_elements.py`: CSS単位付き数値(`38px`等)を許容する`_num()`ヘルパーを追加。

### Fixed
- 小さい`<image>`(アイコン画像)がマッピングのコンテナ候補になり、直下のラベルが画像の狭い幅に
  引きずられて単語途中で折り返される問題(ADR-012、`mapping.image_container_min_px`を追加)。
- テキスト折り返しが1文字単位だったため、英単語混じりのラベル(例:「DuckDB + VSS」)が
  単語の途中で改行される問題(ADR-013、トークン単位の折り返しに変更)。
- `code/qa_capture.py`: `PowerPoint.Application`を`Quit()`していなかったため、実行のたびに
  隠れたPowerPointプロセスが残り続け、次回以降の`build_pptx.py`によるPPTX上書き保存が
  `PermissionError`で失敗する場合があった。既存インスタンスに接続した場合(`GetActiveObject`で
  取得できた場合)はQuit()せず、自分で新規起動した場合のみQuit()するよう修正。

### Changed
- なし

---

## [0.1.0] — 2026-09-13

消失したスキル・プロジェクト一式を `REF/` の復旧ドキュメントをもとに再構築。

### Added
- `code/svg_elements.py`: SVGパース層。`rect`/`circle`/`ellipse`/`line`/`polyline`/`polygon`/`path`/`image`/`text`に対応。
  `path`はM/L/H/V/C/Q/A/Z(絶対・相対)を折れ線サンプリングして近似(旧仕様は直線近似のみだったが、
  テストSVG Aのアイコン曲線に対応するためベジエ・楕円弧サンプリングを新規実装)。
- `code/mapping.py`: テキスト→図形マッピング層(内包判定+フォールバック、ADR-003)。
- `code/concat_text.py`: 隣接テキストのconcat層(縦近接+横近接、ADR-006)。
- `code/fit_text.py`: PIL実測メトリクスによる折り返し・フォントサイズ自動フィット層。
- `code/build_pptx.py`: 変換パイプライン本体・CLI(`--session`/`--no-text`/`--text-only`)。
- `code/qa_capture.py`: PowerPoint COM経由のQA用PNGエクスポート。
- `code/config.yaml`: フォント・レイアウト・マッピング閾値設定。
- `uv`ベースのプロジェクト構成(`pyproject.toml`)。依存: python-pptx / Pillow / pywin32 / PyYAML。
- テストSVG(input/test/、リポジトリには非同梱)を追加し、初回QAサイクルを実施(docs/TEST_CASES.md参照)。

### Fixed
- (再構築時の初回QAで発見・修正。詳細は docs/DECISIONS.md ADR-007〜011)
  - 背景全面矩形が全テキストを誤って「内包」してしまう問題(ADR-007)
  - 疑似ボールドの実測幅不足による意図しない折り返し・はみ出し(ADR-008)
  - 折り返しで複数行化した際のコンテナからのはみ出し(ADR-009)
  - 孤立テキストへの不要な白背景(ADR-010)
  - 開いたpath/polylineがデフォルトの青塗りで残る問題(ADR-011)
  - Connectorに`.fill`属性が無く例外になる問題
  - 1行ブロックの行box高さがline_spacing分膨らみ、隣接ブロックと重なる問題
