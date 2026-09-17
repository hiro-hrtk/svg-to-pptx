# アーキテクチャ設計書 — svg-to-pptx

## 概要

SVGファイルをPowerPointのネイティブオブジェクト(図形・コネクタ・テキストボックス)に変換するパイプライン。
PNG貼り付けではなく、PPTXで完全に編集可能なスライドを生成することが目標。

実行は `.\code\run.ps1 code\build_pptx.py ...` / `.\code\run.ps1 code\qa_capture.py ...`
(このマシンではvenvランチャーがアプリケーション制御ポリシーでブロックされるため`uv run`は使わない。
詳細はCLAUDE.md「実行環境」、ADR-020参照)。

---

## モジュール構成と責務

```
code/
├── svg_elements.py    [パース層]   SVG XML → ShapeElem / TextElem
├── mapping.py         [マッピング層] TextElem → 所属ShapeElem の対応付け
├── concat_text.py     [集約層]     隣接TextElem を TextBlock に結合
├── fit_text.py         [計測層]     PIL実測幅でのテキスト折り返し・サイズフィット
├── build_pptx.py      [出力層]     全体パイプライン + PPTX生成 + CLI
├── qa_capture.py       [QA層]       PowerPoint COM → PNG出力
├── run.ps1              [実行ラッパー] venvランチャーblock対策(ADR-020)
└── config.yaml         [設定]       閾値・フォント・レイアウト定数
```

### `svg_elements.py` — パース層

- 入力: no_text.svg / text_only.svg (どちらも同じ座標系を前提)
- 出力: `ShapeElem[]` / `TextElem[]`
- SVG属性とinline styleを統合して正規化(`get_prop`関数)
- 対応要素: `<rect>` `<circle>` `<ellipse>` `<line>` `<polyline>` `<polygon>` `<path>` `<image>` `<text>`(`<tspan>`込み)
- `<path>` の `d` 属性は `parse_path_d()` がトークナイズし、`M/L/H/V/C/Q/A/Z`(絶対/相対)を折れ線サンプリングする。
  三次/二次ベジエは `_cubic_bezier_points`/`_quad_bezier_points`、楕円弧は `_arc_to_points`
  (SVG仕様のcenter parameterization変換)でサンプル点列に変換する。
- `<g>` はフラット化して子孫を辿る(`_iter_recursive`)。`transform="translate(x y)"` は累積オフセットとして
  追跡し、`_translate_shape()`で子要素の座標に加算する(ADR-018)。scale/rotate/matrixは非対応。
- 祖先`<g>`の直接指定(fill/stroke/stroke-width/opacity/font-size/font-weight/font-family/text-anchor)を
  `_own_specified_props()`で蓄積し、子要素が自分で持たない場合のフォールバック(`_prop()`)として使う
  (SVGの継承ルール対応、ADR-018)。
- `fill="url(#id)"` は `_parse_defs()` で集めたgradient stopの平均色に解決する。
- `<style>`内の単純なクラスセレクタ(`.name{...}`)と型セレクタ(`text{...}`)を`_parse_style_rules()`で解決し、
  `get_prop()`が `inline style > CSSクラス > CSS型セレクタ > 直接属性 > default`の優先順位で値を返す。
- `marker-end="url(#id)"`を`_parse_markers()`で解決した`<marker>`定義と組み合わせ、線・pathの
  終点に矢じり形状のFreeformを追加描画する(`_build_marker_shape`/`_line_angle`)。
- `<text>` は同一要素内の`<tspan>`群を1つの `TextElem.lines`(複数行)にまとめる。**別々の`<text>`要素同士の
  結合は行わない**(それは`concat_text.py`の責務)。

### `mapping.py` — マッピング層

- 入力: `ShapeElem[]` + `TextElem[]` + `svg_size`(キャンバス全体のサイズ)
- 出力: `{text_id: shape_id}` + `contained_ids` (set)
- コンテナ候補は `kind in (rect, image)` に加え、閉じた塗りつぶし`path`(角丸矩形をpathで描画した
  ヘッダー等、ADR-014)。かつキャンバスの90%以上を覆う背景矩形は除外する(`build_mapping(..., svg_size=...)`)。
  `image`/`path`は幅・高さが`image_container_min_px`(既定120px)未満の場合も除外する(小さい
  アイコンがラベルの幅制約になってしまうのを防ぐ、ADR-012/014)。
- **真の内包**: テキストの代表点(x=anchor_x, y=推定top)がshapeのBBoxに`containment_margin_px`の
  余裕を持って含まれる。複数の候補が内包する場合は面積最小のものを選ぶ。
- **フォールバック**: 内包するshapeがない場合、距離`fallback_max_distance_px`以内の最近傍shapeに割り当てる。
- `contained` フラグで両者を区別 → テキスト配置(白背景の要否)・concat挙動が分岐。

### `concat_text.py` — 集約層

- 入力: `TextElem[]` + マッピング情報
- 出力: `TextBlock[]`(複数TextElemをY昇順に結合したもの)
- 結合条件: 同一コンテナ かつ 縦近接(`max_gap_ratio`) **かつ** 横近接(同程度のanchor_x)
  **かつ** フォントサイズが完全一致(ADR-016。`TextBlock`は単一font_sizeしか持てないため)
- 大面積コンテナ(`container_area_threshold_px2`)・閉じた塗り`path`コンテナ(ADR-014)はconcat対象外
- マッピング先が無い(`shape_id is None`)テキストは、他のテキストと結合されず必ず単独のブロックになる

### `fit_text.py` — 計測層

- PIL (`Pillow`) でMeiryo UIフォントを直接読み込み、`font.getlength()`で実際の描画幅を測定
- 太字テキストは `_BOLD_WIDTH_FACTOR`(1.12)を掛けて幅を安全側に見積もる(疑似ボールド対策)
- 折り返し幅はコンテナ(`content_width_px`)を基準にトークン単位(英数字は単語単位、CJKは1文字単位)で
  計算する(ADR-013)。単独のトークンが幅を超える場合でもトークン内では分割しない。
  `content_width_px is None`の場合は折り返しなし(自然な幅をそのまま使う)
- フォントサイズ自動縮小: `shrink_step_pt`刻みで`max_shrink_steps`回まで試行、下限は`min_size_pt`相当
- ブロックの高さ(`box_height_px`)は `font_size * (1 + (行数-1) * line_spacing)` で計算する。
  1行だけのブロックの高さが`line_spacing`倍に膨らんで隣接ブロックと重ならないようにするための式。
- フィット判定の許容量は `content_width_px - 0.5` のようにマイナス側(ADR-017)。プラス側だと
  「収まった」判定でも実際の描画で枠線に接して見えることがあったため。

### `build_pptx.py` — 出力層

- `CoordMapper`: SVG座標系 → PPTX座標系(EMU)への線形変換。アスペクト比を保ちつつスライド全体に中央配置。
  `scale`(EMU/px)・`offset_x_emu`・`offset_y_emu`を保持し、`point_emu()`/`len_emu()`/`px_to_pt()`を提供する。
- 図形描画: `draw_rect`/`draw_oval`/`draw_line`(Connector)/`draw_freeform`(polyline・polygon・path)/`draw_image`
  - `apply_shape_style()`が色・線・破線・opacity(alpha)を適用する共通処理。`has_fill_attr=False`は
    Connector(`.fill`属性を持たない)用、`paint_fill=False`は開いたpath/polyline(塗りつぶし対象外)用。
    **`has_fill_attr=True`のときは`paint_fill`の真偽に関わらず必ず`fill.solid()`か`fill.background()`の
    どちらかを呼ぶ**(呼ばないとPowerPointのデフォルトテーマ色(青)が残ってしまうバグがあったため、
    ADR-011参照)。
  - opacityはXML直接操作(`a:alpha`要素をsolidFill配下に追加)で反映する。
- `place_text_block()`: テキスト配置の核心ロジック
  - `anchor='start'` → `box_l = anchor_x` / `'middle'` → `anchor_x - box_width/2` / `'end'` → `anchor_x - box_width`
  - 例外: 小さいコンテナ(面積`small_container_area_px2`以下)内の`anchor='middle'`は、anchor_xを無視して
    コンテナ幅いっぱいに広げ中央揃えにする(`is_small_container`、ADR-015)。`start`/`end`は対象外。
  - 縦位置は原則 `block.top`。ただし「元は1行 かつ fit結果が複数行」の場合のみ、コンテナ内で縦中央揃えに補正
  - `content_width_px`/`content_height_px` は **真に内包されている場合のみ** コンテナの実サイズから計算する。
    フォールバック(浮遊)テキストは制約なし(自然サイズ)とする(ADR-010関連。極小のアイコン用rectに
    フォールバックした浮遊キャプションが不必要に折り返される事故を防ぐため)。
  - 白背景は「浮遊 かつ 近傍にshapeが存在する」場合のみ敷く
  - `tf.word_wrap = False`(折り返しは全てfit_text側で確定済みのため、PPTX側の再折り返しに委ねない)
  - 行間(`paragraph.line_spacing`)は倍率ではなく`Pt()`による絶対値指定(PowerPointの「単一行間隔」が
    フォント内部メトリクス依存でfont_sizeより大きくなることがあるため)
  - `_set_run_lang_and_charset()`: PowerPointのスペルチェック誤検出抑制のため各runに
    `lang="ja-JP"`/`altLang="ja-JP"`、Unicode範囲明示のため`<a:latin charset="0">`を設定する(ADR-019)

### `qa_capture.py` — QA層

- PowerPoint COMオブジェクト経由でスライドをPNG出力(150 DPI)
- Windows専用(COM依存)。`PageSetup.SlideWidth/Height`はポイント単位なので `/72*dpi` でpx換算する
- 生成PPTXの視覚確認に使用。CIには組み込まず手動QAとして運用
- 既存のPowerPointインスタンスに接続できた場合(`GetActiveObject`)は`Quit()`せず、自分で新規起動した
  場合のみ`Quit()`する(ユーザーの他のプレゼンテーションを巻き添えで閉じないため)

### `run.ps1` — 実行ラッパー

- `uv run`が使えない環境(ADR-020)向けに、uvのベースインタプリタを`uv python find 3.12`で動的に
  解決し、venvの`site-packages`(pywin32関連サブディレクトリ込み)を`PYTHONPATH`に設定してから
  直接実行する
- ASCII文字のみで記述する(PowerShell 5.1がUTF-8(BOM無し)の非ASCII文字を含むスクリプトを
  誤ってシステムのコードページで解釈し、パースエラーになる問題を回避するため)

---

## 座標変換の設計

```
SVG座標(px) × scale(EMU/px) + offset(EMU) = PPTX座標(EMU)

scale   = min(slide_w_emu / svg_w, slide_h_emu / svg_h)   ← アスペクト比を崩さない
offset  = スライド全体に対する中央揃え余白
```

Freeform図形は `build_freeform(start_x, start_y, scale=cm.scale)` + `add_line_segments(pts, close=...)` +
`convert_to_shape(origin_x=offset_x_emu, origin_y=offset_y_emu)` を用い、SVG px座標をそのまま
「ローカル座標」として渡すことで、都度の座標変換を省いている。

---

## 入力SVGのパターン

| パターン | no_text引数 | text_only引数 |
|----------|-------------|---------------|
| 単一SVG(図形+テキスト混在) | SVGファイル | 同じSVGファイル(重複指定可) |
| ペア形式(no_text + text_only) | no_text.svg | text_only.svg |

## 既知の制約

[docs/CONCEPT.md](CONCEPT.md) の「既知の制約・今後の検討事項」を参照。
