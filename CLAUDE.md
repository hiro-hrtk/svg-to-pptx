# CLAUDE.md — svg-to-pptx プロジェクトガイド

このファイルはClaude Codeがこのプロジェクトを操作する際に常に参照するルールを定義する。

---

## プロジェクト概要

SVGファイルをPowerPointのネイティブオブジェクト(図形・コネクタ・テキストボックス)に変換するパイプライン。
詳細は [docs/CONCEPT.md](docs/CONCEPT.md) / [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) を参照。

Claude Codeのスキルとしては `svg-to-pptx` という名前で登録されている([REF/SKILL.md](REF/SKILL.md)がベース、
実体は `~/.claude/skills/svg-to-pptx/SKILL.md`)。

---

## 実行環境

- Python環境・依存関係の管理には **uv** を使用する(pip/venv/pyenvは使わない)。
- 依存追加: `uv add <package>` / 実行: `uv run python code/xxx.py`
- 主要依存: `python-pptx` / `Pillow` / `pywin32` / `PyYAML`
- `qa_capture.py` は PowerPoint COM を使うため **Windows + PowerPointインストール必須**。

---

## フォルダ構成

```
svg-to-pptx/
├── CLAUDE.md             ← このファイル
├── pyproject.toml / uv.lock
├── code/                 ← Pythonソースコード
│   ├── svg_elements.py   パース層
│   ├── mapping.py        マッピング層
│   ├── concat_text.py    集約層
│   ├── fit_text.py       計測層
│   ├── build_pptx.py     出力層(パイプライン本体・CLI)
│   ├── qa_capture.py     QA層(PowerPoint COM)
│   └── config.yaml       設定
├── docs/                 ← ドキュメント一式
│   ├── CONCEPT.md
│   ├── ARCHITECTURE.md
│   ├── DECISIONS.md      (ADR)
│   ├── TEST_CASES.md
│   └── CHANGELOG.md
├── input/
│   └── test/             ← テスト用サンプルSVG
├── output/               ← 変換出力(PPTX・JSON・PNG)
│   └── <session>/
└── REF/                  ← 消失前に復旧できた旧ドキュメント一式(履歴参照用、正本はdocs/)
```

---

## ドキュメント更新ルール

### docs/CHANGELOG.md — **コード変更のたびに更新**

`code/` 配下の `.py` / `.yaml` を変更したら、必ず `## [Unreleased]` セクションを更新する。

| 変更の種類 | 記載するセクション |
|-----------|------------------|
| 新機能・新対応要素 | `### Added` |
| バグ修正 | `### Fixed`(症状・原因・修正を簡潔に) |
| 挙動変更・リファクタ | `### Changed` |
| 機能削除 | `### Removed` |

### docs/ARCHITECTURE.md — **設計変更時に更新**

新モジュール追加・既存モジュールの責務変更・データフロー変更のいずれかが発生したら更新する。

### docs/TEST_CASES.md — **SVG追加・QA実施時に更新**

- `input/test/` に新しいSVGを追加したとき → テストSVG一覧に行を追加、`⬜ 未検証`
- QAサイクルでPASSしたとき → `✅ 検証済み` に変更し、過去のQA結果ログに追記

### docs/DECISIONS.md — **重要な設計判断が生まれたとき**

複数方針を検討して一方を選んだ・既存アプローチを根本変更した・重要なトレードオフを選択した場合にADRを追記する(連番)。

---

## 開発フロー(QAサイクル)

```
1. build_pptx.py で PPTX 生成
   uv run python code/build_pptx.py --session <name> \
       --no-text input/test/<svg> --text-only input/test/<svg>

2. qa_capture.py で PNG 出力
   uv run python code/qa_capture.py output/<name>/<file>.pptx

3. 目視確認(PNG)+ マッピング確認(mapping.json)

4. 問題があれば該当モジュールを修正 → 1 に戻る

5. PASS したら docs/CHANGELOG.md / docs/TEST_CASES.md を更新
```

## input/ フォルダの管理ルール

- テスト用SVGは `input/test/` 直下にフラットに配置する(サブフォルダは作らない)
- zip・png 等の非SVGファイルは置かない
- 新しいSVGを追加したら `docs/TEST_CASES.md` を更新する

## output/ フォルダの管理ルール

- `output/<session>/` に `.pptx` / `_mapping.json` / `_qa_slide*.png` を出力する
- 古いセッションフォルダは手動で削除してよい(再生成可能なため)
