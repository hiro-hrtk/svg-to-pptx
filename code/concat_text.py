"""集約層。同一コンテナに所属する隣接TextElemをTextBlockにconcatする。

結合条件: 同一コンテナ かつ 縦近接 かつ 横近接(ADR-006)。
大面積コンテナはconcat対象外(container_area_threshold_px2)。
"""
from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class TextBlock:
    id: str
    shape_id: str | None
    contained: bool
    anchor: str
    left: float          # anchor_x (テキスト自身の元座標)
    top: float            # 先頭要素の推定top(y - font_size)
    font_size: float
    font_weight: str
    fill: str
    paragraphs: list = field(default_factory=list)  # 各paragraph = [line, line, ...]


def _shape_area(shapes_by_id, shape_id):
    if shape_id is None:
        return 0.0
    s = shapes_by_id.get(shape_id)
    return s.area if s else 0.0


def build_text_blocks(texts, mapping, contained_ids, shapes, config):
    ccfg = config.get("concat", {})
    max_gap_ratio = ccfg.get("max_gap_ratio", 2.2)
    area_threshold = ccfg.get("container_area_threshold_px2", 60000)

    shapes_by_id = {s.id: s for s in shapes}

    # コンテナごと(Noneは各text個別扱い)にグルーピング
    groups: dict = {}
    for t in texts:
        shape_id = mapping.get(t.id)
        key = shape_id if shape_id is not None else f"__solo__{t.id}"
        groups.setdefault(key, []).append(t)

    blocks = []
    for key, group_texts in groups.items():
        shape_id = None if key.startswith("__solo__") else key
        allow_concat = (
            shape_id is not None
            and _shape_area(shapes_by_id, shape_id) <= area_threshold
        )

        group_texts.sort(key=lambda t: t.y)

        current = None
        for t in group_texts:
            contained = t.id in contained_ids
            if current is not None and allow_concat:
                prev = current["last_text"]
                vgap = (t.y - t.font_size) - prev.y
                same_anchor_x = abs(t.x - prev.x) < max(prev.font_size, t.font_size) * 3
                if (
                    vgap <= prev.font_size * max_gap_ratio
                    and vgap >= -prev.font_size
                    and same_anchor_x
                    and t.anchor == prev.anchor
                ):
                    for line in t.lines:
                        current["paragraphs"].append(line)
                    current["last_text"] = t
                    continue
            # 新しいブロックを開始
            if current is not None:
                blocks.append(current)
            current = {
                "shape_id": shape_id,
                "contained": contained,
                "anchor": t.anchor,
                "left": t.x,
                "top": t.y - t.font_size,
                "font_size": t.font_size,
                "font_weight": t.font_weight,
                "fill": t.fill,
                "paragraphs": list(t.lines),
                "last_text": t,
            }
        if current is not None:
            blocks.append(current)

    result = []
    for i, b in enumerate(blocks):
        result.append(TextBlock(
            id=f"block_{i}",
            shape_id=b["shape_id"],
            contained=b["contained"],
            anchor=b["anchor"],
            left=b["left"],
            top=b["top"],
            font_size=b["font_size"],
            font_weight=b["font_weight"],
            fill=b["fill"],
            paragraphs=b["paragraphs"],
        ))
    return result
