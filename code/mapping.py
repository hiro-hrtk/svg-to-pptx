"""マッピング層。TextElem を「所属するShapeElem」に対応付ける。

真の内包(contained)と最近傍フォールバックを区別する(ADR-003)。
コンテナ候補は rect / image のみ(円・線・pathなどの装飾図形は文字を内包しない前提)。
"""
from __future__ import annotations

import math


def _contained_point(px, py, bbox, margin):
    l, t, w, h = bbox
    return (l - margin <= px <= l + w + margin) and (t - margin <= py <= t + h + margin)


def _distance_to_bbox(px, py, bbox):
    l, t, w, h = bbox
    r, b = l + w, t + h
    dx = max(l - px, 0, px - r)
    dy = max(t - py, 0, py - b)
    return math.hypot(dx, dy)


def build_mapping(shapes, texts, config, svg_size=None):
    """
    Returns:
        mapping: {text_id: shape_id or None}
        contained_ids: set of text_id that are truly contained
    """
    mcfg = config.get("mapping", {})
    margin = mcfg.get("containment_margin_px", 2)
    fallback_max = mcfg.get("fallback_max_distance_px", 60)

    image_min_px = mcfg.get("image_container_min_px", 120)
    containers = [
        s for s in shapes
        if s.kind == "rect" or (s.kind == "image" and s.bbox[2] >= image_min_px and s.bbox[3] >= image_min_px)
    ]
    # 小さい<image>(アイコン画像)はコンテナ候補から除外する(ADR-012)。
    # アイコン画像の直下に置かれたキャプションが、本来無関係なアイコンの狭い幅に
    # 「内包」判定されてしまい、単語の途中で折り返される事故を防ぐため。

    # キャンバス全面を覆う背景矩形は「内包図形」の候補から除外する。
    # 除外しないと、すべてのテキストがこの巨大矩形に内包され(面積最小選択が機能せず)、
    # 本来のフォールバック/floating判定が働かなくなる。
    if svg_size:
        svg_w, svg_h = svg_size
        containers = [
            s for s in containers
            if not (s.bbox[2] >= 0.9 * svg_w and s.bbox[3] >= 0.9 * svg_h)
        ]

    mapping = {}
    contained_ids = set()

    for t in texts:
        # テキストの代表点: x=anchor_x, y=推定top(ベースラインからfont_size分上)
        px = t.x
        py = t.y - t.font_size

        best_shape = None
        best_area = None
        for s in containers:
            if _contained_point(px, py, s.bbox, margin):
                area = s.area
                if best_area is None or area < best_area:
                    best_area = area
                    best_shape = s

        if best_shape is not None:
            mapping[t.id] = best_shape.id
            contained_ids.add(t.id)
            continue

        nearest = None
        nearest_dist = None
        for s in containers:
            d = _distance_to_bbox(px, py, s.bbox)
            if nearest_dist is None or d < nearest_dist:
                nearest_dist = d
                nearest = s

        if nearest is not None and nearest_dist <= fallback_max:
            mapping[t.id] = nearest.id
        else:
            mapping[t.id] = None

    return mapping, contained_ids
