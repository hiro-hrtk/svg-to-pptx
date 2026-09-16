"""SVGパース層。SVG XML を ShapeElem / TextElem に正規化する。

対応要素: rect, line, polyline, polygon, circle, ellipse, path, image, text
path は M/L/H/V/C/Q/A/Z (相対/絶対) をサンプリングして折れ線近似する。

スタイルの解決優先順位: inline style > <style>内の単純なクラスセレクタ(.name{...}) >
直接属性 > default (get_prop/_parse_style_rules参照)。

`marker-end="url(#id)"` (SVG標準の矢じり) は <defs><marker>...</marker></defs> を解決し、
線・pathの終点に向き(orient=auto相当)を合わせた塗りつぶしFreeformとして追加描画する。
"""
from __future__ import annotations

import base64
import math
import re
from dataclasses import dataclass, field
from typing import Optional
from xml.etree import ElementTree as ET

SVG_NS = "http://www.w3.org/2000/svg"
XLINK_NS = "http://www.w3.org/1999/xlink"


def _tag(elem) -> str:
    t = elem.tag
    if t.startswith("{"):
        return t.split("}", 1)[1]
    return t


NAMED_COLORS = {
    "white": "FFFFFF",
    "black": "000000",
    "none": None,
    "red": "FF0000",
    "green": "008000",
    "blue": "0000FF",
    "gray": "808080",
    "grey": "808080",
    "transparent": None,
}


@dataclass
class ShapeElem:
    id: str
    kind: str  # rect|line|polyline|polygon|circle|ellipse|path|image
    x: float = 0.0
    y: float = 0.0
    w: float = 0.0
    h: float = 0.0
    rx: float = 0.0
    ry: float = 0.0
    points: list = field(default_factory=list)  # [(x,y), ...]
    closed: bool = False
    fill: Optional[str] = None  # hex "RRGGBB" or None(=塗りなし)
    stroke: Optional[str] = None
    stroke_width: float = 0.0
    dasharray: Optional[str] = None
    opacity: float = 1.0
    image_data: Optional[bytes] = None
    image_ext: str = "png"

    @property
    def bbox(self):
        """(left, top, width, height) in SVG px."""
        if self.kind in ("rect", "circle", "ellipse", "image"):
            return (self.x, self.y, self.w, self.h)
        if self.points:
            xs = [p[0] for p in self.points]
            ys = [p[1] for p in self.points]
            l, t = min(xs), min(ys)
            return (l, t, max(xs) - l, max(ys) - t)
        return (self.x, self.y, self.w, self.h)

    @property
    def area(self) -> float:
        _, _, w, h = self.bbox
        return max(w, 0.0) * max(h, 0.0)


@dataclass
class TextElem:
    id: str
    x: float
    y: float
    anchor: str  # start|middle|end
    font_size: float
    font_weight: str
    fill: str
    lines: list = field(default_factory=list)      # 各tspanの文字列
    line_dy: list = field(default_factory=list)     # 各tspanのdy(px)


def _style_dict(style_attr: Optional[str]) -> dict:
    d = {}
    if not style_attr:
        return d
    for part in style_attr.split(";"):
        part = part.strip()
        if not part or ":" not in part:
            continue
        k, v = part.split(":", 1)
        d[k.strip()] = v.strip()
    return d


def _num(value, default: float = 0.0) -> float:
    """'38px'のような単位付き数値文字列にも対応したfloat変換。"""
    if value is None:
        return default
    if isinstance(value, (int, float)):
        return float(value)
    m = re.match(r"\s*(-?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?)", str(value))
    return float(m.group(1)) if m else default


def get_prop(elem, name: str, class_rules: Optional[dict] = None, default=None):
    """優先順位: inline style > CSSクラス(スタイルシート順で後勝ち) > 直接属性 > default。"""
    style = _style_dict(elem.get("style"))
    if name in style:
        return style[name]
    if class_rules:
        classes = (elem.get("class") or "").split()
        value = None
        for cname, rules in class_rules.items():
            if cname in classes and name in rules:
                value = rules[name]
        if value is not None:
            return value
    if name in elem.attrib:
        return elem.get(name)
    return default


_CLASS_RULE_RE = re.compile(r"\.([\w-]+)\s*\{([^}]*)\}")


def _parse_style_rules(root) -> dict:
    """<style>ブロック中の単純なクラスセレクタ(`.name { prop:value; }`)を解決する。
    複雑なセレクタ(結合子・疑似クラス・idセレクタ等)は非対応。
    """
    rules: dict = {}
    for style_elem in root.iter():
        if _tag(style_elem) != "style":
            continue
        text = "".join(style_elem.itertext())
        for cname, body in _CLASS_RULE_RE.findall(text):
            rules[cname] = _style_dict(body)
    return rules


def _resolve_gradient_color(grad_id: str, defs: dict) -> Optional[str]:
    stops = defs.get(grad_id)
    if not stops:
        return None
    rs = gs = bs = 0
    n = len(stops)
    for hexcol in stops:
        r = int(hexcol[0:2], 16)
        g = int(hexcol[2:4], 16)
        b = int(hexcol[4:6], 16)
        rs += r
        gs += g
        bs += b
    return f"{rs // n:02X}{gs // n:02X}{bs // n:02X}"


def parse_color(value: Optional[str], defs: dict) -> Optional[str]:
    if value is None:
        return None
    value = value.strip()
    if value == "" or value.lower() == "none":
        return None
    if value.startswith("url(#"):
        grad_id = value[5:-1]
        return _resolve_gradient_color(grad_id, defs)
    if value.startswith("#"):
        hexval = value[1:]
        if len(hexval) == 3:
            hexval = "".join(c * 2 for c in hexval)
        return hexval.upper()
    low = value.lower()
    if low in NAMED_COLORS:
        return NAMED_COLORS[low]
    return None


def _parse_defs(root) -> dict:
    """gradient id -> [stop_hex, ...] のマップを作る。"""
    defs = {}
    for grad in root.iter():
        if _tag(grad) in ("linearGradient", "radialGradient"):
            gid = grad.get("id")
            if not gid:
                continue
            stops = []
            for stop in grad:
                if _tag(stop) != "stop":
                    continue
                col = get_prop(stop, "stop-color", default="#000000")
                hexcol = parse_color(col, {})
                if hexcol:
                    stops.append(hexcol)
            if stops:
                defs[gid] = stops
    return defs


_NUM_RE = re.compile(r"-?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?")


def _parse_points_attr(points_attr: str):
    nums = [float(n) for n in _NUM_RE.findall(points_attr)]
    return [(nums[i], nums[i + 1]) for i in range(0, len(nums) - 1, 2)]


_PATH_TOKEN_RE = re.compile(r"([MmLlHhVvCcSsQqTtAaZz])|(-?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?)")


def _tokenize_path(d: str):
    tokens = []
    for cmd, num in _PATH_TOKEN_RE.findall(d):
        if cmd:
            tokens.append(("cmd", cmd))
        elif num:
            tokens.append(("num", float(num)))
    return tokens


def _arc_to_points(x1, y1, rx, ry, phi_deg, large_arc, sweep, x2, y2, n_seg=24):
    """SVG elliptical arc -> サンプリング点列(始点は含まない)。"""
    if rx == 0 or ry == 0 or (x1 == x2 and y1 == y2):
        return [(x2, y2)]
    phi = math.radians(phi_deg)
    cos_phi, sin_phi = math.cos(phi), math.sin(phi)
    dx2 = (x1 - x2) / 2.0
    dy2 = (y1 - y2) / 2.0
    x1p = cos_phi * dx2 + sin_phi * dy2
    y1p = -sin_phi * dx2 + cos_phi * dy2

    rx, ry = abs(rx), abs(ry)
    lam = (x1p ** 2) / (rx ** 2) + (y1p ** 2) / (ry ** 2)
    if lam > 1:
        s = math.sqrt(lam)
        rx *= s
        ry *= s

    sign = -1 if large_arc == sweep else 1
    num = rx ** 2 * ry ** 2 - rx ** 2 * y1p ** 2 - ry ** 2 * x1p ** 2
    den = rx ** 2 * y1p ** 2 + ry ** 2 * x1p ** 2
    co = sign * math.sqrt(max(num / den, 0.0)) if den != 0 else 0.0
    cxp = co * (rx * y1p / ry)
    cyp = co * (-ry * x1p / rx)

    cx = cos_phi * cxp - sin_phi * cyp + (x1 + x2) / 2.0
    cy = sin_phi * cxp + cos_phi * cyp + (y1 + y2) / 2.0

    def ang(ux, uy, vx, vy):
        dot = ux * vx + uy * vy
        length = math.sqrt((ux ** 2 + uy ** 2) * (vx ** 2 + vy ** 2))
        a = math.acos(max(-1.0, min(1.0, dot / length))) if length else 0.0
        if ux * vy - uy * vx < 0:
            a = -a
        return a

    theta1 = ang(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry)
    dtheta = ang((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry)
    if sweep == 0 and dtheta > 0:
        dtheta -= 2 * math.pi
    elif sweep == 1 and dtheta < 0:
        dtheta += 2 * math.pi

    pts = []
    steps = max(4, int(n_seg * abs(dtheta) / (2 * math.pi)) + 2)
    for i in range(1, steps + 1):
        t = theta1 + dtheta * i / steps
        ex = cx + rx * math.cos(t) * cos_phi - ry * math.sin(t) * sin_phi
        ey = cy + rx * math.cos(t) * sin_phi + ry * math.sin(t) * cos_phi
        pts.append((ex, ey))
    return pts


def _cubic_bezier_points(p0, p1, p2, p3, n=16):
    pts = []
    for i in range(1, n + 1):
        t = i / n
        mt = 1 - t
        x = mt ** 3 * p0[0] + 3 * mt ** 2 * t * p1[0] + 3 * mt * t ** 2 * p2[0] + t ** 3 * p3[0]
        y = mt ** 3 * p0[1] + 3 * mt ** 2 * t * p1[1] + 3 * mt * t ** 2 * p2[1] + t ** 3 * p3[1]
        pts.append((x, y))
    return pts


def _quad_bezier_points(p0, p1, p2, n=12):
    pts = []
    for i in range(1, n + 1):
        t = i / n
        mt = 1 - t
        x = mt ** 2 * p0[0] + 2 * mt * t * p1[0] + t ** 2 * p2[0]
        y = mt ** 2 * p0[1] + 2 * mt * t * p1[1] + t ** 2 * p2[1]
        pts.append((x, y))
    return pts


def parse_path_d(d: str):
    """path 'd' -> (points, closed) の折れ線近似。"""
    tokens = _tokenize_path(d)
    i = 0
    n = len(tokens)
    cur = (0.0, 0.0)
    start = (0.0, 0.0)
    points = []
    closed = False
    cmd = None

    def nums(count):
        nonlocal i
        vals = []
        for _ in range(count):
            vals.append(tokens[i][1])
            i += 1
        return vals

    while i < n:
        kind, val = tokens[i]
        if kind == "cmd":
            cmd = val
            i += 1
        # 数値が続く限り同じコマンドの繰り返し(暗黙の連続座標)
        if cmd in ("M", "m"):
            x, y = nums(2)
            if cmd == "m":
                x, y = cur[0] + x, cur[1] + y
            cur = (x, y)
            start = cur
            points.append(cur)
            cmd = "L" if cmd == "M" else "l"
        elif cmd in ("L", "l"):
            x, y = nums(2)
            if cmd == "l":
                x, y = cur[0] + x, cur[1] + y
            cur = (x, y)
            points.append(cur)
        elif cmd in ("H", "h"):
            (x,) = nums(1)
            x = x if cmd == "H" else cur[0] + x
            cur = (x, cur[1])
            points.append(cur)
        elif cmd in ("V", "v"):
            (y,) = nums(1)
            y = y if cmd == "V" else cur[1] + y
            cur = (cur[0], y)
            points.append(cur)
        elif cmd in ("C", "c"):
            x1, y1, x2, y2, x, y = nums(6)
            if cmd == "c":
                x1, y1 = cur[0] + x1, cur[1] + y1
                x2, y2 = cur[0] + x2, cur[1] + y2
                x, y = cur[0] + x, cur[1] + y
            points.extend(_cubic_bezier_points(cur, (x1, y1), (x2, y2), (x, y)))
            cur = (x, y)
        elif cmd in ("Q", "q"):
            x1, y1, x, y = nums(4)
            if cmd == "q":
                x1, y1 = cur[0] + x1, cur[1] + y1
                x, y = cur[0] + x, cur[1] + y
            points.extend(_quad_bezier_points(cur, (x1, y1), (x, y)))
            cur = (x, y)
        elif cmd in ("A", "a"):
            rx, ry, rot, laf, sf, x, y = nums(7)
            if cmd == "a":
                x, y = cur[0] + x, cur[1] + y
            arc_pts = _arc_to_points(cur[0], cur[1], rx, ry, rot, int(laf), int(sf), x, y)
            points.extend(arc_pts)
            cur = (x, y)
        elif cmd in ("Z", "z"):
            if cur != start:
                points.append(start)
            cur = start
            closed = True
        else:
            i += 1
    return points, closed


def _element_opacity(elem, class_rules=None) -> float:
    try:
        return float(get_prop(elem, "opacity", class_rules, 1.0))
    except (TypeError, ValueError):
        return 1.0


def _load_image_data(href: str):
    m = re.match(r"data:image/(\w+);base64,(.*)", href, re.DOTALL)
    if not m:
        return None, "png"
    ext, b64 = m.group(1), m.group(2)
    try:
        return base64.b64decode(b64), ext
    except Exception:
        return None, ext


def _iter_recursive(elem):
    """<g> をフラット化しつつ子孫を全て辿る(この描画順序のSVGはtransform無し前提)。"""
    for child in elem:
        tag = _tag(child)
        if tag == "g":
            yield from _iter_recursive(child)
        else:
            yield child


def _parse_markers(root) -> dict:
    """<marker>定義を id -> {points, closed, fill, mw, mh, rx, ry, units} に変換する。"""
    markers = {}
    for elem in root.iter():
        if _tag(elem) != "marker":
            continue
        mid = elem.get("id")
        if not mid:
            continue
        pts, closed, fill = [], False, "000000"
        for child in elem:
            ctag = _tag(child)
            if ctag == "path":
                pts, closed = parse_path_d(child.get("d", ""))
                fill = parse_color(get_prop(child, "fill", default="#000000"), {}) or "000000"
                break
            if ctag == "polygon":
                pts = _parse_points_attr(child.get("points", ""))
                closed = True
                fill = parse_color(get_prop(child, "fill", default="#000000"), {}) or "000000"
                break
        if not pts:
            continue
        markers[mid] = dict(
            points=pts, closed=closed, fill=fill,
            mw=float(elem.get("markerWidth", 3) or 3),
            mh=float(elem.get("markerHeight", 3) or 3),
            rx=float(elem.get("refX", 0) or 0),
            ry=float(elem.get("refY", 0) or 0),
            units=elem.get("markerUnits", "strokeWidth"),
        )
    return markers


def _marker_url_id(value: Optional[str]) -> Optional[str]:
    if not value:
        return None
    m = re.match(r"url\(#(.+)\)", value.strip())
    return m.group(1) if m else None


def _build_marker_shape(marker: dict, end_point, angle_rad: float, stroke_width: float, sid: str):
    """marker定義 + 線端の位置・向きから矢じり形状のShapeElemを組み立てる(orient=auto相当)。"""
    scale = stroke_width if marker["units"] != "userSpaceOnUse" else 1.0
    if scale <= 0:
        scale = 1.0
    rx, ry = marker["rx"], marker["ry"]
    ca, sa = math.cos(angle_rad), math.sin(angle_rad)
    ex, ey = end_point
    pts = []
    for lx, ly in marker["points"]:
        dx = (lx - rx) * scale
        dy = (ly - ry) * scale
        # ローカルX軸=進行方向、ローカルY軸=進行方向に垂直な向きとして回転
        rxp = dx * ca - dy * sa
        ryp = dx * sa + dy * ca
        pts.append((ex + rxp, ey + ryp))
    return ShapeElem(sid, "path", points=pts, closed=True, fill=marker["fill"])


def _line_angle(points) -> float:
    if len(points) < 2:
        return 0.0
    (x1, y1), (x2, y2) = points[-2], points[-1]
    if x1 == x2 and y1 == y2:
        return 0.0
    return math.atan2(y2 - y1, x2 - x1)


def _parse_shape_elem(elem, tag, defs, counter, class_rules=None, markers=None):
    fill = parse_color(get_prop(elem, "fill", class_rules), defs)
    stroke = parse_color(get_prop(elem, "stroke", class_rules), defs)
    stroke_width = _num(get_prop(elem, "stroke-width", class_rules, 0), 0.0)
    dasharray = get_prop(elem, "stroke-dasharray", class_rules)
    opacity = _element_opacity(elem, class_rules)
    sid = f"shape_{counter}"

    if tag == "rect":
        x = float(elem.get("x", 0))
        y = float(elem.get("y", 0))
        w = float(elem.get("width", 0))
        h = float(elem.get("height", 0))
        rx = float(elem.get("rx", 0) or 0)
        ry = float(elem.get("ry", rx) or rx)
        return ShapeElem(sid, "rect", x=x, y=y, w=w, h=h, rx=rx, ry=ry,
                          fill=fill, stroke=stroke, stroke_width=stroke_width,
                          dasharray=dasharray, opacity=opacity)
    if tag == "circle":
        cx = float(elem.get("cx", 0))
        cy = float(elem.get("cy", 0))
        r = float(elem.get("r", 0))
        return ShapeElem(sid, "circle", x=cx - r, y=cy - r, w=2 * r, h=2 * r,
                          fill=fill, stroke=stroke, stroke_width=stroke_width,
                          dasharray=dasharray, opacity=opacity)
    if tag == "ellipse":
        cx = float(elem.get("cx", 0))
        cy = float(elem.get("cy", 0))
        rx = float(elem.get("rx", 0))
        ry = float(elem.get("ry", 0))
        return ShapeElem(sid, "ellipse", x=cx - rx, y=cy - ry, w=2 * rx, h=2 * ry,
                          fill=fill, stroke=stroke, stroke_width=stroke_width,
                          dasharray=dasharray, opacity=opacity)
    if tag == "line":
        x1 = float(elem.get("x1", 0))
        y1 = float(elem.get("y1", 0))
        x2 = float(elem.get("x2", 0))
        y2 = float(elem.get("y2", 0))
        return ShapeElem(sid, "line", points=[(x1, y1), (x2, y2)],
                          stroke=stroke, stroke_width=stroke_width,
                          dasharray=dasharray, opacity=opacity)
    if tag in ("polyline", "polygon"):
        pts = _parse_points_attr(elem.get("points", ""))
        return ShapeElem(sid, tag, points=pts, closed=(tag == "polygon"),
                          fill=fill, stroke=stroke, stroke_width=stroke_width,
                          dasharray=dasharray, opacity=opacity)
    if tag == "path":
        pts, closed = parse_path_d(elem.get("d", ""))
        return ShapeElem(sid, "path", points=pts, closed=closed,
                          fill=fill, stroke=stroke, stroke_width=stroke_width,
                          dasharray=dasharray, opacity=opacity)
    if tag == "image":
        href = elem.get(f"{{{XLINK_NS}}}href") or elem.get("href") or ""
        data, ext = _load_image_data(href)
        x = float(elem.get("x", 0))
        y = float(elem.get("y", 0))
        w = float(elem.get("width", 0))
        h = float(elem.get("height", 0))
        return ShapeElem(sid, "image", x=x, y=y, w=w, h=h,
                          image_data=data, image_ext=ext, opacity=opacity)
    return None


def parse_shapes(svg_path: str):
    """no_text(相当)のSVGから図形要素を抽出する。<text>は無視する。"""
    tree = ET.parse(svg_path)
    root = tree.getroot()
    defs = _parse_defs(root)
    class_rules = _parse_style_rules(root)
    markers = _parse_markers(root)
    shapes = []
    counter = 0
    for elem in _iter_recursive(root):
        tag = _tag(elem)
        if tag not in ("rect", "circle", "ellipse", "line", "polyline", "polygon", "path", "image"):
            continue
        shape = _parse_shape_elem(elem, tag, defs, counter, class_rules, markers)
        if shape is None:
            continue
        shapes.append(shape)
        counter += 1

        # marker-end(SVG標準の矢じり)は別のFreeform図形として追加する。
        if tag in ("line", "polyline", "path") and shape.points:
            marker_id = _marker_url_id(get_prop(elem, "marker-end", class_rules))
            marker = markers.get(marker_id) if marker_id else None
            if marker:
                angle = _line_angle(shape.points)
                arrow = _build_marker_shape(
                    marker, shape.points[-1], angle, shape.stroke_width, f"shape_{counter}"
                )
                shapes.append(arrow)
                counter += 1
    return shapes


def parse_texts(svg_path: str):
    """text_only(相当)のSVGからテキスト要素を抽出する。<text>のみ対象。"""
    tree = ET.parse(svg_path)
    root = tree.getroot()
    defs = _parse_defs(root)
    class_rules = _parse_style_rules(root)
    texts = []
    counter = 0
    for elem in _iter_recursive(root):
        if _tag(elem) != "text":
            continue
        x = float(elem.get("x", 0))
        y = float(elem.get("y", 0))
        anchor = get_prop(elem, "text-anchor", class_rules, "start")
        font_size = _num(get_prop(elem, "font-size", class_rules, 16), 16.0)
        font_weight = str(get_prop(elem, "font-weight", class_rules, "400"))
        fill = parse_color(get_prop(elem, "fill", class_rules, "#000000"), defs) or "000000"

        lines = []
        line_dy = []
        tspans = [c for c in elem if _tag(c) == "tspan"]
        if tspans:
            for ts in tspans:
                lines.append("".join(ts.itertext()))
                try:
                    dy = float(ts.get("dy", 0) or 0)
                except ValueError:
                    dy = 0.0
                line_dy.append(dy)
        else:
            text = "".join(elem.itertext()).strip()
            if text:
                lines.append(text)
                line_dy.append(0.0)

        if not lines:
            continue

        tid = f"text_{counter}"
        texts.append(TextElem(tid, x, y, anchor, font_size, font_weight, fill, lines, line_dy))
        counter += 1
    return texts
