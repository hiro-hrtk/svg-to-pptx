"""計測層。PIL実測メトリクスでテキストの折り返し・フォントサイズ自動フィットを行う。

全ての計算はSVG px単位で行う(PPTX inchesへの変換はbuild_pptx側で行う)。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from PIL import ImageFont

# 英数字(記号込み)の連続を1トークンとして扱い、それ以外(CJK文字・記号等)は1文字ずつ
# 独立したトークンにする。空白はトークン間の区切りとして扱う(行頭には出さない)。
_TOKEN_RE = re.compile(r"[A-Za-z0-9_+\-./]+|\s+|.")

_font_cache: dict = {}


def _load_font(size_px: int, config: dict):
    key = size_px
    if key in _font_cache:
        return _font_cache[key]
    fcfg = config.get("font", {})
    font_file = fcfg.get("font_file", "C:/Windows/Fonts/meiryo.ttc")
    index = fcfg.get("font_file_index", 2)
    font = ImageFont.truetype(font_file, size=max(size_px, 1), index=index)
    _font_cache[key] = font
    return font


def _is_bold_weight(weight) -> bool:
    weight = str(weight or "").strip().lower()
    if weight == "bold":
        return True
    if weight.isdigit():
        return int(weight) >= 600
    return False


# meiryo.ttcにはUI Bold書体が無く、PPTX側はfont.bold=Trueによる疑似ボールドで
# 表現する(config.yaml参照)。疑似ボールドは実測(Regular)幅より広くレンダリング
# されるため、折り返し判定を安全側に倒す補正係数を掛ける。
_BOLD_WIDTH_FACTOR = 1.12


def measure_width_px(text: str, font_size_px: float, config: dict, bold: bool = False) -> float:
    if not text:
        return 0.0
    font = _load_font(round(font_size_px), config)
    width = font.getlength(text)
    return width * _BOLD_WIDTH_FACTOR if bold else width


def _wrap_paragraph(text: str, font_size_px: float, content_width_px, config, bold: bool = False) -> list:
    if content_width_px is None or content_width_px <= 0:
        return [text]
    if measure_width_px(text, font_size_px, config, bold) <= content_width_px:
        return [text]
    # トークン単位で折り返す(英数字は単語単位、CJK等は1文字単位)。
    # 単独のトークンがcontent_widthを超える場合でも、そのトークン内では分割しない
    # (「Embedding」→「Embed」/「ding」のような英単語の途中改行を避けるため)。
    lines = []
    current = ""
    for tok in _TOKEN_RE.findall(text):
        if tok.isspace() and not current:
            continue
        trial = current + tok
        if current and measure_width_px(trial, font_size_px, config, bold) > content_width_px:
            lines.append(current.rstrip())
            current = "" if tok.isspace() else tok
        else:
            current = trial
    if current.strip():
        lines.append(current.rstrip())
    return lines or [text]


@dataclass
class FitResult:
    font_size_px: float
    lines: list = field(default_factory=list)  # 折り返し後の全行(paragraph境界含む)
    box_width_px: float = 0.0
    box_height_px: float = 0.0


def fit_block(block, content_width_px, content_height_px, config) -> FitResult:
    fcfg = config.get("font", {})
    tcfg = config.get("text_box", {})
    min_size_pt_equiv_px = fcfg.get("min_size_pt", 8)
    max_steps = fcfg.get("max_shrink_steps", 20)
    step = fcfg.get("shrink_step_pt", 0.5)
    line_spacing = tcfg.get("line_spacing", 1.15)

    font_size = block.font_size
    bold = _is_bold_weight(block.font_weight)
    # min_size_pt はpt単位設定値だが、ここではpx単位のfont_sizeをそのまま下限として
    # 同程度の小ささまで許容する(px≒pt換算は呼び出し側のスケールに依存するため、
    # ここでは比率でなく直接pxベースで安全側に丸める)。
    min_size_px = min_size_pt_equiv_px

    for _ in range(max_steps + 1):
        lines = []
        for para in block.paragraphs:
            lines.extend(_wrap_paragraph(para, font_size, content_width_px, config, bold))

        max_w = max((measure_width_px(l, font_size, config, bold) for l in lines), default=0.0)
        # line_spacingは「行間」にのみ適用する(1行だけのブロックの外枠がfont_sizeより
        # 不必要に大きくなると、隣接する独立したテキストブロック同士が視覚的に重なるため)。
        total_h = font_size * (1 + (len(lines) - 1) * line_spacing)

        # 許容量はプラスではなくマイナス(=わずかに余裕を持たせる側)にする。
        # +0.5だと境界ぎりぎりで「収まった」と判定され、実際の描画で枠線に接して
        # 見えることがあったため(特に3行以上のブロックで顕著)。
        width_ok = content_width_px is None or max_w <= content_width_px - 0.5
        height_ok = content_height_px is None or total_h <= content_height_px - 0.5

        if (width_ok and height_ok) or font_size <= min_size_px:
            return FitResult(font_size_px=font_size, lines=lines, box_width_px=max_w, box_height_px=total_h)

        font_size = max(min_size_px, font_size - step)

    return FitResult(font_size_px=font_size, lines=lines, box_width_px=max_w, box_height_px=total_h)
