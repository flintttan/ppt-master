#!/usr/bin/env python3
"""Create a PPTX with editable text from SVG-image deck.

Input deck is an SVG-picture-based PPTX (each slide is a single picture).
This tool rebuilds a new PPTX where:
- Slide background is rendered from the original SVG with ALL <text> removed.
- Each SVG <text> element becomes a PowerPoint textbox (editable in WPS/Office).

This prioritizes editability of text (requirement A) while keeping high visual
fidelity via a high-res background render.

Usage:
  python3 ppt-master/tools/svg_text_overlay_to_pptx.py \
    FabriqueAIHub_答辩_ppt169_20260119/FabriqueAIHub_答辩.pptx \
    -o FabriqueAIHub_答辩_ppt169_20260119/FabriqueAIHub_答辩_可编辑文字版.pptx
"""

from __future__ import annotations

import argparse
import re
import subprocess
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET

from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor

EMU_PER_INCH = 914400
EMU_PER_PIXEL = EMU_PER_INCH / 96  # PPT uses 96dpi mapping in many tooling chains

SVG_NS = "http://www.w3.org/2000/svg"


def _local_name(tag: str) -> str:
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def _parse_float(value: str | None, default: float = 0.0) -> float:
    if value is None:
        return default
    try:
        return float(value)
    except Exception:
        return default


def _parse_color_hex(color: str | None, default: tuple[int, int, int] = (0, 0, 0)) -> tuple[int, int, int]:
    if not color:
        return default
    color = color.strip()
    if not color.startswith("#"):
        return default
    hexv = color[1:]
    if len(hexv) == 3:
        try:
            r = int(hexv[0] * 2, 16)
            g = int(hexv[1] * 2, 16)
            b = int(hexv[2] * 2, 16)
            return (r, g, b)
        except Exception:
            return default
    if len(hexv) == 6:
        try:
            r = int(hexv[0:2], 16)
            g = int(hexv[2:4], 16)
            b = int(hexv[4:6], 16)
            return (r, g, b)
        except Exception:
            return default
    return default


def _blend_with_white(rgb: tuple[int, int, int], opacity: float) -> tuple[int, int, int]:
    # pptx font color doesn't support alpha, approximate by blending with white.
    o = min(1.0, max(0.0, opacity))
    r, g, b = rgb
    r2 = int(round(r * o + 255 * (1.0 - o)))
    g2 = int(round(g * o + 255 * (1.0 - o)))
    b2 = int(round(b * o + 255 * (1.0 - o)))
    return (r2, g2, b2)


def _first_font_name(font_family: str | None) -> str | None:
    if not font_family:
        return None
    # e.g. "Montserrat, Source Han Sans SC, ..."
    first = font_family.split(",", 1)[0].strip()
    first = first.strip('"\'')
    return first or None


def _is_cjk(s: str) -> bool:
    # Cheap heuristic: any CJK Unified ideographs or fullwidth.
    for ch in s:
        o = ord(ch)
        if 0x4E00 <= o <= 0x9FFF:
            return True
        if 0x3400 <= o <= 0x4DBF:
            return True
        if 0xF900 <= o <= 0xFAFF:
            return True
    return False


def _estimate_text_width_px(text: str, font_px: float) -> float:
    # Heuristic: wide chars ~1.0em, latin ~0.55em, punctuation ~0.35em.
    width_em = 0.0
    for ch in text:
        o = ord(ch)
        if ch.isspace():
            width_em += 0.35
        elif 0x4E00 <= o <= 0x9FFF:
            width_em += 1.0
        elif 0x3400 <= o <= 0x4DBF:
            width_em += 1.0
        elif 0xF900 <= o <= 0xFAFF:
            width_em += 1.0
        elif ch in "•·-–—|丨/\\:：,，.。;；()（）[]【】{}“”\"'" :
            width_em += 0.35
        else:
            width_em += 0.55
    # Add a bit of padding.
    return max(4.0, width_em * font_px + 6.0)


@dataclass(frozen=True)
class SvgText:
    x: float
    y: float
    anchor: str
    text: str
    font_family: str | None
    font_px: float
    font_weight: str | None
    fill: str | None
    fill_opacity: float


def extract_svg_texts(svg_root: ET.Element) -> list[SvgText]:
    texts: list[SvgText] = []
    for el in svg_root.iter():
        if _local_name(el.tag) != "text":
            continue

        # Skip empty/whitespace-only nodes.
        raw = "".join(el.itertext()).strip("\n")
        if raw.strip() == "":
            continue

        x = _parse_float(el.attrib.get("x"), 0.0)
        y = _parse_float(el.attrib.get("y"), 0.0)
        anchor = (el.attrib.get("text-anchor") or "start").strip()

        font_family = el.attrib.get("font-family")
        font_px = _parse_float(el.attrib.get("font-size"), 16.0)
        font_weight = el.attrib.get("font-weight")
        fill = el.attrib.get("fill")

        # Prefer element fill-opacity, fall back to generic opacity if present.
        fill_opacity = _parse_float(el.attrib.get("fill-opacity"), 1.0)
        if "opacity" in el.attrib and "fill-opacity" not in el.attrib:
            fill_opacity = _parse_float(el.attrib.get("opacity"), fill_opacity)

        texts.append(
            SvgText(
                x=x,
                y=y,
                anchor=anchor,
                text=raw,
                font_family=font_family,
                font_px=font_px,
                font_weight=font_weight,
                fill=fill,
                fill_opacity=fill_opacity,
            )
        )

    return texts


def remove_all_text_elements(svg_root: ET.Element) -> None:
    # ElementTree has no parent pointers, remove via parent traversal.
    for parent in list(svg_root.iter()):
        children = list(parent)
        if not children:
            continue
        for child in children:
            if _local_name(child.tag) == "text":
                parent.remove(child)


def render_svg_to_png(svg_path: Path, png_path: Path, width_px: int, height_px: int) -> None:
    cmd = [
        "rsvg-convert",
        "-f",
        "png",
        "-w",
        str(width_px),
        "-h",
        str(height_px),
        "-o",
        str(png_path),
        str(svg_path),
    ]
    subprocess.run(cmd, check=True)


def _px_to_emu(px: float) -> int:
    return int(round(px * EMU_PER_PIXEL))


def _px_to_pt(px: float) -> float:
    return px * 72.0 / 96.0


def _ppt_align_from_anchor(anchor: str) -> int:
    a = anchor.lower()
    if a == "middle":
        return PP_ALIGN.CENTER
    if a == "end":
        return PP_ALIGN.RIGHT
    return PP_ALIGN.LEFT


def _bold_from_weight(font_weight: str | None) -> bool:
    if not font_weight:
        return False
    w = font_weight.strip().lower()
    if w == "bold":
        return True
    m = re.match(r"^(\d+)$", w)
    if m:
        try:
            return int(m.group(1)) >= 600
        except Exception:
            return False
    return False


def _rgb_from_svg(text: SvgText) -> RGBColor:
    rgb = _parse_color_hex(text.fill, default=(0, 0, 0))
    rgb2 = _blend_with_white(rgb, text.fill_opacity)
    return RGBColor(rgb2[0], rgb2[1], rgb2[2])


def extract_notes_text(pptx_zip: zipfile.ZipFile, slide_num: int) -> str | None:
    notes_path = f"ppt/notesSlides/notesSlide{slide_num}.xml"
    try:
        raw = pptx_zip.read(notes_path)
    except KeyError:
        return None

    try:
        root = ET.fromstring(raw)
    except Exception:
        return None

    # Extract text paragraphs.
    # notesSlide uses drawingml a:p/a:r/a:t for runs.
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    paras: list[str] = []
    for p in root.findall(".//a:p", ns):
        parts = [t.text or "" for t in p.findall(".//a:t", ns)]
        s = "".join(parts).strip()
        if s:
            paras.append(s)

    if not paras:
        return None
    return "\n".join(paras)


def build_editable_pptx(
    input_pptx: Path,
    output_pptx: Path,
    scale: float,
    max_slides: int | None,
) -> None:
    with zipfile.ZipFile(input_pptx) as z:
        slide_xmls = [n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)]
        slide_count = len(slide_xmls)

        # Use embedded SVGs in order to avoid filename ordering issues.
        svg_members = []
        for i in range(1, slide_count + 1):
            name = f"ppt/media/image{i}.svg"
            try:
                z.getinfo(name)
            except KeyError:
                # Some decks can skip numbers; ignore.
                continue
            svg_members.append((i, name))

        if max_slides is not None:
            svg_members = svg_members[: max_slides]

        prs = Presentation()
        prs.slide_width = _px_to_emu(1280)
        prs.slide_height = _px_to_emu(720)

        blank_layout = prs.slide_layouts[6]

        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            for slide_num, member in svg_members:
                svg_bytes = z.read(member)
                svg_root = ET.fromstring(svg_bytes)

                texts = extract_svg_texts(svg_root)

                # Build background SVG by removing text.
                remove_all_text_elements(svg_root)
                bg_svg_path = tmp / f"bg_{slide_num:02d}.svg"
                bg_svg_path.write_text(ET.tostring(svg_root, encoding="unicode"), encoding="utf-8")

                bg_png_path = tmp / f"bg_{slide_num:02d}.png"
                render_svg_to_png(
                    bg_svg_path,
                    bg_png_path,
                    width_px=int(round(1280 * scale)),
                    height_px=int(round(720 * scale)),
                )

                slide = prs.slides.add_slide(blank_layout)

                # Background image.
                slide.shapes.add_picture(
                    str(bg_png_path),
                    Emu(0),
                    Emu(0),
                    width=Emu(prs.slide_width),
                    height=Emu(prs.slide_height),
                )

                # Speaker notes (best-effort).
                notes = extract_notes_text(z, slide_num)
                if notes:
                    try:
                        slide.notes_slide.notes_text_frame.text = notes
                    except Exception:
                        pass

                # Overlay editable text.
                for t in texts:
                    font_pt = max(6.0, _px_to_pt(t.font_px))
                    # Approximate baseline -> top.
                    top_px = t.y - 0.8 * t.font_px

                    est_w_px = _estimate_text_width_px(t.text, t.font_px)
                    if t.anchor == "middle":
                        left_px = t.x - est_w_px / 2.0
                    elif t.anchor == "end":
                        left_px = t.x - est_w_px
                    else:
                        left_px = t.x

                    # Clamp within canvas.
                    left_px = max(0.0, min(1280.0 - 2.0, left_px))
                    top_px = max(0.0, min(720.0 - 2.0, top_px))

                    # Avoid wrapping: give a generous width for safety.
                    w_px = min(1280.0 - left_px, max(est_w_px, 20.0))
                    h_px = max(t.font_px * 1.5, 10.0)

                    shape = slide.shapes.add_textbox(
                        Emu(_px_to_emu(left_px)),
                        Emu(_px_to_emu(top_px)),
                        Emu(_px_to_emu(w_px)),
                        Emu(_px_to_emu(h_px)),
                    )

                    # Transparent box.
                    try:
                        shape.fill.background()
                        shape.line.fill.background()
                    except Exception:
                        pass

                    tf = shape.text_frame
                    tf.clear()
                    tf.word_wrap = False
                    try:
                        tf.margin_left = 0
                        tf.margin_right = 0
                        tf.margin_top = 0
                        tf.margin_bottom = 0
                    except Exception:
                        pass

                    p = tf.paragraphs[0]
                    p.alignment = _ppt_align_from_anchor(t.anchor)
                    run = p.add_run()
                    run.text = t.text

                    font = run.font
                    font.size = Pt(font_pt)

                    # Pick first font in stack; fall back to a CN/EN reasonable default.
                    name = _first_font_name(t.font_family)
                    if not name:
                        name = "Source Han Sans SC" if _is_cjk(t.text) else "Montserrat"
                    font.name = name
                    font.color.rgb = _rgb_from_svg(t)
                    font.bold = _bold_from_weight(t.font_weight)

        output_pptx.parent.mkdir(parents=True, exist_ok=True)
        prs.save(str(output_pptx))


def main() -> None:
    parser = argparse.ArgumentParser(description="Rebuild PPTX with editable text extracted from embedded SVGs")
    parser.add_argument("input_pptx", type=Path)
    parser.add_argument("-o", "--output", type=Path, required=True)
    parser.add_argument("--scale", type=float, default=2.0, help="Background render scale (default 2.0)")
    parser.add_argument("--max-slides", type=int, default=None, help="Limit slides for a quick dry run")
    args = parser.parse_args()

    if not args.input_pptx.exists():
        raise SystemExit(f"Input not found: {args.input_pptx}")

    build_editable_pptx(
        input_pptx=args.input_pptx,
        output_pptx=args.output,
        scale=args.scale,
        max_slides=args.max_slides,
    )


if __name__ == "__main__":
    main()
