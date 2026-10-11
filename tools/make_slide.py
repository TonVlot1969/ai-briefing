#!/usr/bin/env python3
"""Render the weekly 'The Best AI is Yet to Come!' teaser slide for the AI Briefing.

Usage: python3 tools/make_slide.py spec.json out.pptx

spec.json:
{
  "week": "Week of 11 October 2026",                       (max 34 chars)
  "items": [                                                 (exactly 7, in this order)
    {"tag": "General AI", "text": "one line, max 56 chars", "note": "2 to 4 sentences for the speaker notes, incl. date and source"},
    {"tag": "EPAM", ...}, {"tag": "Governance", ...}, {"tag": "Engineering", ...},
    {"tag": "BI", ...}, {"tag": "Architecture", ...}, {"tag": "Delivery", ...}
  ],
  "cta": "Pick your role, press Play, listen on the way",   (max 52 chars)
  "url": "https://tonvlot1969.github.io/ai-briefing/ai-briefing.html"
}

One EPAM-branded content slide in the style of the session deck: Snow background,
Night header bar, Calibri, Cobalt and Slate accents, one line per topic, QR code to
the app, footer "EPAM Proprietary & Confidential.", and the explanation of every
line in the speaker notes. The script refuses text that is too long, so the caller
shortens the wording instead of the font.
"""
import io
import json
import sys

import qrcode
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

NIGHT = RGBColor(0x06, 0x06, 0x06)
SNOW = RGBColor(0xFB, 0xFA, 0xFA)
SEA = RGBColor(0x00, 0xF6, 0xFF)
COBALT = RGBColor(0x00, 0x47, 0xFF)
SLATE = RGBColor(0x4A, 0x71, 0xBD)
TINT = RGBColor(0xE9, 0xEF, 0xFF)

TAGS = ["General AI", "EPAM", "Governance", "Engineering", "BI", "Architecture", "Delivery"]
LIMITS = {"week": 34, "text": 56, "cta": 52}


def check(spec):
    errs = []
    for k in ("week", "cta"):
        if len(spec.get(k, "")) > LIMITS[k]:
            errs.append(f"{k} is {len(spec[k])} chars, max {LIMITS[k]}: {spec[k]!r}")
    items = spec.get("items", [])
    if [i.get("tag") for i in items] != TAGS:
        errs.append(f"items must be exactly these tags in this order: {TAGS}")
    for i in items:
        if len(i.get("text", "")) > LIMITS["text"]:
            errs.append(f"line too long ({len(i['text'])}>{LIMITS['text']}): {i['text']!r}")
        if not i.get("note", "").strip():
            errs.append(f"missing note for {i.get('tag')}")
    if "—" in json.dumps(spec, ensure_ascii=False):
        errs.append("em-dash found, use a period or a comma")
    if errs:
        sys.exit("SLIDE SPEC REJECTED\n" + "\n".join(errs))


def box(slide, x, y, w, h, text, size, color, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0)
    tf.margin_top = tf.margin_bottom = Inches(0)
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = "Calibri"
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    return tb


def rect(slide, x, y, w, h, fill, shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    s.shadow.inherit = False
    return s


def main(spec_path, out_path):
    spec = json.load(open(spec_path, encoding="utf-8"))
    check(spec)

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.background.fill
    bg.solid()
    bg.fore_color.rgb = SNOW

    # header bar
    rect(s, 0, 0, 13.333, 1.0, NIGHT)
    box(s, 0.6, 0.15, 8.4, 0.7, "AI Briefing  ·  " + spec["week"], 26, SNOW, bold=True)
    box(s, 8.9, 0.15, 3.85, 0.7, "The Best AI is Yet to Come!", 16, SEA, bold=True, align=PP_ALIGN.RIGHT)

    # seven one-liners, one per topic
    y, step, h = 1.3, 0.66, 0.56
    for it in spec["items"]:
        rect(s, 0.6, y, 8.75, h, TINT, MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.2
        rect(s, 0.6, y, 0.1, h, COBALT)
        box(s, 0.85, y, 1.75, h, it["tag"].upper(), 13, SLATE, bold=True)
        box(s, 2.6, y, 6.65, h, it["text"], 18, NIGHT)
        y += step

    # call to action
    box(s, 0.6, y + 0.12, 8.75, 0.5, spec["cta"], 20, COBALT, bold=True)

    # QR card on the right
    rect(s, 9.75, 1.3, 3.0, 5.15, TINT, MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.06
    qr = qrcode.QRCode(border=1, box_size=12, error_correction=qrcode.constants.ERROR_CORRECT_M)
    qr.add_data(spec["url"])
    qr.make(fit=True)
    img = qr.make_image(fill_color="#060606", back_color="#FBFAFA")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    s.shapes.add_picture(buf, Inches(9.95), Inches(1.5), Inches(2.6), Inches(2.6))
    box(s, 9.85, 4.3, 2.8, 0.5, "Scan and listen", 22, COBALT, bold=True, align=PP_ALIGN.CENTER)
    box(s, 9.85, 4.85, 2.8, 0.4, "NL or EN, 7 topics", 15, SLATE, align=PP_ALIGN.CENTER)
    box(s, 9.85, 5.25, 2.8, 0.4, "pick your role", 15, SLATE, align=PP_ALIGN.CENTER)
    box(s, 9.85, 5.85, 2.8, 0.4, "tonvlot1969.github.io/ai-briefing", 11, SLATE, align=PP_ALIGN.CENTER)

    # footer, as in the session deck
    box(s, 0.6, 6.95, 6.0, 0.3, "EPAM Proprietary & Confidential.", 10, SLATE)

    # speaker notes: the explanation behind every line
    lines = [f"AI Briefing, {spec['week']}. One line per topic, explained below.", ""]
    for it in spec["items"]:
        lines.append(f"{it['tag'].upper()}: {it['text']}")
        lines.append(it["note"].strip())
        lines.append("")
    lines.append("Close: " + spec["cta"] + ". " + spec["url"])
    s.notes_slide.notes_text_frame.text = "\n".join(lines)

    prs.save(out_path)
    print("saved", out_path)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2])
