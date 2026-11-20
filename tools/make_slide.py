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

Since 20 November 2026 the slide is filled into tools/slide_template.pptx, Ton's
layout on the EPAM session template: gradient title with a headphones icon, the
app screenshot top right, seven topic rows, the QR card on the right, and the EPAM
master footer (logo, "EPAM Proprietary & Confidential.", page number). Only the
week, the seven lines, the call to action and the speaker notes change. The QR
code in the template points at the app URL below. The script refuses text that is
too long, so the caller shortens the wording instead of the font.
"""
import json
import os
import sys

from pptx import Presentation

TEMPLATE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "slide_template.pptx")
APP_URL = "https://tonvlot1969.github.io/ai-briefing/ai-briefing.html"

TAGS = ["General AI", "EPAM", "Governance", "Engineering", "BI", "Architecture", "Delivery"]
LIMITS = {"week": 34, "text": 56, "cta": 52}

# shape names in the template
TITLE = "Title 1"
TAG_BOXES = ["TextBox 6", "TextBox 10", "TextBox 14", "TextBox 18", "TextBox 22", "TextBox 26", "TextBox 30"]
LINE_BOXES = ["TextBox 7", "TextBox 11", "TextBox 15", "TextBox 19", "TextBox 23", "TextBox 27", "TextBox 31"]
CTA = "TextBox 32"


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
    if spec.get("url") != APP_URL:
        errs.append(f"url must be {APP_URL} (the template's QR code points there)")
    if "—" in json.dumps(spec, ensure_ascii=False):
        errs.append("em-dash found, use a period or a comma")
    if errs:
        sys.exit("SLIDE SPEC REJECTED\n" + "\n".join(errs))


def set_text(shape, text):
    """Replace a text box's words, keeping the first run's formatting."""
    p = shape.text_frame.paragraphs[0]
    runs = p.runs
    runs[0].text = text
    for r in runs[1:]:
        r._r.getparent().remove(r._r)


def main(spec_path, out_path):
    spec = json.load(open(spec_path, encoding="utf-8"))
    check(spec)

    prs = Presentation(TEMPLATE)
    s = prs.slides[0]
    shapes = {sh.name: sh for sh in s.shapes}

    set_text(shapes[TITLE], "AI Briefing  ·  " + spec["week"])
    for i, it in enumerate(spec["items"]):
        tag = shapes[TAG_BOXES[i]].text_frame.text.strip()
        if tag.upper() != it["tag"].upper():
            sys.exit(f"template row {i + 1} is {tag!r}, expected {it['tag']!r}")
        set_text(shapes[LINE_BOXES[i]], it["text"])
    # the first row's text box sits a little lower in the template: align it with the others
    g, ref = shapes[LINE_BOXES[0]], shapes[LINE_BOXES[1]]
    g.top, g.height = shapes[TAG_BOXES[0]].top, ref.height
    g.text_frame._txBody.bodyPr.set("anchor", ref.text_frame._txBody.bodyPr.get("anchor") or "ctr")
    set_text(shapes[CTA], spec["cta"])

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
