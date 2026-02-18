#!/usr/bin/env python3
"""Generate custom dark-mode visuals for the WolfXL Twitter thread.

Usage:
    uv run python scripts/generate_twitter_visuals.py [--output-dir twitter_visuals]

Produces 5 PNGs sized for Twitter (1200x675, 16:9):
  1. tradeoff_table.png   — Speed/fidelity/caps comparison
  2. feedback_loop.png    — AI agent development cycle
  3. speedup_card.png     — The multiplier numbers (hero image)
  4. architecture.png     — Rust/Python layer stack
  5. code_snippet.png     — Quick-start code
"""

from __future__ import annotations

import argparse
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

# ── Dimensions ────────────────────────────────────────────────────
W, H = 1200, 675  # 16:9

# ── Colors ────────────────────────────────────────────────────────
BG = "#0a0a0a"
CARD = "#161616"
CARD_BORDER = "#2a2a2a"
TEXT = "#ededed"
TEXT2 = "#a0a0a0"
ORANGE = "#f97316"
ORANGE_DIM = "#7c3a0d"
GREEN = "#62c073"
RED = "#ff6066"
BLUE = "#51a8ff"
YELLOW = "#fbbf24"

# ── Fonts ─────────────────────────────────────────────────────────
_SF = "/System/Library/Fonts/SFNS.ttf"
_SF_BOLD = "/System/Library/Fonts/SFNS.ttf"  # SF has weight axes; we fake bold via size
_MONO = "/System/Library/Fonts/Menlo.ttc"
_ARIAL = "/System/Library/Fonts/Supplemental/Arial Bold.ttf"
_ARIAL_REG = "/System/Library/Fonts/Supplemental/Arial.ttf"


def _font(size: int, bold: bool = False, mono: bool = False) -> ImageFont.FreeTypeFont:
    if mono:
        return ImageFont.truetype(_MONO, size)
    path = _ARIAL if bold else _ARIAL_REG
    return ImageFont.truetype(path, size)


def _card_rect(
    draw: ImageDraw.ImageDraw,
    x0: int, y0: int, x1: int, y1: int,
    fill: str = CARD, border: str = CARD_BORDER, radius: int = 16,
) -> None:
    draw.rounded_rectangle([x0, y0, x1, y1], radius=radius, fill=fill, outline=border, width=1)


def _center_text(
    draw: ImageDraw.ImageDraw,
    text: str,
    y: int,
    font: ImageFont.FreeTypeFont,
    fill: str = TEXT,
    x_center: int = W // 2,
) -> None:
    bbox = draw.textbbox((0, 0), text, font=font)
    tw = bbox[2] - bbox[0]
    draw.text((x_center - tw // 2, y), text, font=font, fill=fill)


# ====================================================================
#  1. Tradeoff Table
# ====================================================================


def generate_tradeoff_table(out: Path) -> None:
    img = Image.new("RGB", (W, H), BG)
    draw = ImageDraw.Draw(img)

    # Title
    _center_text(draw, "The Python Excel Tradeoff", 30, _font(36, bold=True), TEXT2)

    # Table area
    table_x = 80
    table_y = 100
    col_widths = [280, 240, 240, 280]  # Library, Read Speed, Fidelity, Caps
    row_h = 72
    header_h = 56

    headers = ["Library", "Read Speed", "Fidelity", "Capabilities"]
    rows = [
        ("openpyxl",        "811 cells/s",  "16/16",    "R + W"),
        ("calamine",        "Fast",         "2/16",     "R only"),
        ("xlsxwriter",      "—",            "15/16",    "W only"),
        ("WolfXL",          "5K cells/s",   "14/16",    "R + W + Patch"),
    ]

    # Draw header
    x = table_x
    for i, hdr in enumerate(headers):
        draw.text((x + 16, table_y + 14), hdr, font=_font(20, bold=True), fill=TEXT2)
        x += col_widths[i]

    # Separator
    draw.line(
        [(table_x, table_y + header_h), (table_x + sum(col_widths), table_y + header_h)],
        fill=CARD_BORDER, width=1,
    )

    # Rows
    for row_idx, (lib, speed, fid, caps) in enumerate(rows):
        ry = table_y + header_h + row_idx * row_h + 8
        is_wolf = lib == "WolfXL"

        if is_wolf:
            # Highlight row
            _card_rect(
                draw,
                table_x - 10, ry - 8,
                table_x + sum(col_widths) + 10, ry + row_h - 12,
                fill="#1a0800", border=ORANGE, radius=10,
            )

        x = table_x
        lib_color = ORANGE if is_wolf else TEXT
        lib_font = _font(24, bold=True) if is_wolf else _font(22)
        draw.text((x + 16, ry + 14), lib, font=lib_font, fill=lib_color)
        x += col_widths[0]

        speed_color = GREEN if "5K" in speed else (RED if speed == "—" else TEXT2)
        draw.text((x + 16, ry + 14), speed, font=_font(22), fill=speed_color)
        x += col_widths[1]

        fid_val = int(fid.split("/")[0])
        fid_color = GREEN if fid_val >= 14 else (YELLOW if fid_val >= 10 else RED)
        draw.text((x + 16, ry + 14), fid, font=_font(22), fill=fid_color)
        x += col_widths[2]

        caps_color = ORANGE if is_wolf else TEXT2
        draw.text((x + 16, ry + 14), caps, font=_font(22), fill=caps_color)

    # Footnote
    draw.text(
        (table_x, H - 50),
        "Fidelity = features at full score (3/3). Speed = cell read throughput (p50).",
        font=_font(15), fill="#666666",
    )

    img.save(out, "PNG")


# ====================================================================
#  2. Feedback Loop
# ====================================================================


def generate_feedback_loop(out: Path) -> None:
    img = Image.new("RGB", (W, H), BG)
    draw = ImageDraw.Draw(img)

    _center_text(draw, "AI-Driven Development with ExcelBench", 28, _font(32, bold=True), TEXT2)

    # Boxes in a cycle
    boxes = [
        (100, 220, "Excel Fixtures", "Ground truth\n(.xlsx from Excel)", TEXT2),
        (420, 220, "ExcelBench", "17 features\n238 test cases", BLUE),
        (740, 220, "Score: 0-3", "Binary pass/fail\nper test case", GREEN),
        (740, 420, "AI Agent", "Claude Code\nwrites adapter", ORANGE),
        (420, 420, "Run Tests", "Automated\nvalidation", YELLOW),
        (100, 420, "Iterate", "Until all\nfeatures = 3", TEXT2),
    ]

    for bx, by, title, sub, color in boxes:
        _card_rect(draw, bx, by, bx + 260, by + 120, fill="#111111", border=color, radius=12)
        draw.text((bx + 20, by + 16), title, font=_font(20, bold=True), fill=color)
        draw.text((bx + 20, by + 48), sub, font=_font(15), fill=TEXT2)

    # Arrows (right across top row, down, left across bottom row, up)
    arrow_color = "#444444"
    aw = 2
    # Top row: right arrows
    draw.line([(360, 280), (420, 280)], fill=arrow_color, width=aw)
    draw.line([(680, 280), (740, 280)], fill=arrow_color, width=aw)
    # Down on right
    draw.line([(870, 340), (870, 420)], fill=arrow_color, width=aw)
    # Bottom row: left arrows
    draw.line([(740, 480), (680, 480)], fill=arrow_color, width=aw)
    draw.line([(420, 480), (360, 480)], fill=arrow_color, width=aw)
    # Up on left
    draw.line([(230, 420), (230, 340)], fill=arrow_color, width=aw)

    # Arrow heads (simple triangles)
    for ax, ay, direction in [
        (420, 280, "right"), (740, 280, "right"),
        (870, 420, "down"),
        (680, 480, "left"), (360, 480, "left"),
        (230, 340, "up"),
    ]:
        s = 8
        if direction == "right":
            draw.polygon([(ax, ay - s), (ax + s, ay), (ax, ay + s)], fill=arrow_color)
        elif direction == "left":
            draw.polygon([(ax, ay - s), (ax - s, ay), (ax, ay + s)], fill=arrow_color)
        elif direction == "down":
            draw.polygon([(ax - s, ay), (ax, ay + s), (ax + s, ay)], fill=arrow_color)
        elif direction == "up":
            draw.polygon([(ax - s, ay), (ax, ay - s), (ax + s, ay)], fill=arrow_color)

    # Bottom label
    _center_text(draw, '"Done" is precisely defined — no ambiguity, no manual review', H - 55,
                 _font(18), fill="#666666")

    img.save(out, "PNG")


# ====================================================================
#  3. Speedup Card (hero)
# ====================================================================


def generate_speedup_card(out: Path) -> None:
    img = Image.new("RGB", (W, H), BG)
    draw = ImageDraw.Draw(img)

    # Wolf icon area
    _center_text(draw, "WolfXL", 30, _font(28, bold=True), ORANGE)

    # The three big numbers
    metrics = [
        ("6x",    "faster reads"),
        ("4x",    "faster writes"),
        ("10-14x","faster modify"),
    ]

    card_w = 300
    gap = 50
    total = card_w * 3 + gap * 2
    start_x = (W - total) // 2

    for i, (number, label) in enumerate(metrics):
        cx = start_x + i * (card_w + gap)
        cy = 160

        _card_rect(draw, cx, cy, cx + card_w, cy + 340, fill="#111111", border=CARD_BORDER)

        # Big number
        num_font = _font(96, bold=True)
        bbox = draw.textbbox((0, 0), number, font=num_font)
        nw = bbox[2] - bbox[0]
        draw.text((cx + (card_w - nw) // 2, cy + 60), number, font=num_font, fill=ORANGE)

        # Label
        lbl_font = _font(24)
        bbox = draw.textbbox((0, 0), label, font=lbl_font)
        lw = bbox[2] - bbox[0]
        draw.text((cx + (card_w - lw) // 2, cy + 220), label, font=lbl_font, fill=TEXT)

    # Baseline
    _center_text(draw, "vs openpyxl  •  Hybrid Rust + Python", H - 65, _font(20), TEXT2)
    _center_text(draw, "No fidelity tradeoff — 14/16 features at full score", H - 35,
                 _font(17), fill="#666666")

    img.save(out, "PNG")


# ====================================================================
#  4. Architecture Stack
# ====================================================================


def generate_architecture(out: Path) -> None:
    img = Image.new("RGB", (W, H), BG)
    draw = ImageDraw.Draw(img)

    _center_text(draw, "WolfXL Architecture", 25, _font(30, bold=True), TEXT2)

    # Layer boxes (top to bottom)
    lx = 160
    lw = W - 320
    layers = [
        (95,  "Python API", "load_workbook()  •  .save()  •  Worksheet['A1']", BLUE, 80),
        (205, "PyO3 Bridge", "Zero-copy data transfer between Python ↔ Rust", TEXT2, 65),
        (310, None, None, None, None),  # Split into 3 sub-boxes
    ]

    for ly, title, sub, color, lh in layers:
        if title is None:
            continue
        _card_rect(draw, lx, ly, lx + lw, ly + lh, fill="#111111", border=color, radius=12)
        draw.text((lx + 24, ly + 14), title, font=_font(24, bold=True), fill=color)
        if sub:
            draw.text((lx + 24, ly + 46), sub, font=_font(16), fill=TEXT2)

    # Rust layer — 3 boxes side by side
    rust_y = 310
    rust_h = 160
    box_w = (lw - 40) // 3
    rust_boxes = [
        ("calamine", "Read Engine", "Styled cell parsing\nFormat extraction", GREEN),
        ("rust_xlsxwriter", "Write Engine", "Full-featured .xlsx\ngeneration", ORANGE),
        ("XlsxPatcher", "Modify Engine", "Surgical ZIP patching\n10-14x vs rewrite", YELLOW),
    ]

    for i, (name, role, desc, color) in enumerate(rust_boxes):
        bx = lx + i * (box_w + 20)
        _card_rect(draw, bx, rust_y, bx + box_w, rust_y + rust_h, fill="#111111",
                   border=color, radius=12)
        draw.text((bx + 16, rust_y + 14), name, font=_font(20, bold=True), fill=color)
        draw.text((bx + 16, rust_y + 42), role, font=_font(15), fill=TEXT)
        draw.text((bx + 16, rust_y + 80), desc, font=_font(14), fill=TEXT2)

    # Rust badge
    _card_rect(draw, lx, rust_y + rust_h + 18, lx + lw, rust_y + rust_h + 55,
               fill="#1a0800", border=ORANGE_DIM, radius=8)
    _center_text(draw, "All Rust layers compiled to native — zero Python overhead in hot paths",
                 rust_y + rust_h + 26, _font(16), ORANGE)

    # Connection lines
    for from_y, to_y in [(175, 205), (270, 310)]:
        mid = W // 2
        draw.line([(mid, from_y), (mid, to_y)], fill="#444444", width=2)

    # Footnote
    _center_text(draw, "pip install wolfxl  •  Pre-built wheels for macOS / Linux / Windows",
                 H - 40, _font(16), "#666666")

    img.save(out, "PNG")


# ====================================================================
#  5. Code Snippet
# ====================================================================


def generate_code_snippet(out: Path) -> None:
    img = Image.new("RGB", (W, H), BG)
    draw = ImageDraw.Draw(img)

    _center_text(draw, "Get started in 4 lines", 30, _font(30, bold=True), TEXT2)

    # Code card
    cx, cy = 120, 100
    cw, ch = W - 240, 420
    _card_rect(draw, cx, cy, cx + cw, cy + ch, fill="#111111", border=CARD_BORDER, radius=14)

    # Window dots
    for i, dot_color in enumerate(["#ff5f57", "#febc2e", "#28c840"]):
        draw.ellipse([cx + 20 + i * 22, cy + 16, cx + 32 + i * 22, cy + 28], fill=dot_color)

    # Code lines
    code_font = _font(24, mono=True)
    line_h = 38
    code_x = cx + 30
    code_y = cy + 55

    lines: list[list[tuple[str, str]]] = [
        [("from", "#c678dd"), (" wolfxl ", TEXT),
         ("import", "#c678dd"), (" load_workbook", YELLOW)],
        [],
        [("wb ", TEXT), ("= ", "#c678dd"), ("load_workbook", YELLOW), ("(", TEXT),
         ('"report.xlsx"', GREEN), (")", TEXT)],
        [("ws ", TEXT), ("= ", "#c678dd"), ("wb", TEXT), (".", TEXT2), ("active", BLUE)],
        [],
        [("# Read, modify, save — full fidelity", "#666666")],
        [("ws", TEXT), ("[", TEXT2), ('"A1"', GREEN), ("]", TEXT2),
         (".", TEXT2), ("value ", BLUE), ("= ", "#c678dd"), ('"Updated"', GREEN)],
        [("wb", TEXT), (".", TEXT2), ("save", YELLOW), ("(", TEXT),
         ('"report.xlsx"', GREEN), (")", TEXT)],
    ]

    for i, line_parts in enumerate(lines):
        x = code_x
        y = code_y + i * line_h
        for text, color in line_parts:
            draw.text((x, y), text, font=code_font, fill=color)
            bbox = draw.textbbox((0, 0), text, font=code_font)
            x += bbox[2] - bbox[0]

    # Bottom labels
    _center_text(draw, "pip install wolfxl", H - 85, _font(24, bold=True), ORANGE)
    _center_text(draw, "github.com/SynthGL/wolfxl  •  excelbench.vercel.app",
                 H - 48, _font(18), TEXT2)

    img.save(out, "PNG")


# ====================================================================
#  Main
# ====================================================================


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate Twitter thread visuals")
    parser.add_argument("--output-dir", default="twitter_visuals", help="Output directory")
    args = parser.parse_args()

    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    generators = [
        ("tradeoff_table.png", generate_tradeoff_table),
        ("feedback_loop.png", generate_feedback_loop),
        ("speedup_card.png", generate_speedup_card),
        ("architecture.png", generate_architecture),
        ("code_snippet.png", generate_code_snippet),
    ]

    for name, func in generators:
        path = out_dir / name
        func(path)
        print(f"  ✓ {path}")

    print(f"\nDone — {len(generators)} visuals in {out_dir}/")


if __name__ == "__main__":
    main()
