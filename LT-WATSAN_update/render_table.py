"""Draws a pipeline-summary table into a vector EMF - same layout as render_table.ps1.

Used by update_schematic.py when Windows PowerShell is not available (Linux/macOS).
Text is measured with Carlito, which has the same character widths as Calibri, and the
EMF asks for Calibri, so the table looks the same in Visio as the PowerShell version.
Needs:  pip install fonttools   and the Carlito font (e.g. apt install fonts-crosextra-carlito).
"""
import glob, os, struct
from fontTools.ttLib import TTFont

COLS = (40, 380, 208)                    # S/N | DESCRIPTION | PIPELINE LENGTH (m), in points
DEV_PX, DEV_MM = (1920, 1080), (344, 193)  # reference device written in the EMF header
PX_PER_PT = DEV_PX[0] / DEV_MM[0] * 25.4 / 72
U = 100                                  # logical units per point

FILL = {"title": 0xDDEBF7, "head": 0xBDD7EE, "grand": 0x9BC2E6}
BLACK, GREEN, RED, WHITE = 0x000000, 0x00B050, 0xFF0000, 0xFFFFFF


def _find_font(style):
    pats = [f"/usr/share/fonts/**/Carlito-{style}.ttf", os.path.expanduser(f"~/.fonts/**/Carlito-{style}.ttf"),
            f"/Library/Fonts/**/Carlito-{style}.ttf", os.path.join(os.path.dirname(__file__), f"Carlito-{style}.ttf")]
    for p in pats:
        hit = glob.glob(p, recursive=True)
        if hit:
            return hit[0]
    raise FileNotFoundError(f"Carlito-{style}.ttf not found (install fonts-crosextra-carlito)")


class Metrics:
    def __init__(self, style):
        f = TTFont(_find_font(style))
        self.upm = f["head"].unitsPerEm
        self.asc, self.desc = f["OS/2"].usWinAscent, f["OS/2"].usWinDescent
        self.cmap = f.getBestCmap()
        self.hmtx = f["hmtx"].metrics

    def advances(self, text, size):
        return [self.hmtx[self.cmap.get(ord(ch), ".notdef")][0] * size / self.upm for ch in text]

    def width(self, text, size):
        return sum(self.advances(text, size))


def rgb(c):  # 0xRRGGBB -> COLORREF 0x00BBGGRR
    return ((c >> 16) & 255) | (c & 0xFF00) | ((c & 255) << 16)


class Emf:
    def __init__(self):
        self.recs, self.handles = [], 1

    def rec(self, typ, payload=b""):
        self.recs.append(struct.pack("<II", typ, 8 + len(payload)) + payload)

    def handle(self):
        self.handles += 1
        return self.handles - 1

    def pen(self, color, width):
        h = self.handle()
        self.rec(38, struct.pack("<IIiiI", h, 0, int(width), 0, rgb(color)))
        return h

    def brush(self, color):
        h = self.handle()
        self.rec(39, struct.pack("<IIII", h, 0, rgb(color), 0))
        return h

    def font(self, size_pt, bold):
        h = self.handle()
        face = "Calibri".encode("utf-16-le").ljust(64, b"\0")
        lf = struct.pack("<iiiii8B", -round(size_pt * U), 0, 0, 0, 700 if bold else 400,
                         0, 0, 0, 1, 0, 0, 5, 0) + face
        self.rec(82, struct.pack("<I", h) + lf)
        return h

    def select(self, h):
        self.rec(37, struct.pack("<I", h))

    def delete(self, h):
        self.rec(40, struct.pack("<I", h))

    def rect(self, x0, y0, x1, y1):
        self.rec(43, struct.pack("<4i", *(round(v * U) for v in (x0, y0, x1, y1))))

    def text(self, x, y, s, dx):
        s16 = s.encode("utf-16-le")
        s16 += b"\0" * (-len(s16) % 4)
        off_str = 8 + 16 + 4 + 8 + 40            # record hdr, bounds, mode, scales, EMRTEXT
        off_dx = off_str + len(s16)
        emrtext = struct.pack("<iiIIIiiiiI", round(x * U), round(y * U), len(s), off_str, 0, 0, 0, -1, -1, off_dx)
        body = struct.pack("<4iIff", 0, 0, -1, -1, 1, 0.0, 0.0) + emrtext + s16
        body += struct.pack(f"<{len(dx)}i", *dx)
        self.rec(84, body)

    def build(self, w_pt, h_pt):
        self.rec(14, struct.pack("<III", 0, 16, 20))
        px = (round(w_pt * PX_PER_PT), round(h_pt * PX_PER_PT))
        frame = (0, 0, round(w_pt * 2540 / 72), round(h_pt * 2540 / 72))
        size = 108 + sum(len(r) for r in self.recs)
        hdr = struct.pack("<II4i4iIIIIHHIIIiiiiIIIii", 1, 108, 0, 0, px[0], px[1], *frame,
                          0x464D4520, 0x10000, size, len(self.recs) + 1, self.handles, 0, 0, 0, 0,
                          DEV_PX[0], DEV_PX[1], DEV_MM[0], DEV_MM[1], 0, 0, 0,
                          DEV_MM[0] * 1000, DEV_MM[1] * 1000)
        assert len(hdr) == 108
        return hdr + b"".join(self.recs)


def render(rows, aspect=0):
    """rows: list of dicts as built by update_schematic.table_rows. Returns (emf_bytes, width_pt, height_pt, font_pt)."""
    W = sum(COLS)
    RH = round(W * aspect / len(rows), 2) if aspect > 0 else 13
    FS = round(RH * 0.75, 1)
    H = len(rows) * RH
    reg, bold = Metrics("Regular"), Metrics("Bold")

    # shrink the font until the longest description fits (GDI+ MeasureString adds 1/6 em padding)
    longest = [str(r["text"]) for r in rows if r["kind"] in ("section", "pipe")] + ["PIPELINE LENGTH (m)"]
    while FS > 5 and max(bold.width(t, FS) + FS / 6 for t in longest) > COLS[1] - 6:
        FS -= 0.25

    e = Emf()
    e.rec(17, struct.pack("<I", 8))                                   # MM_ANISOTROPIC
    e.rec(9, struct.pack("<ii", round(W * U), round(H * U)))          # window ext (logical = 1/100 pt)
    e.rec(11, struct.pack("<ii", round(W * PX_PER_PT), round(H * PX_PER_PT)))
    e.rec(18, struct.pack("<I", 1))                                   # transparent text background
    e.rec(22, struct.pack("<I", 0))                                   # TA_LEFT | TA_TOP
    pen = e.pen(BLACK, 0.5 * U)
    fonts = {False: e.font(FS, False), True: e.font(FS, True)}
    brushes = {k: e.brush(v) for k, v in FILL.items()}
    white = e.brush(WHITE)
    NULL_PEN, NULL_BRUSH = 0x80000008, 0x80000005

    e.select(NULL_PEN); e.select(white); e.rect(0, 0, W, H)
    e.select(pen)
    state = {"color": None, "font": None}

    def cell(x, y, w, text, is_bold, color, center, bg):
        e.select(brushes[bg] if bg else NULL_BRUSH)
        e.rect(x, y, x + w, y + RH)
        if not text:
            return
        m = bold if is_bold else reg
        text = str(text)
        adv = m.advances(text, FS)
        tw = sum(adv)
        tx = x + (w - tw) / 2 if center else x + 2 + FS / 12
        ty = y + (RH - (m.asc + m.desc) * FS / m.upm) / 2
        if state["font"] != is_bold:
            e.select(fonts[is_bold]); state["font"] = is_bold
        if state["color"] != color:
            e.rec(24, struct.pack("<I", rgb(color))); state["color"] = color
        dx, acc, prev = [], 0.0, 0
        for a in adv:                       # integer advances without rounding drift
            acc += a * U
            dx.append(round(acc) - prev); prev = round(acc)
        e.text(tx, ty, text, dx)

    x1, x2 = COLS[0], COLS[0] + COLS[1]
    y = 0
    for r in rows:
        k = r["kind"]
        if k == "title":
            cell(0, y, W, r["text"], True, BLACK, True, "title")
        elif k == "head":
            cell(0, y, COLS[0], "S/N", True, BLACK, True, "head")
            cell(x1, y, COLS[1], "DESCRIPTION", True, BLACK, True, "head")
            cell(x2, y, COLS[2], "PIPELINE LENGTH (m)", True, BLACK, True, "head")
        elif k == "section":
            cell(0, y, COLS[0], r["sn"], True, BLACK, True, None)
            cell(x1, y, COLS[1], r["text"], True, BLACK, False, None)
            cell(x2, y, COLS[2], "", False, BLACK, True, None)
        elif k == "pipe":
            cell(0, y, COLS[0], "", False, BLACK, True, None)
            cell(x1, y, COLS[1], r["text"], False, BLACK, False, None)
            cell(x2, y, COLS[2], r["len"], False, BLACK, True, None)
        elif k == "sub":
            cell(0, y, COLS[0], "", False, BLACK, True, None)
            cell(x1, y, COLS[1], "Sub-total", True, GREEN, True, None)
            cell(x2, y, COLS[2], r["len"], True, GREEN, True, None)
        elif k == "grand":
            cell(0, y, COLS[0], "", True, RED, True, "grand")
            cell(x1, y, COLS[1], r["text"], True, RED, True, "grand")
            cell(x2, y, COLS[2], r["len"], True, RED, True, "grand")
        else:
            for x, w in ((0, COLS[0]), (x1, COLS[1]), (x2, COLS[2])):
                cell(x, y, w, "", False, BLACK, True, None)
        y += RH
    e.select(NULL_PEN); e.select(NULL_BRUSH); e.select(0x8000000D)  # stock objects, so ours can be deleted
    for h in [pen, white, *fonts.values(), *brushes.values()]:
        e.delete(h)
    return e.build(W, H), W, H, FS
