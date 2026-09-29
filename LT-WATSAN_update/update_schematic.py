"""Push values from the LT-WATSAN project summary Excel into the Visio schematic labels.

Only the numbers inside existing labels are rewritten; Visio formatting (fonts,
superscript m3, positions) is left untouched. The original .vsdx is never modified.

Usage:  python update_schematic.py [--excel X.xlsx] [--vsdx IN.vsdx] [--out OUT.vsdx] [--pages "OPT 4A,OPT 4B"]
        python update_schematic.py --vsdx IN.vsdx --dump "OPT 4A"     (list shape IDs + text on a page)
"""
import argparse, csv, json, os, re, shutil, subprocess, sys, tempfile, zipfile
import xml.etree.ElementTree as ET
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
DL = os.path.dirname(HERE)
DEF_XLSX = os.path.join(DL, "LT-WATSAN PROJECT SUMMARY OPTIONS_OPT01_OPT02_OPT03_OPT04_OPT05_2030.xlsx")
DEF_VSDX = os.path.join(DL, "LT-WATSAN SCHEMATIC DIAGRAM_rev03.vsdx")
DEF_OUT = os.path.join(HERE, "LT-WATSAN SCHEMATIC DIAGRAM_rev03_UPDATED.vsdx")

T = r"(?:<[^>]*>)*"  # any Visio formatting tags (<cp IX='n'/>) sitting inside a value


def norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip().upper()


def fmt_int(n):
    return f"{int(round(n)):,}"


def num(s):
    return float(str(s).replace(",", ""))


# ---------------------------------------------------------------- Excel
class Sheet:
    def __init__(self, ws):
        self.rows = [(r[0].value, r[1].value, r[2].value, r[2].coordinate)
                     for r in ws.iter_rows(min_col=1, max_col=3)]

    def region(self, name):
        start = next((i for i, r in enumerate(self.rows) if norm(r[1]) == norm(name)), None)
        if start is None:
            raise KeyError(f"region '{name}' not found")
        end = next((i for i in range(start + 1, len(self.rows))
                    if norm(self.rows[i][1]).endswith(" OPTION")), len(self.rows))
        return start, end

    def anchor(self, reg, text):
        for i in range(*reg):
            if norm(self.rows[i][1]) == norm(text):
                return i
        raise KeyError(f"heading '{text}' not found")

    def pipe(self, reg, text):
        i = self.anchor(reg, text) + 1
        segs, subtotal = [], None
        while i < reg[1]:
            a, b, c, ref = self.rows[i]
            m = re.match(r"DI-DN([\d,]+)-PN(\d+)", norm(b))  # the Excel sometimes writes DN1,100
            if m and c is not None:
                segs.append((int(num(m[1])), int(m[2]), num(c), ref))
            elif norm(b) == "SUB-TOTAL":
                subtotal = num(c) if c is not None else None  # no cached value until Excel recalculates
                break
            elif segs or b is not None:
                break
            i += 1
        if not segs:
            raise KeyError(f"no DI-DN rows under '{text}'")
        return segs, subtotal

    def item(self, reg, text, anchor):
        i = self.anchor(reg, anchor)
        for j in range(i + 1, reg[1]):
            a, b, c, ref = self.rows[j]
            if isinstance(a, (int, float)):  # next numbered section
                break
            if norm(b) == norm(text):
                return str(c), ref
        raise KeyError(f"'{text}' not found under '{anchor}'")


def parse_pump(s):
    m = re.search(r"H\s*=\s*([\d.,]+)\s*m.*?\((\d+)\s*W\s*\+\s*(\d+)\s*S\)\s*@\s*([\d,.]+)\s*m3/h", s, re.I)
    if not m:
        raise ValueError(f"cannot read pump text '{s}'")
    v = re.search(r"([\d.]+)\s*m/s", s[m.end():])
    return dict(H=m[1], W=m[2], S=m[3], Q=fmt_int(num(m[4])), V=v[1] if v else None)


def parse_vol(s):
    m = re.search(r"([\d,]+)\s*m3", s)
    if not m:
        raise ValueError(f"cannot read volume '{s}'")
    return fmt_int(num(m[1]))


# ---------------------------------------------------------------- Visio text edits
def sub_nth(text, pattern, repl, n=0):
    """Replace the n-th match of pattern. Returns (text, old, new) or (text, None, None)."""
    ms = list(re.finditer(pattern, text))
    if len(ms) <= n:
        return text, None, None
    m = ms[n]
    new = repl(m)
    return text[:m.start()] + new + text[m.end():], m[0], new


def plain(s):
    return re.sub(r"<[^>]*>", "", s)


def build_edits(entry, sh, reg, default_hours=None):
    """Return list of (field, pattern, replacement_fn, excel_ref) plus notes."""
    t, edits, notes = entry["type"], [], []
    if t == "pipe":
        segs, subtotal = sh.pipe(reg, entry["anchor"])
        dns = sorted({s[0] for s in segs})
        pns = sorted({s[1] for s in segs})
        total = sum(s[2] for s in segs)
        if len(dns) > 1:
            notes.append(f"Excel has mixed diameters {dns} - label shows DN{dns[-1]}")
        if subtotal is not None and abs(subtotal - total) > 0.5:
            notes.append(f"Excel Sub-total {fmt_int(subtotal)} != sum of rows {fmt_int(total)} (used sum)")
        spec = f"DN{dns[-1]}-" + "-".join(f"PN{p}" for p in ([pns[0], pns[-1]] if len(pns) > 1 else pns))
        ref = f"{segs[0][3]}:{segs[-1][3]}"
        edits += [("DN/PN", r"DN[\d,]+(?:-PN\d+)+", lambda m, v=spec: v, ref),
                  ("Length", r"(L\s*=\s*)([\d,]+)(\s*m\b)", lambda m, v=fmt_int(total): m[1] + v + m[3], ref)]
    elif t == "tank":
        val, ref = sh.item(reg, entry["item"], entry["anchor"])
        edits.append(("Capacity", r"(CAPACITY: )([\d,]+)", lambda m, v=parse_vol(val): m[1] + v, ref))
    elif t == "station":
        if "daily" in entry:
            val, ref = sh.item(reg, entry["daily"], entry["anchor"])
            m = re.search(r"([\d,]+)\s*m3", val)
            edits.append(("Q m3/day", r"(Q = )([\d,]+)(\s*m" + T + "3" + T + "/day)",
                          lambda mm, v=fmt_int(num(m[1])): mm[1] + v + mm[3], ref))
        if "pump" in entry:
            val, ref = sh.item(reg, entry["pump"], entry["anchor"])
            p = parse_pump(val)
            edits += [("Head", r"(H = )([\d.,]+)(\s*m\b)", lambda m, v=p["H"]: m[1] + v + m[3], ref),
                      ("Pump Q", r"(H = [^<]*?, Q = )([\d,]+)", lambda m, v=p["Q"]: m[1] + v, ref),
                      ("Pumps", r"\(\d+w\+\d+s\)", lambda m, v=f"({p['W']}w+{p['S']}s)": v, ref),
]
            if p["V"]:
                edits.append(("Velocity", r"([Vv] = )([\d.]+)(\s*m/s)", lambda m, v=p["V"]: m[1] + v + m[3], ref))
            else:
                notes.append("pump velocity not in Excel - V kept")
        if "hours" in entry:
            val, ref = sh.item(reg, entry["hours"], entry["anchor"])
            h = re.search(r"(\d+)", val)[1]
            edits.append(("Hours", r"(\d+)( hrs)", lambda m, v=h: v + m[2], ref))
        elif default_hours:
            edits.append(("Hours", r"(\d+)( hrs)", lambda m, v=str(default_hours): v + m[2], "default"))
        else:
            notes.append("operating time not in Excel - hrs kept")
        if "sump" in entry:
            val, ref = sh.item(reg, entry["sump"], entry["anchor"])
            edits.append(("Sump", r"(Sump Capacity = )([\d,]+)", lambda m, v=parse_vol(val): m[1] + v, ref))
    return edits, notes


def page_files(z):
    ns = {"v": "http://schemas.microsoft.com/office/visio/2012/main",
          "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    rels = ET.fromstring(z.read("visio/pages/_rels/pages.xml.rels"))
    target = {r.get("Id"): r.get("Target") for r in rels}
    pages = ET.fromstring(z.read("visio/pages/pages.xml"))
    out = {}
    for p in pages.findall("v:Page", ns):
        f = "visio/pages/" + target[p.find("v:Rel", ns).get(f"{{{ns['r']}}}id")]
        out[p.get("Name")] = f
        out[p.get("NameU")] = f
    return out


def text_span(xml, sid):
    i = xml.find(f"<Shape ID='{sid}'")
    if i < 0:
        raise KeyError(f"shape {sid} not on page")
    j = xml.find("<Text>", i)
    k = xml.find("</Text>", j)
    nxt = min(x for x in (xml.find("<Shape ", i + 1), xml.find("</Shape>", i)) if x > 0)
    if j < 0 or j > nxt:
        raise KeyError(f"shape {sid} has no text")
    return j + 6, k


# ---------------------------------------------------------------- pipe summary table
def table_rows(sh, reg, t):
    """Every heading in the region that has DI-DN rows under it becomes a table section."""
    rows = [{"kind": "title", "text": "LT-WATSAN  PROJECT"},
            {"kind": "title", "text": t["title"]},
            {"kind": "title", "text": t["subtitle"]},
            {"kind": "head"}]
    sections, last, cur = [], None, None
    for a, b, c, ref in sh.rows[reg[0] + 1:reg[1]]:
        if re.match(r"DI-DN[\d,]+-PN\d+", norm(b)) and c is not None:
            if cur is None:
                cur = {"title": re.sub(r"^[A-Z]\.\s*", "", norm(last)), "pipes": []}
                sections.append(cur)
            cur["pipes"].append((norm(b), num(c)))
        else:
            cur = None
            if b is not None and norm(b) != "SUB-TOTAL":
                last = b
    overall = sum(l for s in sections for _, l in s["pipes"])
    first, last_sn = t.get("sections", [1, len(sections)])
    grand = 0
    for i, s in list(enumerate(sections, 1))[first - 1:last_sn]:
        rows.append({"kind": "section", "sn": str(i), "text": s["title"]})
        rows += [{"kind": "pipe", "text": p, "len": fmt_int(l)} for p, l in s["pipes"]]
        tot = sum(l for _, l in s["pipes"])
        grand += tot
        if len(s["pipes"]) > 1:
            rows.append({"kind": "sub", "len": fmt_int(tot)})
        rows.append({"kind": "blank"})
    rows.append({"kind": "grand", "text": t["grand"], "len": fmt_int(grand)})
    if t.get("overall"):
        rows.append({"kind": "grand", "text": t["overall"], "len": fmt_int(overall)})
    return rows, last_sn - first + 1, grand


def render_emf(rows, tmpdir, name, aspect):
    if not shutil.which("powershell"):  # not on Windows: same table drawn by the Python renderer
        import render_table
        emf, w, h, fs = render_table.render(rows, aspect)
        print(f"  {name} table: font {fs} pt")
        return emf, w, h
    jp, ep = os.path.join(tmpdir, name + ".json"), os.path.join(tmpdir, name + ".emf")
    json.dump(rows, open(jp, "w", encoding="utf-8"))
    out = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
                          os.path.join(HERE, "render_table.ps1"), jp, ep, str(aspect)],
                         capture_output=True, text=True, check=True).stdout.split()
    print(f"  {name} table: font {out[1]} pt")
    return open(ep, "rb").read(), float(out[-2]), float(out[-1])


def set_cell(shape, n, v):
    new, k = re.subn(rf"<Cell N='{n}' V='[^']*'(?: F='[^']*')?", f"<Cell N='{n}' V='{v:.12g}'", shape, count=1)
    return new


def shape_size(xml, sid):
    s = re.search(rf"<Shape ID='{sid}' Type='Foreign'.*?</Shape>", xml, re.S).group(0)
    c = dict(re.findall(r"<Cell N='(Width|Height)' V='([^']*)'", s))
    return float(c["Width"]), float(c["Height"])


def swap_table(xml, sid, rid, ew, eh):
    """Turn the linked Excel object into a plain EMF picture that fits the old box (top edge kept)."""
    m = re.search(rf"<Shape ID='{sid}' Type='Foreign'.*?</Shape>", xml, re.S)
    if not m:
        raise KeyError(f"table shape {sid} not found")
    s = m.group(0)
    c = {k: float(v) for k, v in re.findall(r"<Cell N='(PinX|PinY|Width|Height)' V='([^']*)'", s)}
    k = min(c["Width"] / ew, c["Height"] / eh)
    w, h = ew * k, eh * k
    top = c["PinY"] + c["Height"] / 2
    for n, v in (("Width", w), ("Height", h), ("PinY", top - h / 2), ("LocPinX", w / 2), ("LocPinY", h / 2),
                 ("ImgWidth", w), ("ImgHeight", h), ("ImgOffsetX", 0), ("ImgOffsetY", 0)):
        s = set_cell(s, n, v)
    s = re.sub(r"<ForeignData .*?</ForeignData>", f"<ForeignData ForeignType='EnhMetaFile'><Rel r:id='{rid}'/></ForeignData>", s, flags=re.S)
    return xml[:m.start()] + s + xml[m.end():]


def dump_page(vsdx, page):
    z = zipfile.ZipFile(vsdx)
    files = page_files(z)
    if page not in files:
        sys.exit(f"page '{page}' not found; pages: {sorted(set(files))}")
    xml = z.read(files[page]).decode("utf-8")
    for m in re.finditer(r"<Shape ID='(\d+)'", xml):
        try:
            s, e = text_span(xml, m[1])
        except KeyError:
            continue
        t = re.sub(r"\s+", " ", plain(xml[s:e])).strip()
        if t:
            print(f"{m[1]:>6}  {t}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", default=DEF_XLSX)
    ap.add_argument("--vsdx", default=DEF_VSDX)
    ap.add_argument("--out", default=DEF_OUT)
    ap.add_argument("--map", default=os.path.join(HERE, "mapping.json"))
    ap.add_argument("--pages", help="comma-separated page names to update (default: all pages in mapping)")
    ap.add_argument("--dump", metavar="PAGE", help="print every shape ID and its text on PAGE, then stop")
    a = ap.parse_args()

    if a.dump:
        return dump_page(a.vsdx, a.dump)
    cfg = json.load(open(a.map, encoding="utf-8"))
    if a.pages:
        want = [p.strip() for p in a.pages.split(",")]
        missing = [p for p in want if p not in cfg["pages"]]
        if missing:
            sys.exit(f"not in mapping: {missing}")
        cfg["pages"] = {p: cfg["pages"][p] for p in want}
    wb = openpyxl.load_workbook(a.excel, data_only=True)
    zin = zipfile.ZipFile(a.vsdx)
    files = page_files(zin)
    changed, added, log, problems = {}, {}, [], 0
    tmpdir = tempfile.mkdtemp(prefix="ltw_")

    for page, pc in cfg["pages"].items():
        if page not in files:
            problems += 1
            log.append([page, "", "ERROR", "", "", "", f"page '{page}' not in drawing (pages: {sorted(set(files))})"])
            continue
        sh = Sheet(wb[pc["sheet"]])
        reg = sh.region(pc["region"])
        fname = files[page]
        xml = changed.get(fname) or zin.read(fname).decode("utf-8")
        for key, entry in pc["labels"].items():
            sid, _, n = key.partition("#")
            n = int(n or 0)
            if not sid.isdigit():  # placeholder: shape ID still to be looked up with --dump
                log.append([page, key, "SKIPPED", "", "", "", f"shape ID not set in mapping ({entry.get('anchor')})"])
                continue
            try:
                edits, notes = build_edits(entry, sh, reg, cfg.get("default_hours"))
                s, e = text_span(xml, sid)
            except (KeyError, ValueError) as ex:
                problems += 1
                log.append([page, sid, "ERROR", "", "", "", str(ex)])
                continue
            txt = xml[s:e]
            for field, pat, fn, ref in edits:
                txt, old, new = sub_nth(txt, pat, fn, n)
                if old is None:
                    log.append([page, sid, field, "", "", ref, "field not found on label - skipped"])
                elif plain(old) != plain(new):
                    log.append([page, sid, field, plain(old), plain(new), ref, ""])
            if re.search(r"\bv = ", txt) and entry["type"] == "pipe":
                notes.append("pipe Q (m3/day) and v are not in Excel - kept as-is, please check")
            for nt in notes:
                log.append([page, sid, "NOTE", "", "", "", nt])
            xml = xml[:s] + txt + xml[e:]
        tables = pc.get("table", [])
        for t in tables if isinstance(tables, list) else [tables]:
            rows, nsec, grand = table_rows(sh, reg, t)
            base = os.path.basename(fname)[:-4]
            bw, bh = shape_size(xml, t["shape"])
            emf, ew, eh = render_emf(rows, tmpdir, f"{base}_{t['shape']}", bh / bw)
            media = f"visio/media/ltw_table_{base}_{t['shape']}.emf"
            added[media] = emf
            rels_name = f"visio/pages/_rels/{base}.xml.rels"
            rels = changed.get(rels_name) or zin.read(rels_name).decode("utf-8")
            rid = f"rIdLTW{t['shape']}"
            if f'Id="{rid}"' not in rels:
                rels = rels.replace("</Relationships>", f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/'
                                    f'officeDocument/2006/relationships/image" Target="../media/{os.path.basename(media)}"/></Relationships>')
            changed[rels_name] = rels
            xml = swap_table(xml, t["shape"], rid, ew, eh)
            log.append([page, t["shape"], "TABLE", "linked Excel table", f"{nsec} sections, grand total {fmt_int(grand)} m",
                        pc["sheet"], "rebuilt from Excel; old link to external pipe-summary file removed"])
        ET.fromstring(xml.encode("utf-8"))  # must still be valid XML
        changed[fname] = xml

    with zipfile.ZipFile(a.out, "w") as zout:
        for info in zin.infolist():
            if info.filename in added:  # table picture from an earlier run - replaced below
                continue
            data = changed[info.filename].encode("utf-8") if info.filename in changed else zin.read(info)
            zout.writestr(info, data, compress_type=info.compress_type)
        for name, data in added.items():
            zout.writestr(name, data, compress_type=zipfile.ZIP_DEFLATED)

    rep = os.path.splitext(a.out)[0] + "_CHANGES.csv"
    with open(rep, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["Page", "Shape ID", "Field", "Old (drawing)", "New (Excel)", "Excel cells", "Note"])
        w.writerows(log)

    n_chg = sum(1 for r in log if r[3] or r[4])
    print(f"{n_chg} values changed, {problems} errors")
    print(f"drawing : {a.out}")
    print(f"report  : {rep}")


if __name__ == "__main__":
    sys.exit(main())
