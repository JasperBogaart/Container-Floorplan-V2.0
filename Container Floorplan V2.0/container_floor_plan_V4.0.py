import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET
import re
import sys
import os
import subprocess
from datetime import datetime
from xml.sax.saxutils import escape

HERE = (Path(sys.executable).resolve().parent if getattr(sys, "frozen", False)
        else Path(__file__).resolve().parent)
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
TRANSPORTS = {
    "20FT General": (5.89, 2.35),
    "20FT Reefer": (5.44, 2.26),
    "40FT General": (12.05, 2.35),
    "40FT Reefer": (11.55, 2.26),
    "45FT Reefer": (13.30, 2.44),
    "Trailer": (13.60, 2.48),
    "Koel Trailer": (13.30, 2.46),
}
DEFAULT_TRANSPORT = "40FT General"
CONTAINER_L, CONTAINER_W = (value * 100 for value in TRANSPORTS[DEFAULT_TRANSPORT])

def col_index(ref):
    n = 0
    for c in re.match(r"[A-Z]+", ref).group():
        n = n * 26 + ord(c) - 64
    return n - 1

def read_xlsx(path):
    with ZipFile(path) as z:
        if "xl/workbook.xml" not in z.namelist():
            raise ValueError("Dit bestand is geen echt XLSX-bestand.")
        ss = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            ss = ["".join(t.text or "" for t in si.findall(".//m:t", NS))
                  for si in root.findall("m:si", NS)]
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        relroot = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rels = {x.attrib["Id"]: x.attrib["Target"] for x in relroot}
        sheet = wb.find("m:sheets/m:sheet", NS)
        target = rels[sheet.attrib["{" + NS["r"] + "}id"]].lstrip("/")
        target = target if target.startswith("xl/") else "xl/" + target
        root = ET.fromstring(z.read(target))
        rows = []
        for row in root.findall(".//m:sheetData/m:row", NS):
            values = {}
            for c in row.findall("m:c", NS):
                v = c.find("m:v", NS)
                if c.attrib.get("t") == "inlineStr":
                    value = "".join(t.text or "" for t in c.findall(".//m:t", NS))
                elif v is None:
                    value = ""
                elif c.attrib.get("t") == "s":
                    value = ss[int(v.text)]
                else:
                    value = v.text
                values[col_index(c.attrib["r"])] = value
            rows.append([values.get(i, "") for i in range(max(values, default=-1)+1)])
    if not rows:
        raise ValueError("Geen gegevens gevonden.")
    headers = [str(x).strip().lower() for x in rows[0]]
    wi = next((i for i,h in enumerate(headers) if "breedte" in h or "width" in h), None)
    li = next((i for i,h in enumerate(headers) if "lengte" in h or "length" in h), None)
    ri = next((i for i,h in enumerate(headers) if "draai" in h or "rotat" in h), None)
    pi = next((i for i,h in enumerate(headers)
               if "pallet" in h and ("nummer" in h or "number" in h or "nr" in h)), None)
    if wi is None or li is None or pi is None:
        raise ValueError("Vereiste kolommen: pallet nummer, breedte en lengte.")
    crates, missing, seen = [], [], set()
    for lineno, row in enumerate(rows[1:], 2):
        number = str(row[pi]).strip() if pi < len(row) else ""
        w = row[wi] if wi < len(row) else ""
        l = row[li] if li < len(row) else ""
        if not number and not str(w).strip() and not str(l).strip():
            continue
        if not number:
            missing.append(lineno)
            continue
        if number in seen:
            raise ValueError(f"Dubbel palletnummer {number} op Excel-rij {lineno}.")
        try:
            w, l = float(str(w).replace(",", ".")), float(str(l).replace(",", "."))
            if w <= 0 or l <= 0:
                raise ValueError
        except (ValueError, TypeError):
            missing.append(lineno)
            continue
        rotation_value = str(row[ri]).strip().lower() if ri is not None and ri < len(row) else ""
        can_rotate = rotation_value in {"ja", "yes", "true", "1", "y"}
        seen.add(number)
        crates.append((number, w, l, can_rotate))
    return crates, missing

def contact_length(rect, placed, length, width, epsilon=0.01):
    x, y, rw, rh = rect
    contact = 0.0
    if abs(x) <= epsilon:
        contact += rh
    if abs(x + rw - length) <= epsilon:
        contact += rh
    if abs(y) <= epsilon:
        contact += rw
    if abs(y + rh - width) <= epsilon:
        contact += rw
    for _ident, px, py, pw, ph, _turned in placed:
        if abs(x + rw - px) <= epsilon or abs(px + pw - x) <= epsilon:
            contact += max(0.0, min(y + rh, py + ph) - max(y, py))
        if abs(y + rh - py) <= epsilon or abs(py + ph - y) <= epsilon:
            contact += max(0.0, min(x + rw, px + pw) - max(x, px))
    return contact


def lateral_position_score(rect, placed, width, max_gap=35):
    x, y, crate_length, crate_width = rect
    middle_x = x + crate_length / 2
    intervals = [(py, py + ph) for _ident, px, py, pw, ph, _turned in placed
                 if px < middle_x < px + pw]
    intervals.append((y, y + crate_width))
    intervals.sort()
    gaps = []
    cursor = 0.0
    for start, end in intervals:
        gaps.append(max(0.0, start - cursor))
        cursor = max(cursor, end)
    gaps.append(max(0.0, width - cursor))
    if len(intervals) == 1:
        # During packing, keep a lone first pallet at a wall so a second pallet
        # can join it. If it remains alone after the full run, center it below.
        return (0.0, min(gaps[0], gaps[-1]), 0.0)
    return (sum(max(0.0, gap - max_gap) for gap in gaps),
            gaps[0] + gaps[-1], 0.0)


def pack_once(ordered_crates, length, width, heuristic, lateral_bias=False):
    free = [(0.0, 0.0, length, width)]
    placed, failed = [], []
    for ident, crate_width, crate_length, can_rotate in ordered_crates:
        choices = []
        orientations = ([(crate_length, crate_width, False), (crate_width, crate_length, True)]
                        if can_rotate and crate_width != crate_length
                        else [(crate_length, crate_width, False)])
        for fi, (x, y, fw, fh) in enumerate(free):
            for rw, rh, turned in orientations:
                if rw > fw or rh > fh:
                    continue
                y_positions = {y, y + fh - rh}
                center_y = (width - rh) / 2
                right_wall_y = width - rh
                for anchor_y in (center_y, right_wall_y):
                    if y - 0.01 <= anchor_y and anchor_y + rh <= y + fh + 0.01:
                        y_positions.add(anchor_y)
                centered_single = (
                    lateral_bias and abs(y) <= 0.01 and abs(fh - width) <= 0.01
                    and 35 < width - rh <= 70
                    and y - 0.01 <= center_y and center_y + rh <= y + fh + 0.01
                )
                if centered_single:
                    y_positions = {center_y}
                for place_y in y_positions:
                    dx, dy = fw - rw, fh - rh
                    short_side, long_side = min(dx, dy), max(dx, dy)
                    area_left = fw * fh - rw * rh
                    contact = contact_length((x, place_y, rw, rh), placed, length, width)
                    if lateral_bias:
                        lateral = lateral_position_score((x, place_y, rw, rh), placed, width)
                        if heuristic == "BSSF":
                            fit_score = (short_side, long_side)
                        elif heuristic == "BLSF":
                            fit_score = (long_side, short_side)
                        elif heuristic == "BAF":
                            fit_score = (area_left, short_side, long_side)
                        elif heuristic == "Bottom Left":
                            fit_score = (place_y + rh, short_side, long_side)
                        else:  # Contact Point
                            fit_score = (short_side, long_side, -contact)
                        score = (x + rw, *lateral, *fit_score, place_y, x)
                    elif heuristic == "BSSF":
                        score = (short_side, long_side, place_y, x)
                    elif heuristic == "BLSF":
                        score = (long_side, short_side, place_y, x)
                    elif heuristic == "BAF":
                        score = (area_left, short_side, long_side, place_y, x)
                    elif heuristic == "Bottom Left":
                        score = (place_y + rh, x, short_side, long_side)
                    else:  # Contact Point
                        score = (-contact, short_side, long_side, place_y, x)
                    score += (int(turned), fi)
                    choices.append((score, x, place_y, rw, rh, turned))
        if not choices:
            failed.append(ident)
            continue
        _, x, y, rw, rh, turned = min(choices)
        placed.append((ident, x, y, rw, rh, turned))
        new = []
        for fx, fy, fw, fh in free:
            if x >= fx + fw or x + rw <= fx or y >= fy + fh or y + rh <= fy:
                new.append((fx, fy, fw, fh))
                continue
            if x > fx:
                new.append((fx, fy, x - fx, fh))
            if x + rw < fx + fw:
                new.append((x + rw, fy, fx + fw - x - rw, fh))
            if y > fy:
                new.append((fx, fy, fw, y - fy))
            if y + rh < fy + fh:
                new.append((fx, y + rh, fw, fy + fh - y - rh))
        free = [
            rect for i, rect in enumerate(new)
            if rect[2] > 0 and rect[3] > 0 and not any(
                i != j and rect[0] >= other[0] and rect[1] >= other[1]
                and rect[0] + rect[2] <= other[0] + other[2]
                and rect[1] + rect[3] <= other[1] + other[3]
                for j, other in enumerate(new)
            )
        ]
    # A pallet occupying an otherwise empty longitudinal slice is centered
    # across the container width after the layout is complete. This translation
    # does not change its fit or interfere with any other pallet.
    centered = []
    for ident, x, y, rw, rh, turned in placed:
        has_side_by_side_neighbor = any(
            other_id != ident and min(x + rw, ox + ow) - max(x, ox) > 0.01
            for other_id, ox, _oy, ow, _oh, _ot in placed
        )
        if not has_side_by_side_neighbor:
            y = (width - rh) / 2
        centered.append((ident, x, y, rw, rh, turned))
    return centered, failed


def evaluate_plan(placed, failed, length, width):
    loaded_length = max((x + rw for _id, x, _y, rw, _rh, _turned in placed), default=0.0)
    contacts = 0.0
    for i, (_id, x, y, rw, rh, _turned) in enumerate(placed):
        contacts += contact_length((x, y, rw, rh), placed[:i], length, width)
    return (len(placed), -loaded_length, contacts, -len(failed)), loaded_length, contacts


def evaluate_lateral_gaps(placed, length, width, max_gap=35):
    if not placed:
        return 0.0, 0.0, 0.0
    bounds = sorted({value for _ident, x, _y, crate_length, _w, _turned in placed
                     for value in (x, x + crate_length)})
    unsafe_area = wall_gap_cost = single_center_cost = 0.0
    for left, right in zip(bounds, bounds[1:]):
        span = right - left
        if span <= 0.01:
            continue
        middle_x = (left + right) / 2
        intervals = sorted((y, y + crate_width)
                           for _ident, x, y, crate_length, crate_width, _turned in placed
                           if x < middle_x < x + crate_length)
        if not intervals:
            unsafe_area += width * span
            continue
        gaps, cursor = [], 0.0
        for start, end in intervals:
            gaps.append(max(0.0, start - cursor))
            cursor = max(cursor, end)
        gaps.append(max(0.0, width - cursor))
        unsafe_area += sum(max(0.0, gap - max_gap) for gap in gaps) * span
        if len(intervals) == 1:
            single_center_cost += abs(gaps[0] - gaps[-1]) * span
        else:
            wall_gap_cost += (gaps[0] + gaps[-1]) * span
    return unsafe_area, wall_gap_cost, single_center_cost


def pack_candidates(crates, length, width):
    orderings = [
        ("area-aflopend", sorted(crates, key=lambda c: c[1] * c[2], reverse=True)),
        ("langste-zijde", sorted(crates, key=lambda c: max(c[1], c[2]), reverse=True)),
        ("lengte-aflopend", sorted(crates, key=lambda c: c[2], reverse=True)),
        ("breedte-aflopend", sorted(crates, key=lambda c: c[1], reverse=True)),
        ("Excel-volgorde", list(crates)),
    ]
    heuristics = ["BSSF", "BLSF", "BAF", "Bottom Left", "Contact Point"]
    candidates = []
    for heuristic in heuristics:
        best = None
        for order_name, ordered in orderings:
            for lateral_bias in (False, True):
                placed, failed = pack_once(ordered, length, width, heuristic, lateral_bias)
                _base_score, used_length, contacts = evaluate_plan(placed, failed, length, width)
                unsafe_area, wall_gap_cost, single_center_cost = evaluate_lateral_gaps(
                    placed, length, width
                )
                score = (len(placed), -unsafe_area, -wall_gap_cost,
                         -single_center_cost, -used_length, contacts, -len(failed))
                plan = {
                    "placed": placed, "failed": failed, "heuristic": heuristic,
                    "order": order_name, "score": score,
                    "used_length": used_length, "contacts": contacts,
                    "unsafe_area": unsafe_area, "wall_gap_cost": wall_gap_cost,
                    "single_center_cost": single_center_cost,
                    "attempts": len(orderings) * 2,
                }
                if best is None or plan["score"] > best["score"]:
                    best = plan
        candidates.append(best)
    candidates.sort(key=lambda plan: plan["score"], reverse=True)
    return candidates

def hatch_segments(x0, y0, x1, y1, spacing=8):
    width, height = x1-x0, y1-y0
    for k in range(-int(height), int(width)+1, spacing):
        points = []
        for px, py in ((0, -k), (width, width-k), (k, 0), (height+k, height)):
            if -0.01 <= px <= width+0.01 and -0.01 <= py <= height+0.01:
                point = (x0+px, y0+py)
                if point not in points:
                    points.append(point)
        if len(points) >= 2:
            yield points[0], points[1]

def pallet_color_index(number, count=5):
    value = str(number).strip()
    try:
        return (int(float(value))-1) % count
    except (ValueError, OverflowError):
        return sum(ord(char) for char in value) % count

def _pdf_text(text):
    return str(text).encode("ascii", "replace").decode("ascii").replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")

def export_plan_pdf(path, placed, length, width, failed=(), missing=(), transport=DEFAULT_TRANSPORT, method=''):
    page_w, page_h = 595, 842
    scale = min(170/width, 640/length)
    box_w, box_h = width*scale, length*scale
    x0, y0 = 38, 105
    colors = [(0.66,0.84,0.73),(0.72,0.84,0.95),(0.96,0.84,0.55),(0.85,0.76,0.94),(0.95,0.72,0.72)]
    cmds = ["0.12 0.18 0.25 rg", f"BT /F1 18 Tf 38 812 Td ({_pdf_text(transport)} - vloerplan) Tj ET",
            f"BT /F1 10 Tf 38 793 Td (Binnenmaat: {length:g} x {width:g} cm) Tj ET",
            f"BT /F1 8 Tf 38 777 Td ({_pdf_text(method)}) Tj ET",
            "BT /F1 9 Tf 38 761 Td (KOPSCHOT - eerste pallets laden) Tj ET",
            "0.12 0.18 0.25 RG 1.3 w", f"{x0} {y0} {box_w:.2f} {box_h:.2f} re S"]
    for ident,x,y,w,h,turned in placed:
        rx, ry, rw, rh = x0+y*scale, y0+box_h-(x+w)*scale, h*scale, w*scale
        cr,cg,cb = colors[pallet_color_index(ident, len(colors))]
        cmds += [f"{cr:.3f} {cg:.3f} {cb:.3f} rg", f"{rx:.2f} {ry:.2f} {rw:.2f} {rh:.2f} re f"]
        if turned:
            cmds.append("0.36 0.42 0.47 RG 0.45 w")
            for (ax,ay),(bx,by) in hatch_segments(rx,ry,rx+rw,ry+rh,6):
                cmds.append(f"{ax:.2f} {ay:.2f} m {bx:.2f} {by:.2f} l S")
        cmds += ["0.12 0.18 0.25 RG 0.65 w", f"{rx:.2f} {ry:.2f} {rw:.2f} {rh:.2f} re S"]
        fs = max(6, min(9, rw/7, rh/4))
        cmds.append(f"0.08 0.12 0.18 rg BT /F1 {fs:.1f} Tf {rx+2:.2f} {ry+rh/2:.2f} Td ({_pdf_text('Pallet '+str(ident))}) Tj ET")
    cmds.append(f"BT /F1 9 Tf {x0} {y0-17} Td (DEUREN) Tj ET")
    cmds.append("BT /F1 10 Tf 225 761 Td (Palletposities - x vanaf kopschot, y vanaf linkerzijde (cm)) Tj ET")
    lines = [f"Pallet {ident}: x={x:.1f}, y={y:.1f}; lengte={w:g}, breedte={h:g}" + ("; GEDRAAID" if turned else "") for ident,x,y,w,h,turned in placed]
    lines += [f"Pallet {ident}: NIET GEPLAATST" for ident in failed]
    lines += [f"Excel-rij {row}: ontbrekende of ongeldige maten" for row in missing]
    for i, line in enumerate(lines):
        col = i // 44
        row = i % 44
        tx = 225 + col*185
        ty = 742 - row*14
        cmds.append(f"BT /F1 8 Tf {tx} {ty} Td ({_pdf_text(line)}) Tj ET")
    content = "\n".join(cmds).encode("ascii", "replace")
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {page_w} {page_h}] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>".encode(),
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
        b"<< /Length " + str(len(content)).encode() + b" >>\nstream\n" + content + b"\nendstream",
    ]
    pdf = bytearray(b"%PDF-1.4\n")
    offsets = [0]
    for i,obj in enumerate(objects,1):
        offsets.append(len(pdf)); pdf += f"{i} 0 obj\n".encode() + obj + b"\nendobj\n"
    xref = len(pdf)
    pdf += f"xref\n0 {len(offsets)}\n0000000000 65535 f \n".encode()
    for off in offsets[1:]: pdf += f"{off:010d} 00000 n \n".encode()
    pdf += f"trailer\n<< /Size {len(offsets)} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF".encode()
    Path(path).write_bytes(pdf)

class App:
    def __init__(self, root):
        self.root = root
        root.title("Container vloerplan V4.0")
        root.geometry("1450x760")
        bar = ttk.Frame(root, padding=10); bar.pack(fill="x")
        ttk.Label(bar, text="XLSX-bestand").pack(side="left")
        self.file = tk.StringVar()
        self.combo = ttk.Combobox(bar, textvariable=self.file, width=27, state="readonly")
        self.combo.pack(side="left", padx=5)
        ttk.Label(bar, text="Vervoersmiddel").pack(side="left", padx=(10,3))
        self.transport = tk.StringVar(value=DEFAULT_TRANSPORT)
        self.transport_combo = ttk.Combobox(
            bar, textvariable=self.transport, values=list(TRANSPORTS), width=18, state="readonly"
        )
        self.transport_combo.pack(side="left", padx=4)
        self.transport_combo.bind("<<ComboboxSelected>>", self.update_transport)
        ttk.Label(bar, text="Lengte (cm)").pack(side="left", padx=(7,2))
        self.length = tk.StringVar()
        ttk.Entry(bar, textvariable=self.length, width=8, state="readonly").pack(side="left")
        ttk.Label(bar, text="Breedte (cm)").pack(side="left", padx=(7,2))
        self.width = tk.StringVar()
        ttk.Entry(bar, textvariable=self.width, width=7, state="readonly").pack(side="left")
        ttk.Button(bar, text="Bereken vloerplan", command=self.calculate).pack(side="left", padx=8)
        body = ttk.Panedwindow(root, orient="horizontal"); body.pack(fill="both", expand=True, padx=10, pady=4)
        diagram = ttk.Frame(body); body.add(diagram, weight=4)
        self.canvas = tk.Canvas(diagram, bg="#f2f4f7", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind("<Configure>", self._redraw_plan)
        sidebar = ttk.Frame(body, width=300); body.add(sidebar, weight=1)
        ttk.Label(sidebar, text="Beste gevonden indelingen").pack(anchor="w")
        self.plan_choice = ttk.Combobox(sidebar, state="readonly", width=42)
        self.plan_choice.pack(fill="x", pady=(4, 2))
        self.plan_choice.bind("<<ComboboxSelected>>", self.choose_plan)
        self.plan_summary = tk.StringVar(value="Bereken eerst een vloerplan.")
        ttk.Label(sidebar, textvariable=self.plan_summary, wraplength=285).pack(anchor="w", pady=(2, 6))
        ttk.Button(sidebar, text="Sla gekozen plan op en open PDF",
                   command=self.create_and_open_pdf).pack(fill="x", pady=(0, 8))
        ttk.Label(sidebar, text="Palletposities vanaf kopschot (cm)", wraplength=270).pack(anchor="w")
        self.list = tk.Listbox(sidebar, width=36); self.list.pack(fill="both", expand=True, pady=6)
        self.status = tk.StringVar(value="")
        ttk.Label(root, textvariable=self.status, anchor="w", padding=8).pack(fill="x")
        self.current_plan = None
        self.plan_options = []
        self.total_crates = 0
        self.missing_rows = []
        self.update_transport()
        self.refresh_files()

    def update_transport(self, _event=None):
        length_m, width_m = TRANSPORTS[self.transport.get()]
        self.length.set(f"{length_m * 100:g}")
        self.width.set(f"{width_m * 100:g}")

    def refresh_files(self):
        files = []
        for p in HERE.glob("*.xlsx"):
            try:
                read_xlsx(p)
                files.append(p.name)
            except ValueError as error:
                if str(error).startswith("Dubbel palletnummer "):
                    files.append(p.name)
            except Exception:
                pass
        self.combo["values"] = files
        if files:
            self.combo.current(0)
        else:
            self.status.set("Geen echte .xlsx-bestanden gevonden in de scriptmap.")

    def calculate(self):
        if not self.file.get():
            return
        try:
            length, width = float(self.length.get()), float(self.width.get())
            if length <= 0 or width <= 0:
                raise ValueError
            crates, missing = read_xlsx(HERE / self.file.get())
            candidates = pack_candidates(crates, length, width)
            if not candidates:
                raise ValueError("Er kon geen plaatsingsplan worden gemaakt.")
        except Exception as error:
            messagebox.showerror("Kan vloerplan niet berekenen", str(error))
            return
        self.plan_options = candidates
        self.total_crates = len(crates)
        self.missing_rows = missing
        self.plan_choice["values"] = [
            self.plan_label(index, plan) for index, plan in enumerate(candidates)
        ]
        self.plan_choice.current(0)
        self.show_plan(0)

    def plan_label(self, index, plan):
        used_m = plan["used_length"] / 100
        rank = "Aanbevolen" if index == 0 else f"Alternatief {index}"
        return (f"{rank}: {len(plan['placed'])}/{self.total_crates} pallets · "
                f"{used_m:.2f} m gebruikt · {plan['heuristic']} / {plan['order']}")

    def choose_plan(self, _event=None):
        index = self.plan_choice.current()
        if 0 <= index < len(self.plan_options):
            self.show_plan(index)

    def show_plan(self, index):
        plan = self.plan_options[index]
        placed, failed = plan["placed"], plan["failed"]
        length, width = float(self.length.get()), float(self.width.get())
        transport = self.transport.get()
        method = f"{plan['heuristic']} / {plan['order']} (plan {index + 1} van {len(self.plan_options)})"
        self.current_plan = (placed, length, width, failed, self.missing_rows, transport, method)
        self.list.delete(0, tk.END)
        for ident, x, y, w, h, turned in placed:
            turn_note = " · GEDRAAID" if turned else ""
            self.list.insert(
                tk.END,
                f"Pallet {ident}: x={x:.1f} vanaf kopschot, y={y:.1f} vanaf links · "
                f"L {w:g} × B {h:g} cm{turn_note}"
            )
        for ident in failed:
            self.list.insert(tk.END, f"Pallet {ident}: past niet")
        score_text = (
            f"{len(placed)}/{self.total_crates} pallets geplaatst. "
            f"Gebruikte lengte: {plan['used_length']/100:.2f} m. "
            f"Contact met pallet/wand: {plan['contacts']:.0f} cm. "
            f"Onondersteunde zijruimte boven 35 cm (lengte-gewogen): {plan['unsafe_area']:.0f} cm². "
            f"Rangschikking: meeste pallets, zijruimtes, enkele pallets centreren, buitenste pallets tegen wanden, compactheid. "
            f"Beste van {plan['attempts']} volgorde-/breedtevarianten voor {plan['heuristic']}."
        )
        if self.missing_rows:
            score_text += f" Onvolledige Excel-rijen: {self.missing_rows}."
        self.plan_summary.set(score_text)
        self.status.set(
            f"Plan {index + 1}/{len(self.plan_options)}: {len(placed)}/{self.total_crates} "
            f"pallets geplaatst in {transport}. Niet geplaatst: {failed or 'geen'}."
        )
        self.draw(placed, length, width)

    def create_and_open_pdf(self):
        if not self.current_plan:
            return
        transport = self.current_plan[5]
        method = self.current_plan[6]
        safe_transport = re.sub(r"[^A-Za-z0-9_-]+", "_", transport).strip("_")
        safe_method = re.sub(r"[^A-Za-z0-9_-]+", "_", method).strip("_")
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = HERE / f"vloerplan_{safe_transport}_{safe_method}_{stamp}.pdf"
        try:
            export_plan_pdf(path, *self.current_plan)
        except Exception as error:
            messagebox.showerror("PDF maken mislukt", str(error))
            return
        try:
            if sys.platform == "win32":
                os.startfile(str(path))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
            self.status.set(f"PDF opgeslagen en geopend: {path.name}")
        except Exception as error:
            messagebox.showwarning(
                "PDF opgeslagen",
                f"De PDF staat hier, maar kon niet automatisch worden geopend:\n{path}\n\n{error}",
            )

    def _redraw_plan(self, _event=None):
        if self.current_plan:
            self.draw(*self.current_plan[:3])

    def draw(self, placed, length, width):
        c = self.canvas
        c.delete("all")
        c.update_idletasks()
        cw, ch = max(c.winfo_width()-60, 300), max(c.winfo_height()-130, 180)
        scale = min(cw/width, ch/length)
        box_w, box_h = width*scale, length*scale
        ox = (c.winfo_width()-box_w)/2
        oy = 55
        c.create_text(ox+box_w/2, oy-34, text=f"Container · {length:g} × {width:g} cm",
                      font=("Sans", 11, "bold"))
        c.create_text(ox+box_w/2, oy-16, text="KOPSCHOT - eerste pallets laden",
                      font=("Sans", 9, "bold"), fill="#a33a2b")
        c.create_rectangle(ox, oy, ox+box_w, oy+box_h, fill="#ffffff", outline="#243247", width=2)
        colors = ["#a8d5ba","#b9d6f2","#f5d58b","#d8c2f0","#f2b6b6"]
        for ident, x, y, w, h, turned in placed:
            x0, y0 = ox+y*scale, oy+x*scale
            x1, y1 = ox+(y+h)*scale, oy+(x+w)*scale
            c.create_rectangle(x0, y0, x1, y1, fill=colors[pallet_color_index(ident, len(colors))],
                               outline="#263445", width=1)
            if turned:
                for (ax,ay),(bx,by) in hatch_segments(x0+1,y0+1,x1-1,y1-1,8):
                    c.create_line(ax, ay, bx, by, fill="#71808A", width=1)
            font_size = max(7, min(11, int(min(x1-x0, y1-y0)/4)))
            c.create_text((x0+x1)/2, (y0+y1)/2, text=f"Pallet {ident}",
                          width=max(12, x1-x0-3), font=("Sans", font_size, "bold"))
        c.create_text(ox+box_w/2, oy+box_h+18, text="DEUREN", font=("Sans", 9, "bold"))

if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
