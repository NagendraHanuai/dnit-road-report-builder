#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ROAD ATHENA  -  DNIT 006/2003-PRO REPORT BUILDER   (PyQt5 & CLI)
=================================================================
Input : the MAIN pavement-report .xlsx   (sheets "Pavement Anomaly" + "Chainage wise report")
        OR a folder containing multiple pavement report .xlsx files.
Output: New .xlsx report(s)  ->  Summary | Search | Section sheets (100 m, 200 m, 300 m, etc.)

Features:
  * User-configurable section length (e.g. 100 m, 200 m, 300 m).
  * Station evaluation per 20 m interval (n = section_len / 20).
  * Re-added & enhanced Chainage Diagram:
      - Plotted along exact chainage distance axis with ticks every 5 m.
      - Vertical station boundary lines every 20 m.
      - Top/middle: individual anomalies with simple rectangular vertical markers, individual
        defect colours, and vertically staggered ID labels for close chainages.
      - Bottom station summary bar: colored and labeled with the WORST defect category
        selected using priority ranking (fp 1.0 > 0.9 > 0.8 > 0.6 > 0.5 > 0.3 > 0.2).
      - Note lines explicitly explaining Top/middle (individual) vs Bottom (station summary).
  * Support for single file OR entire folder batch processing.

Run:   python road_report_builder_copy.py                       (GUI)
       python road_report_builder_copy.py in.xlsx out.xlsx      (CLI single file)
       python road_report_builder_copy.py in_folder/ out_folder/ (CLI folder batch)
"""
from __future__ import annotations

import io
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime

import matplotlib
matplotlib.use("Agg")
from matplotlib.backends.backend_agg import FigureCanvasAgg
from matplotlib.figure import Figure
from matplotlib.patches import Rectangle

import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.properties import PageSetupProperties
from PIL import Image as PILImage

# --------------------------------------------------------------------------------------
# 1.  DNIT TABLES & CONSTANTS
# --------------------------------------------------------------------------------------
DEFAULT_SECTION_LEN = 100.0   # m per sheet (configurable: 100, 200, 300, etc.)
STATION_LEN = 20.0            # m per DNIT station
LANE_WIDTH = 3.7              # m (default, editable)

TYPES = {  # type: (codes, description, fp)
    1: ("TLC, TLL, TTC, TTL, other isolated cracks", "Isolated cracks (Type 1) - e.g. TTL long transverse, TTC transverse", 0.2),
    2: ("J, TB", "FC-2 - alligator / block cracks WITHOUT edge erosion", 0.5),
    3: ("JE, TBE", "FC-3 - alligator / block cracks WITH edge erosion", 0.8),
    4: ("ALP, ATP, ALC, ATC", "Plastic deformation (rutting) / corrugation", 0.9),
    5: ("O, P, E", "Pothole, deteriorated patch, edge break", 1.0),
    6: ("EX", "Existing patch (any condition)", 0.5),
    7: ("D", "Surface disintegration / delamination", 0.3),
    8: ("R", "Raveling (loss of aggregate)", 0.6),
}
FP = {t: v[2] for t, v in TYPES.items()}
FP[9] = 0.0

KNOWN = {
    "ALLIGATOR CRACK": ("J", 2, "J (FC-2): alligator cracking, no marked edge erosion visible. FC-3 (JE) would need erosion."),
    "ALLIGATOR CRACKING": ("J", 2, "J (FC-2): alligator cracking, no marked edge erosion visible. FC-3 (JE) would need erosion."),
    "CRACK L H": ("TTL", 1, "Isolated long/transverse crack (TTL). FC-1/2/3 class needs opening + edge erosion check."),
    "CRACK L/H": ("TTL", 1, "Isolated long/transverse crack (TTL). FC-1/2/3 class needs opening + edge erosion check."),
    "LONGITUDINAL CRACK": ("TTL", 1, "Isolated long/transverse crack (TTL)."),
    "TRANSVERSE CRACK": ("TTC", 1, "Isolated long/transverse crack (TTC)."),
    "SURFACE DELAMINATION": ("D", 7, "D: surface disintegration / delamination of the wearing layer."),
    "DELAMINATION": ("D", 7, "D: surface disintegration / delamination of the wearing layer."),
    "PATCH": ("EX", 6, "EX: existing patch, any condition."),
    "PATCHING": ("EX", 6, "EX: existing patch, any condition."),
    "POTHOLE": ("O", 5, "O: pothole (panela). Any pothole is Type 5, fp 1.0."),
    "POT HOLE": ("O", 5, "O: pothole (panela). Any pothole is Type 5, fp 1.0."),
    "RAVELLING": ("R", 8, "R: raveling - loss of aggregate / surface wear."),
    "RAVELING": ("R", 8, "R: raveling - loss of aggregate / surface wear."),
    "BLEEDING": ("-", 9, "Bleeding is NOT one of the 8 DNIT types: recorded only, fp 0 (not in IGG)."),
}

MARK = {1: "#7ED37E", 2: "#6FA8DC", 3: "#E57373", 4: "#FFD54F", 5: "#B48CE0", 6: "#4DD0E1", 7: "#F6A25B", 8: "#E07BC4", 9: "#BDBDBD"}
LIGHT = {1: "C6EFCE", 2: "BDD7EE", 3: "F8CBD6", 4: "FFE699", 5: "D9C7F0", 6: "CCF2F4", 7: "FAD7B5", 8: "F4B6D9"}
COND = [("Very good", "0 < IGG <= 20", 20, "92D050"), ("Good", "20 < IGG <= 40", 40, "C6EFCE"),
        ("Regular", "40 < IGG <= 80", 80, "FFE066"), ("Bad", "80 < IGG <= 160", 160, "F4B183"),
        ("Lousy", "IGG > 160", None, "FF7C80")]
NAVY, BLUE, PALE, GREY = "1F2A44", "2F5597", "DCE6F1", "F2F2F2"


def norm(name: str) -> str:
    """Only used to LOOK UP a name in the mapping - the name itself is never changed."""
    return re.sub(r"[\s/_\-]+", " ", str(name).strip().upper())


def classify(raw: str):
    return KNOWN.get(norm(raw))      # None for unknown names


def km(m: float) -> str:
    m = int(round(m))
    return f"{m // 1000}+{m % 1000:03d}"


def condition_of(igg: float):
    for name, _rng, up, col in COND:
        if up is None or igg <= up:
            return name, col
    return COND[-1][0], COND[-1][3]


# --------------------------------------------------------------------------------------
# 2.  READ THE MAIN REPORT
# --------------------------------------------------------------------------------------
class ReportInfo:
    pass


def _num(v):
    try:
        return float(str(v).replace(",", "").strip())
    except Exception:
        return None


def _url(cell):
    v = cell.value
    if isinstance(v, str):
        m = re.search(r'HYPERLINK\(\s*"([^"]+)"', v, re.I)
        if m:
            return m.group(1)
        if cell.hyperlink is not None and cell.hyperlink.target:
            return cell.hyperlink.target
        return v.strip()
    return "" if v is None else str(v)


def parse_report(path: str, side_override: str = "Auto", section_len: float = DEFAULT_SECTION_LEN, progress=None) -> ReportInfo:
    P = progress or (lambda p, t: None)
    P(1, f"Step 1/5  Opening main report: {os.path.basename(path)} ...")
    wb = openpyxl.load_workbook(path)
    ws_a = next((wb[n] for n in wb.sheetnames if "anomal" in n.lower()), wb.worksheets[0])
    ws_c = next((wb[n] for n in wb.sheetnames if "chainage" in n.lower()), None)

    info = ReportInfo()
    info.path, info.file = path, os.path.basename(path)
    info.section_len = section_len
    info.warnings = []

    # ---- header block ------------------------------------------------------------
    info.title = str(ws_a["A1"].value or "AI Based road condition assessment detailed report")
    txt2 = str(ws_a["A2"].value or "")
    m = re.search(r"(\d{2}[-/]\d{2}[-/]\d{4})", txt2)
    info.date = m.group(1) if m else txt2.replace("Date of survey:", "").strip()
    seg = str(ws_a["A3"].value or "")
    info.segment = re.sub(r"^\s*Segment\s*Name\s*:\s*", "", seg, flags=re.I).strip()
    start = end = None
    hdr_row = None
    for r in range(1, 15):
        a = str(ws_a.cell(r, 1).value or "").strip().lower()
        vals = [_num(ws_a.cell(r, c).value) for c in range(2, 9)]
        vals = [v for v in vals if v is not None]
        if a.startswith("start chainage") and vals:
            start = vals[0]
        elif a.startswith("end chainage") and vals:
            end = vals[0]
        elif a.startswith("anomaly number"):
            hdr_row = r
    if hdr_row is None:
        raise ValueError("Could not find the table header 'Anomaly Number' in the main sheet.")
    cols = {}
    for c in range(1, ws_a.max_column + 1):
        h = str(ws_a.cell(hdr_row, c).value or "").strip().lower()
        if h.startswith("anomaly number"): cols["no"] = c
        elif h.startswith("timestamp"): cols["ts"] = c
        elif h.startswith("anomaly type"): cols["type"] = c
        elif h.startswith("latitude"): cols["lat"] = c
        elif h.startswith("longitude"): cols["lon"] = c
        elif h.startswith("distance"): cols["dist"] = c
        elif h.startswith("image"): cols["url"] = c

    # ---- anomalies ---------------------------------------------------------------
    P(3, "Step 1/5  Reading anomaly rows ...")
    an = []
    for r in range(hdr_row + 1, ws_a.max_row + 1):
        d = _num(ws_a.cell(r, cols["dist"]).value)
        if d is None:
            continue
        raw = ws_a.cell(r, cols["type"]).value
        an.append(dict(
            no=ws_a.cell(r, cols["no"]).value, ts=str(ws_a.cell(r, cols["ts"]).value or ""),
            raw=("" if raw is None else str(raw).strip()),
            lat=_num(ws_a.cell(r, cols["lat"]).value), lon=_num(ws_a.cell(r, cols["lon"]).value),
            dist=d, url=_url(ws_a.cell(r, cols["url"]))))
    if not an:
        raise ValueError("No anomaly rows were found in the main sheet.")
    info.anoms = an
    maxd = max(a["dist"] for a in an)
    if start is None:
        start = 0.0
        info.warnings.append("Start chainage not found - 0 used.")
    if end is None:
        end = start + maxd
        info.warnings.append("End chainage not found - taken from the last anomaly distance.")
    info.start, info.end = start, end

    # ---- LHS / RHS ---------------------------------------------------------------
    P(4, "Step 2/5  Detecting lane and LHS / RHS ...")
    ms = re.search(r"\b(LHS|RHS)\b", info.segment, re.I)
    name_side = ms.group(1).upper() if ms else None
    trend_side = "LHS" if end < start else "RHS"
    if ws_c is not None:
        fr_to = []
        for r in range(1, ws_c.max_row + 1):
            a, b = _num(ws_c.cell(r, 1).value), _num(ws_c.cell(r, 2).value)
            if a is not None and b is not None and ws_c.cell(r, 3).value is not None:
                fr_to.append((a, b))
        if fr_to:
            trend_side = "LHS" if fr_to[0][1] < fr_to[0][0] else "RHS"
    info.side_name, info.side_trend = name_side, trend_side
    if side_override in ("LHS", "RHS"):
        info.side, info.side_src = side_override, "set manually"
    elif name_side:
        info.side, info.side_src = name_side, "segment name in main sheet"
    else:
        info.side, info.side_src = trend_side, "chainage trend (decreasing = LHS, increasing = RHS)"
    if name_side and name_side != trend_side:
        info.warnings.append(f"Segment name says {name_side} but the chainage runs like {trend_side}. "
                             f"Using {info.side}.")
    info.sign = -1 if info.side == "LHS" else 1
    ml = re.search(r"Lane\s*#?\s*(\d+)", info.segment, re.I)
    info.lane = ml.group(1) if ml else "1"
    md = re.search(r"\(\s*([A-Za-z]+\s+to\s+[A-Za-z]+)\s*\)", info.segment)
    info.direction = md.group(1) if md else ("East to West" if info.side == "LHS" else "West to East")
    info.total = abs(end - start)
    info.nsec = max(1, math.ceil(info.total / section_len - 1e-9))

    # ---- main-sheet section counts -------------------------------------------------
    info.main_counts = {}
    if ws_c is not None:
        cK = cN = cQ = cT = None
        for r in (6, 7, 8):
            for c in range(1, ws_c.max_column + 1):
                h = str(ws_c.cell(r, c).value or "").lower()
                if "alligator" in h: cK = c
                elif "longitudinal" in h: cN = c
                elif h == "count" and c > 14: cQ = c
                elif h == "rank": cT = c
        cK, cN, cQ, cT = cK or 11, cN or 14, cQ or 17, cT or 20
        for r in range(1, ws_c.max_row + 1):
            a, b = _num(ws_c.cell(r, 1).value), _num(ws_c.cell(r, 2).value)
            if a is None or b is None or ws_c.cell(r, 3).value is None:
                continue
            k = int(round(abs(a - start) / section_len))
            info.main_counts[k] = tuple(_num(ws_c.cell(r, c).value) for c in (cK, cN, cQ, cT))

    # ---- logos --------------------------------------------------------------------
    imgs = []
    for im in ws_a._images:
        try:
            imgs.append((im.anchor._from.col, im._data()))
        except Exception:
            pass
    imgs.sort(key=lambda x: x[0])
    info.logo_left = imgs[0][1] if imgs else None
    info.logo_right = imgs[-1][1] if len(imgs) > 1 else None

    # ---- names found -------------------------------------------------------------
    names = {}
    for a in an:
        names[a["raw"]] = names.get(a["raw"], 0) + 1
    info.names = sorted(names.items(), key=lambda kv: -kv[1])
    info.unknown = [n for n, _ in info.names if classify(n) is None]
    if info.unknown:
        info.warnings.append("Not in the DNIT mapping (left EMPTY, excluded from IGG): " + ", ".join(info.unknown))
    P(6, "Step 2/5  Report understood.")
    return info


# --------------------------------------------------------------------------------------
# 3.  SECTION GEOMETRY + DNIT ANALYSIS
# --------------------------------------------------------------------------------------
def sections(info, section_len: float = DEFAULT_SECTION_LEN):
    out = []
    nsec = max(1, math.ceil(info.total / section_len - 1e-9))
    for k in range(nsec):
        d0 = k * section_len
        d1 = min((k + 1) * section_len, info.total)
        L = d1 - d0
        n = max(1, math.ceil(L / STATION_LEN - 1e-9))
        c0, c1 = info.start + info.sign * d0, info.start + info.sign * d1
        out.append(dict(k=k, d0=d0, d1=d1, L=L, n=n, c0=c0, c1=c1, section_len=section_len,
                        name=f"Km {km(c0)}-{km(c1)}", anoms=[]))
    return out


def assign(info, secs, section_len: float = DEFAULT_SECTION_LEN):
    warn = 0
    for a in info.anoms:
        k = int(a["dist"] // section_len)
        if k < 0 or k >= len(secs):
            warn += 1
            k = min(max(k, 0), len(secs) - 1)
        s = secs[k]
        a["sec"] = k
        a["station"] = min(s["n"], max(1, int((a["dist"] - s["d0"]) // STATION_LEN) + 1))
        c = classify(a["raw"])
        a["code"], a["type_no"], a["basis"] = (c if c else ("", None, ""))
        a["fp"] = FP[a["type_no"]] if a["type_no"] else None
        s["anoms"].append(a)
    if warn:
        info.warnings.append(f"{warn} anomaly distance(s) lie outside survey length - placed in nearest section.")
    for s in secs:
        s["anoms"].sort(key=lambda x: (x["dist"], str(x["no"])))


def analyse(s):
    n = s["n"]
    st = {i: [] for i in range(1, n + 1)}
    for a in s["anoms"]:
        if a["type_no"] and a["fp"] and a["fp"] > 0:
            st[a["station"]].append(a)
    worst = {}
    for i, lst in st.items():
        if lst:
            # Highest weighting factor (fp)
            fp = max(x["fp"] for x in lst)
            # Tie-break: lowest type_no (e.g. Type 2 J before Type 6 EX if fp is equal 0.5)
            t = min(x["type_no"] for x in lst if x["fp"] == fp)
            codes = sorted({x["code"] for x in lst if x["type_no"] == t})
            worst[i] = (t, fp, "/".join(codes))
    FA = {t: 0 for t in range(1, 9)}
    for t, _, _ in worst.values():
        FA[t] += 1
    igg = sum(FA[t] / n * 100 * FP[t] for t in FA)
    s["worst"], s["FA"], s["igg"] = worst, FA, igg
    s["cond"], s["cond_col"] = condition_of(igg)
    s["affected"] = len(worst)


def locate(info, metres: float, section_len: float = DEFAULT_SECTION_LEN):
    """sheet index (0-based) for an absolute chainage in metres, or None if outside survey."""
    d = info.sign * (metres - info.start)
    if d < -1e-9 or d > info.total + 1e-9:
        return None
    nsec = max(1, math.ceil(info.total / section_len - 1e-9))
    return min(nsec - 1, int(d // section_len))


def parse_chainage(text: str):
    t = str(text).strip().replace(" ", "").replace(",", "")
    if not t:
        return None
    if "+" in t:
        a, b = t.split("+", 1)
        try:
            return float(a) * 1000 + float(b)
        except ValueError:
            return None
    try:
        return float(t)
    except ValueError:
        return None


# --------------------------------------------------------------------------------------
# 4.  CHAINAGE DIAGRAM (matplotlib) - Simple vertical markers + Bottom station summary
# --------------------------------------------------------------------------------------
def draw_diagram(info, s, lane_w=LANE_WIDTH, section_len=DEFAULT_SECTION_LEN) -> bytes:
    sec_l = s.get("section_len", section_len)
    fig_w = min(22.0, max(14.7, 14.7 * (sec_l / 100.0) ** 0.65))
    fig = Figure(figsize=(fig_w, 3.8), dpi=100)
    canvas = FigureCanvasAgg(fig)

    ax = fig.add_axes([0.085, 0.33, 0.83, 0.48])
    ax2 = fig.add_axes([0.085, 0.11, 0.83, 0.14])
    L, W = s["L"], lane_w
    ax.set_xlim(0, sec_l)
    ax.set_ylim(-W / 2 - 0.05, W / 2 + 1.7)
    ax2.set_xlim(0, sec_l)
    ax2.set_ylim(0, 1)
    for a in (ax, ax2):
        a.axis("off")

    ax.add_patch(Rectangle((0, -W / 2), L, W, fc="#F7F8FC", ec="none", zorder=0))

    # Station vertical boundary lines every 20 m
    ticks = [t for t in range(0, int(L) + 1, 20)]
    if L - ticks[-1] > 0.5:
        ticks.append(L)
    for t in ticks:
        x = min(t, L)
        ax.vlines(x, W / 2, W / 2 + 1.55, color="#8C8C8C", lw=1)
        ax.vlines(x, -W / 2, W / 2, color="#8C8C8C", lw=1)
        ax.text(x, W / 2 + 1.62, km(s["c0"] + info.sign * x), ha="center", va="bottom", fontsize=8.5, fontweight="bold",
                bbox=dict(boxstyle="round,pad=0.25", fc="white", ec="black", lw=1), clip_on=False)

    # Sub-ticks every 5 m
    for x in [i * 5 for i in range(1, int(L // 5) + 1)]:
        if x % 20:
            ax.vlines(x, -W / 2, W / 2, color="#D0D0D0", lw=0.5, ls=(0, (1, 3)))

    ax.hlines([W / 2, -W / 2], 0, L, color="black", lw=2.4)
    ax.hlines(0, 0, L, color="#555555", lw=1.1, ls=(0, (4, 3)))

    # Group anomalies into chainage clusters for vertical marker stacking within lane boundaries
    clusters = []
    anoms_sorted = sorted(s["anoms"], key=lambda a: a["dist"])
    threshold = 2.0 * (sec_l / 100.0)

    for a in anoms_sorted:
        x_val = min(max(a["dist"] - s["d0"], 0), L)
        a_item = dict(a, x=x_val)
        if not clusters:
            clusters.append([a_item])
        else:
            if abs(x_val - clusters[-1][0]["x"]) <= threshold:
                clusters[-1].append(a_item)
            else:
                clusters.append([a_item])

    last_single_x, single_level = -99, 0
    for cluster in clusters:
        N = len(cluster)
        if N == 1:
            a = cluster[0]
            x = a["x"]
            col = MARK.get(a["type_no"], "#9E9E9E")
            # Single anomaly centered in lane
            ax.add_patch(Rectangle((x - 0.2, -0.35), 0.4, 0.7, fc=col, ec="#333333", lw=0.6, zorder=3,
                                   hatch=None if a["type_no"] else "////"))
            single_level = single_level + 1 if abs(x - last_single_x) < (2.2 * sec_l / 100.0) else 0
            last_single_x = x
            ax.text(x, W / 2 + 0.28 + single_level * 0.46, str(a["no"]), ha="center", va="center", fontsize=7, fontweight="bold",
                    bbox=dict(boxstyle="round,pad=0.18", fc="white", ec="#555555", lw=0.6), zorder=4)
        else:
            # Multiple anomalies stacked vertically inside lane boundaries
            max_usable_span = W - 0.8
            spacing = min(0.75, max_usable_span / (N - 1))
            h_m = min(0.45, max(0.22, spacing * 0.75))
            w_m = 0.4
            mid = (N - 1) / 2.0
            for k, a in enumerate(cluster):
                x = a["x"]
                col = MARK.get(a["type_no"], "#9E9E9E")
                y_k = (k - mid) * spacing

                # Separate small vertical rectangular marker patch
                ax.add_patch(Rectangle((x - w_m / 2, y_k - h_m / 2), w_m, h_m, fc=col, ec="#333333", lw=0.6, zorder=3,
                                       hatch=None if a["type_no"] else "////"))

                # ID label centered directly inside/on each stacked marker
                fs = 6.5 if N > 5 else 7.0
                ax.text(x, y_k, str(a["no"]), ha="center", va="center", fontsize=fs, fontweight="bold",
                        bbox=dict(boxstyle="round,pad=0.15", fc="white", ec="#333333", lw=0.5), zorder=4)

    # Station Summary Bar (bottom) - exact consistent colors based on selected worst defect
    for i in range(1, s["n"] + 1):
        x0 = (i - 1) * STATION_LEN
        w = min(STATION_LEN, L - x0)
        w_ = s["worst"].get(i)
        fc = "#" + LIGHT[w_[0]] if (w_ and w_[0] in LIGHT) else "#F2F2F2"
        ax2.add_patch(Rectangle((x0, 0.04), w, 0.92, fc=fc, ec="#444444", lw=0.8))
        code_str = w_[2] if w_ else "-"
        ax2.text(x0 + w / 2, 0.5, f"S{i}\n{code_str}", ha="center", va="center", fontsize=9, fontweight="bold")

    fig.text(0.075, 0.57, f"Lane #{info.lane}\nWidth = {W:g} m", ha="right", va="center", fontsize=9)
    fig.text(0.075, 0.78, "left side \u25B2", ha="right", va="center", fontsize=8, color="#555555")
    fig.text(0.075, 0.37, "right side \u25BC", ha="right", va="center", fontsize=8, color="#555555")
    fig.text(0.075, 0.15, "DNIT stations\n(20 m each)", ha="right", va="center", fontsize=9)
    fig.text(0.925, 0.15, "Worst defect\nper station", ha="left", va="center", fontsize=9)
    fig.text(0.925, 0.57, "- - centre-line", ha="left", va="center", fontsize=8, color="#555555")

    note_line = (
        f"Chainage {km(s['c0'])} to {km(s['c1'])} ({info.direction}, {info.side})  |  "
        'Bottom = station summary ("Worst defect category for this 20 m section")  |  '
        "That separation is important. Do not colour every anomaly according to the station colour. "
        "Each anomaly should retain its own defect colour, while the bottom 20 m station uses the highest-priority defect found within that station."
    )
    fig.text(0.5, 0.010, note_line, ha="center", fontsize=7.2, fontweight="bold", color="#222222")

    buf = io.BytesIO()
    canvas.print_png(buf)
    return buf.getvalue()


# --------------------------------------------------------------------------------------
# 5.  EXCEL STYLE HELPERS
# --------------------------------------------------------------------------------------
def F(sz=10, b=False, col="000000", i=False, u=None):
    return Font(name="Arial", size=sz, bold=b, color=col, italic=i, underline=u)


def fill(c):
    return PatternFill("solid", fgColor=c)


THIN = Side(style="thin", color="A6A6A6")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CEN = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)


def box(ws, rng, value=None, font=None, bg=None, al=CEN, border=BOX, fmt=None):
    cells = ws[rng]
    if not isinstance(cells, tuple):
        cells = ((cells,),)
    elif cells and not isinstance(cells[0], tuple):
        cells = (cells,)
    for row in cells:
        for c in row:
            c.border = border
            if bg: c.fill = fill(bg)
            if font: c.font = font
            c.alignment = al
            if fmt: c.number_format = fmt
    if ":" in rng:
        ws.merge_cells(rng)
    first = ws[rng.split(":")[0]]
    if value is not None:
        first.value = value
    return first


def band(ws, row, text, last_col="T", color=BLUE):
    box(ws, f"A{row}:{last_col}{row}", text, F(11, True, "FFFFFF"), color, LEFT)
    ws.row_dimensions[row].height = 21.75


def add_logos(ws, info, right_col, h=62):
    for data, col, hh in ((info.logo_left, "A", h), (info.logo_right, right_col, int(h * 0.62))):
        if not data:
            continue
        try:
            im = PILImage.open(io.BytesIO(data))
            w = int(im.width * hh / im.height)
            xl = XLImage(io.BytesIO(data))
            xl.width, xl.height = w, hh
            ws.add_image(xl, f"{col}1")
        except Exception:
            pass


def type_cf(ws, rng):
    for t, c in LIGHT.items():
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[str(t)], fill=fill(c)))


def cond_cf(ws, rng):
    for name, _r, _u, col in COND:
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{name}"'], fill=fill(col)))


def qs(name):
    return "'" + name.replace("'", "''") + "'"


# --------------------------------------------------------------------------------------
# 6.  ONE SECTION SHEET (with Diagram band included)
# --------------------------------------------------------------------------------------
WIDTHS = {"A": 6, "B": 10, "C": 9, "D": 20, "E": 20, "F": 15, "G": 14, "H": 9, "I": 24, "J": 10, "K": 10, "L": 8, "M": 12,
          "N": 12, "O": 12, "P": 12, "Q": 12, "R": 13, "S": 13, "T": 20, "U": 3, "V": 14}


def build_sheet(wb, info, s, nsheets, lane_w, section_len=DEFAULT_SECTION_LEN):
    ws = wb.create_sheet(s["name"])
    for c, w in WIDTHS.items():
        ws.column_dimensions[c].width = w
    sgn, start = info.sign, info.start
    nA = max(1, len(s["anoms"]))
    a1 = 28
    a2 = a1 + nA - 1
    rSB = a2 + 2
    s1 = rSB + 2
    s2 = s1 + s["n"] - 1
    rST = s2 + 1
    rGB = rST + 2
    g1 = rGB + 2
    g2 = g1 + 7
    rGT = g2 + 1
    rFB = rGT + 2
    rF = rFB + 1

    # row 1-2 : logos / title / segment
    box(ws, "A1:T1", f"CRACK MAPPING  -  DNIT 006/2003-PRO  (defect-based IGG)  -  {int(section_len)} m sheet", F(15, True, NAVY), "FFFFFF", CEN,
        Border(bottom=Side(style="thick", color=NAVY)))
    ws.row_dimensions[1].height = 54
    add_logos(ws, info, "R")
    box(ws, "A2:T2", info.segment, F(11, True), PALE, LEFT)

    # row 3 / 4 : lane, side, chainage
    lab = F(10, True)
    box(ws, "A3:B3", "Road side", lab, GREY)
    box(ws, "C3:E3", f"{info.side}  -  chainage {'DECREASES' if info.sign < 0 else 'INCREASES'}  ({info.direction})",
        F(11, True, "C00000" if info.side == "LHS" else "1F4E9C"))
    box(ws, "F3", "Lane", lab, GREY); box(ws, "G3", f"Lane #{info.lane}", lab)
    box(ws, "H3:I3", "Lane width (m)", lab, GREY); box(ws, "J3", f"{lane_w:g}", lab, fmt="0.0#")
    box(ws, "K3:L3", "Survey date", lab, GREY); box(ws, "M3:N3", info.date, lab)
    box(ws, "O3:T3", f"Sheet {s['k'] + 1} / {nsheets}   |   Source: {info.file}   ({len(info.anoms)} anomaly rows)", F(8, False, "595959"), al=LEFT)
    box(ws, "A4:B4", "Start chainage (km+m)", lab, GREY)
    X0, X1 = "$V$3", "$V$4"
    kmf = lambda X: f'INT({X}/1000)&"+"&TEXT(MOD({X},1000),"000")'
    box(ws, "C4:E4", f'={kmf(X0)}&"   (= "&TEXT({X0},"#,##0")&" m)"', F(11, True))
    box(ws, "F4:G4", "End chainage (km+m)", lab, GREY)
    box(ws, "H4:J4", f'={kmf(X1)}&"   (= "&TEXT({X1},"#,##0")&" m)"', F(11, True))
    box(ws, "K4:L4", f"Distance of section start from survey start {km(info.start)} (m)", F(8, True), GREY)
    box(ws, "M4:N4", s["d0"], F(10, True, "0000FF"), fmt="#,##0")
    box(ws, "O4:Q4", '=HYPERLINK("#\'Summary\'!A1","\u25C0  Back to Summary")', F(10, True, "0563C1", u="single"))
    box(ws, "R4:T4", '=HYPERLINK("#\'Search\'!A1","\U0001F50D  Search a chainage")', F(10, True, "0563C1", u="single"))
    ws.row_dimensions[3].height, ws.row_dimensions[4].height = 21, 36

    # helper cells (column U/V)
    box(ws, "U2:V2", "calculation helpers", F(8, True, "7F7F7F"), None, CEN, Border())
    helpers = [("start (m)", f"={start}+{sgn}*$M$4", "#,##0"), ("end (m)", f"=$V$3+{sgn}*$V$5", "#,##0"),
               ("length (m)", round(s["L"], 3), "0.0##"), ("n stations", f"=ROUNDUP($V$5/{STATION_LEN},0)", "0")]
    for i, (t, v, f_) in enumerate(helpers):
        r = 3 + i
        ws[f"U{r}"] = None
        cc = ws[f"V{r}"]; cc.value = v; cc.number_format = f_
        cc.font = F(9, False, "0000FF" if i == 2 else "000000"); cc.border = BOX
        ws[f"W{r}"] = t; ws[f"W{r}"].font = F(8, False, "7F7F7F")
    ws.column_dimensions["W"].width = 11

    # row 6/7 : KPI
    for rng, h in (("A6:B6", "Detected anomalies"), ("C6:E6", "Stations affected (of n)"), ("F6:G6", "Stations affected %"),
                   ("H6:I6", "IGG (sum IGI)"), ("J6:N6", "Pavement condition"), ("O6:T6", "DNIT reference / calculation note")):
        box(ws, rng, h, F(10, True, "FFFFFF"), NAVY)
    cexpr = 'IF(H7<=20,"Very good",IF(H7<=40,"Good",IF(H7<=80,"Regular",IF(H7<=160,"Bad","Lousy"))))'
    box(ws, "A7:B7", f"=COUNT($B${a1}:$B${a2})", F(13, True))
    box(ws, "C7:E7", f'=$K${rST}&" of "&$H${rST}', F(13, True))
    box(ws, "F7:G7", f"=IF($V$6=0,0,$K${rST}/$V$6)", F(13, True), fmt="0.0%")
    box(ws, "H7:I7", f"=$K${rGT}", F(13, True), fmt="0.00")
    box(ws, "J7:N7", f"={cexpr}", F(13, True))
    box(ws, "O7:T7", f"IGG = sum(fr x fp). One {int(section_len)} m sheet evaluated in 20 m DNIT stations (n = {s['n']}).",
        F(8, False, "595959"), al=LEFT)
    cond_cf(ws, "J7:N7")
    ws.row_dimensions[6].height, ws.row_dimensions[7].height = 24, 36

    # diagram band + picture + legend explanation
    band(ws, 9, f"CHAINAGE DIAGRAM  -  {info.segment}  -  chainage {km(s['c0'])} to {km(s['c1'])}  ({info.direction}, {info.side})")
    for r in range(10, 25):
        ws.row_dimensions[r].height = 18
    if "png" in s:
        img = XLImage(io.BytesIO(s["png"]))
        img.width, img.height = 1380, 328
        ws.add_image(img, "A10")
    box(ws, "N10:T10", "HOW TO READ THE DIAGRAM", F(10, True, "FFFFFF"), NAVY)
    htr = [
        f"Top/middle = individual defects ('Where exactly is each defect?'). Simple vertical markers & ID labels with individual defect colours.",
        f"Bottom = station summary ('For this 20 m section, which defect category is the worst/selected one?').",
        f"Horizontal axis = chainage along the road ({info.direction}). Vertical lines every 20 m = one DNIT station (S1 ... S{s['n']}). Ticks every 5 m.",
        "Station bar (bottom): colour and code = worst DNIT defect category selected for that 20 m station (highest fp). Grey = nothing counted."
    ]
    for i, t in enumerate(htr):
        box(ws, f"N{11 + 2 * i}:T{12 + 2 * i}", t, F(9), None, LEFT)
    box(ws, "N19:T19", "COLOUR KEY - DNIT defect type (fp = weighting factor)", F(10, True, "FFFFFF"), NAVY)
    names8 = {1: "TTL isolated crack", 2: "J alligator", 3: "JE alligator + erosion", 4: "rutting", 5: "O pothole",
              6: "EX patch", 7: "D delamination", 8: "R ravelling"}
    for t in range(1, 9):
        r, col = 20 + (t - 1) % 4, (("N", "O", "Q") if t <= 4 else ("R", "S", "T"))
        box(ws, f"{col[0]}{r}", t, F(10, True), LIGHT[t])
        if t <= 4:
            box(ws, f"O{r}:Q{r}", f"Type {t} - {names8[t]} - fp {FP[t]:g}", F(9), al=LEFT)
        else:
            box(ws, f"S{r}:T{r}", f"{t} - {names8[t]} - fp {FP[t]:g}", F(9), al=LEFT)
    box(ws, "N24:T24", "Grey = no defect counted in that station (bleeding: fp 0, unknown names: not counted)", F(9, False, "595959"), GREY, LEFT)

    # anomaly table
    band(ws, 26, f"ANOMALY TABLE - every detected anomaly in this {int(section_len)} m section (blue = input)")
    heads = [("A", "No."), ("B", "Anomaly no."), ("C", "Lane"), ("D:E", "Image path (URL of the main sheet)"),
             ("F", "Anomaly chainage (km+m)"), ("G", "Distance from survey start (m)"), ("H", "Station"), ("I", "Type of anomaly (as in main sheet)"),
             ("J", "DNIT code"), ("K", "DNIT type"), ("L", "fp"), ("M:Q", "DNIT classification / basis"), ("R", "Latitude"),
             ("S", "Longitude"), ("T", "Processed video time")]
    for col, h in heads:
        box(ws, f"{col.split(':')[0]}27" + (f":{col.split(':')[1]}27" if ":" in col else ""), h, F(9, True, "FFFFFF"), NAVY)
    ws.row_dimensions[27].height = 40

    if not s["anoms"]:
        box(ws, f"A{a1}:T{a1}", f"No anomaly detected in this {int(section_len)} m section.", F(10, False, "595959", True), None, CEN)
    for i, a in enumerate(s["anoms"]):
        r = a1 + i
        ws.row_dimensions[r].height = 26
        box(ws, f"A{r}", i + 1, F(10)); box(ws, f"B{r}", a["no"], F(10))
        box(ws, f"C{r}", f"Lane #{info.lane}", F(9))
        cu = box(ws, f"D{r}:E{r}", a["url"], F(8, False, "0563C1", u="single"), None, LEFT)
        if a["url"].startswith("http"):
            cu.hyperlink = a["url"]
        box(ws, f"F{r}", f'=INT(({start}+{sgn}*G{r})/1000)&"+"&TEXT(MOD({start}+{sgn}*G{r},1000),"000.00")', F(10, True), fmt="@")
        box(ws, f"G{r}", round(a["dist"], 2), F(10, False, "0000FF"), fmt="#,##0.00")
        box(ws, f"H{r}", f'="S"&MAX(1,MIN($V$6,INT((G{r}-$M$4)/{STATION_LEN})+1))', F(10), al=CEN)
        box(ws, f"I{r}", a["raw"], F(10, True, "0000FF"), al=LEFT)
        box(ws, f"J{r}", a["code"], F(10, True, "008000"), al=CEN)
        box(ws, f"K{r}", a["type_no"] if a["type_no"] else "", F(10, True, "008000"), al=CEN)
        box(ws, f"L{r}", a["fp"] if a["fp"] is not None else "", F(10, False, "008000"), fmt="0.0", al=CEN)
        box(ws, f"M{r}:Q{r}", a["basis"], F(8, False, "008000"), al=LEFT)
        box(ws, f"R{r}", a["lat"], F(9), fmt="0.0000000"); box(ws, f"S{r}", a["lon"], F(9), fmt="0.0000000")
        box(ws, f"T{r}", a["ts"], F(9))
    type_cf(ws, f"K{a1}:K{a2}")

    # stations
    band(ws, rSB, "DNIT STATIONS - 20 m stations, the worst defect of each station is counted")
    for rng, h in (("A", "Station"), ("B:D", "Chainage range (km+m)"), ("E:F", "Anomaly no. in station"), ("G", "Anomalies in station"),
                   ("H", "Worst DNIT type"), ("I", "Defect code(s) of worst type"), ("J", "fp of worst type"), ("K", "Counted in FA?")):
        a_, b_ = (rng.split(":") + [None])[:2]
        box(ws, f"{a_}{rSB + 1}" + (f":{b_}{rSB + 1}" if b_ else ""), h, F(9, True, "FFFFFF"), NAVY)
    ws.row_dimensions[rSB + 1].height = 34
    Hr, Kr, Lr, Br = (f"$H${a1}:$H${a2}", f"$K${a1}:$K${a2}", f"$L${a1}:$L${a2}", f"$B${a1}:$B${a2}")
    codes_expr = ('"TLC, TLL, TTC, TTL, other isolated cracks", "J, TB", "JE, TBE", '
                  '"ALP, ATP, ALC, ATC", "O, P, E", "EX", "D", "R"')
    for i in range(s["n"]):
        r = s1 + i
        box(ws, f"A{r}", f"S{i + 1}", F(10, True), GREY)
        fr_ = f"$V$3+{sgn}*({i})*{STATION_LEN}"
        to_ = f"({fr_})+{sgn}*MIN({STATION_LEN},$V$5-({i})*{STATION_LEN})"
        box(ws, f"B{r}:D{r}", f'={kmf("("+fr_+")")}&" - "&{kmf("("+to_+")")}', F(10))
        box(ws, f"E{r}:F{r}", None, F(10))
        ws[f"E{r}"] = ArrayFormula(f"E{r}", f'=_xlfn.TEXTJOIN(", ",TRUE,IF({Hr}=A{r},{Br},""))')
        box(ws, f"G{r}", f"=COUNTIF({Hr},A{r})", F(10))
        box(ws, f"J{r}", f"=_xlfn.MAXIFS({Lr},{Hr},A{r})", F(10), fmt="0.0")
        box(ws, f"H{r}", f'=IF(J{r}=0,"-",_xlfn.MINIFS({Kr},{Hr},A{r},{Lr},J{r}))', F(10, True))
        box(ws, f"I{r}", f'=IF(ISNUMBER(H{r}),CHOOSE(H{r},{codes_expr}),"-")', F(9), al=CEN)
        box(ws, f"K{r}", f'=IF(J{r}=0,"No","Yes")', F(10))
    type_cf(ws, f"H{s1}:H{s2}")
    box(ws, f"A{rST}:F{rST}", "TOTALS \u2192  anomalies in stations  |  n = stations evaluated  |  stations affected", F(10, True), PALE, RIGHT)
    box(ws, f"G{rST}", f"=SUM(G{s1}:G{s2})", F(10, True), PALE); box(ws, f"H{rST}", "=$V$6", F(10, True), PALE)
    box(ws, f"I{rST}", "stations affected \u2192", F(9, True), PALE); box(ws, f"J{rST}", None, None, PALE)
    box(ws, f"K{rST}", f'=COUNTIF(K{s1}:K{s2},"Yes")', F(10, True), PALE)

    # IGG calculation table
    band(ws, rGB, "IGG CALCULATION - all 8 DNIT defect types (types with no station show FA = 0)")
    for rng, h in (("A", "Type"), ("B:D", "Defect codes"), ("E:F", "Description"), ("G", "fp weight"), ("H", "FA (stations)"),
                   ("I", "n (stations)"), ("J", "fr = FA/n x 100 (%)"), ("K", "IGI = fr x fp")):
        a_, b_ = (rng.split(":") + [None])[:2]
        box(ws, f"{a_}{rGB + 1}" + (f":{b_}{rGB + 1}" if b_ else ""), h, F(9, True, "FFFFFF"), NAVY)
    ws.row_dimensions[rGB + 1].height = 34
    desc_expr = ('"Isolated cracks (Type 1) - e.g. TTL long transverse, TTC transverse", '
                 '"FC-2 - alligator / block cracks WITHOUT edge erosion", '
                 '"FC-3 - alligator / block cracks WITH edge erosion", '
                 '"Plastic deformation (rutting) / corrugation", '
                 '"Pothole, deteriorated patch, edge break", '
                 '"Existing patch (any condition)", '
                 '"Surface disintegration / delamination", '
                 '"Raveling (loss of aggregate)"')
    fp_expr = "0.2, 0.5, 0.8, 0.9, 1.0, 0.5, 0.3, 0.6"
    for t in range(1, 9):
        r = g1 + t - 1
        box(ws, f"A{r}", t, F(10, True), LIGHT[t])
        box(ws, f"B{r}:D{r}", f"=CHOOSE(A{r},{codes_expr})", F(9, False, "008000"))
        box(ws, f"E{r}:F{r}", f"=CHOOSE(A{r},{desc_expr})", F(8, False, "008000"), al=LEFT)
        box(ws, f"G{r}", f"=CHOOSE(A{r},{fp_expr})", F(10, False, "008000"), fmt="0.0")
        box(ws, f"H{r}", f"=COUNTIF($H${s1}:$H${s2},A{r})", F(10))
        box(ws, f"I{r}", "=$V$6", F(10)); box(ws, f"J{r}", f"=H{r}/I{r}*100", F(10), fmt="0.00")
        box(ws, f"K{r}", f"=J{r}*G{r}", F(10, True), fmt="0.00")
        ws.row_dimensions[r].height = 24
    box(ws, f"A{rGT}:F{rGT}", "Total", F(10, True), PALE, RIGHT)
    for col, fm, nf in (("G", None, None), ("H", f"=SUM(H{g1}:H{g2})", "0"), ("I", "=$V$6", "0"), ("J", f"=SUM(J{g1}:J{g2})", "0.00"),
                        ("K", f"=SUM(K{g1}:K{g2})", "0.00")):
        box(ws, f"{col}{rGT}", fm, F(10, True), PALE, fmt=nf)

    # final result
    band(ws, rFB, f"FINAL RESULT FOR THIS {int(section_len)} m SECTION")
    box(ws, f"A{rF}:C{rF}", "IGG = \u03A3 IGI", F(13, True, "FFFFFF"), NAVY)
    box(ws, f"D{rF}", "=$H$7", F(13, True), fmt="0.00")
    box(ws, f"E{rF}", "=$J$7", F(13, True)); cond_cf(ws, f"E{rF}")
    box(ws, f"F{rF}:G{rF}", "Stations affected", F(10, True), GREY); box(ws, f"H{rF}:I{rF}", "=$F$7", F(13, True), fmt="0.0%")
    ws.row_dimensions[rF].height = 30
    ws.sheet_properties.tabColor = s["cond_col"]
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.sheet_view.showGridLines = False
    s["rows"] = dict(a1=a1, a2=a2, s1=s1, s2=s2, rST=rST, rGT=rGT)
    return ws


# --------------------------------------------------------------------------------------
# 7.  SUMMARY + SEARCH
# --------------------------------------------------------------------------------------
def build_summary(wb, info, secs, section_len=DEFAULT_SECTION_LEN):
    ws = wb.create_sheet("Summary", 0)
    ncols = 8
    last = get_column_letter(ncols)
    widths = [6, 22, 26, 11, 15, 13, 12, 13]
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(i + 1)].width = w
    box(ws, f"A1:{last}1", f"CRACK MAPPING  -  LANE {info.lane} {info.side}  |  SUMMARY & DNIT IGG  (every {int(section_len)} m)", F(15, True, NAVY), "FFFFFF", CEN,
        Border(bottom=Side(style="thick", color=NAVY)))
    ws.row_dimensions[1].height = 54
    add_logos(ws, info, "H")
    box(ws, f"A2:{last}2", info.segment, F(11, True), PALE, LEFT)
    items = [("A3:B3", "Road side", "A4:B4", info.side), ("C3", "Lane / direction", "C4", f"Lane #{info.lane}  -  {info.direction}"),
             ("D3", "Start chainage", "D4", f'=INT({info.start}/1000)&"+"&TEXT(MOD({info.start},1000),"000")'),
             ("E3", "End chainage", "E4", f'=INT({info.end}/1000)&"+"&TEXT(MOD({info.end},1000),"000")'),
             ("F3", "Survey length (m)", "F4", info.total), ("G3", "Survey date", "G4", info.date),
             ("H3", f"{int(section_len)} m sheets", "H4", info.nsec)]
    for lr, lt, vr, vv in items:
        box(ws, lr, lt, F(9, True), GREY); box(ws, vr, vv, F(11, True), fmt="#,##0" if isinstance(vv, (int, float)) else None)
    ws["A4"].font = F(12, True, "C00000" if info.side == "LHS" else "1F4E9C")
    ws.row_dimensions[4].height = 26
    box(ws, f"A5:{last}5", f"Chainage {'DECREASES' if info.sign < 0 else 'INCREASES'} along this survey ({info.side}).  Click a sheet name to open it.  "
        "Use Search sheet to jump to any chainage.", F(9, False, "595959", True), None, LEFT, Border())
    hr = 6
    heads = ["#", f"{int(section_len)} m section (sheet)", "Chainage (km+m)", "Detected anomalies", "Stations affected (of n)",
             "Stations affected %", "IGG (sum IGI)", "Condition"]
    for i, h in enumerate(heads):
        box(ws, f"{get_column_letter(i + 1)}{hr}", h, F(9, True, "FFFFFF"), NAVY)
    ws.row_dimensions[hr].height = 42
    r0 = hr + 1
    for s in secs:
        r = r0 + s["k"]
        q = qs(s["name"]); R = s["rows"]
        box(ws, f"A{r}", s["k"] + 1, F(10), GREY)
        box(ws, f"B{r}", f'=HYPERLINK("#{q}!A1","{s["name"]}")', F(10, True, "0563C1", u="single"), al=LEFT)
        box(ws, f"C{r}", f'={q}!$C$4&"  to  "&{q}!$H$4', F(9), al=LEFT)
        ws[f"C{r}"].value = f'=INT({q}!$V$3/1000)&"+"&TEXT(MOD({q}!$V$3,1000),"000")&"  \u2192  "&INT({q}!$V$4/1000)&"+"&TEXT(MOD({q}!$V$4,1000),"000")'
        box(ws, f"D{r}", f"={q}!$A$7", F(10))
        box(ws, f"E{r}", f'=COUNTIF({q}!$K${R["s1"]}:$K${R["s2"]},"Yes")&" of "&{q}!$V$6', F(10))
        box(ws, f"F{r}", f"={q}!$F$7", F(10), fmt="0.0%")
        box(ws, f"G{r}", f"={q}!$H$7", F(10, True), fmt="0.00"); box(ws, f"H{r}", f"={q}!$J$7", F(10, True))
    rl = r0 + len(secs) - 1
    rt = rl + 1
    box(ws, f"A{rt}:C{rt}", f"TOTAL  ({km(info.start)} \u2192 {km(info.end)})", F(10, True), PALE, RIGHT)
    box(ws, f"D{rt}", f"=SUM(D{r0}:D{rl})", F(10, True), PALE)
    box(ws, f"E{rt}", None, None, PALE)
    box(ws, f"F{rt}", f"=AVERAGE(F{r0}:F{rl})", F(10, True), PALE, fmt="0.0%")
    box(ws, f"G{rt}", f"=AVERAGE(G{r0}:G{rl})", F(10, True), PALE, fmt="0.00")
    box(ws, f"H{rt}", "average", F(9, True), PALE)
    cond_cf(ws, f"H{r0}:H{rl}")
    ws.freeze_panes = f"C{r0}"
    ws.auto_filter.ref = f"A{hr}:{last}{rl}"

    # condition distribution
    r = rt + 3
    box(ws, f"A{r}:H{r}", f"CONDITION OF THE ROAD - how many {int(section_len)} m sheets fall in each class", F(11, True, "FFFFFF"), BLUE, LEFT)
    for i, h in enumerate(("#", "Condition", "IGG range", "Sheets", "% of sheets", "Length (m)")):
        box(ws, f"{get_column_letter(i + 1)}{r + 1}", h, F(9, True, "FFFFFF"), NAVY)
    for i, (nm, rng, up, col) in enumerate(COND):
        rr = r + 2 + i
        box(ws, f"A{rr}", i + 1, F(10)); box(ws, f"B{rr}", nm, F(10, True), col, LEFT); box(ws, f"C{rr}", rng, F(10), al=LEFT)
        box(ws, f"D{rr}", f'=COUNTIF($H${r0}:$H${rl},B{rr})', F(10, True))
        box(ws, f"E{rr}", f"=IF({len(secs)}=0,0,D{rr}/{len(secs)})", F(10), fmt="0.0%")
        box(ws, f"F{rr}", f"=D{rr}*{section_len:g}", F(10), fmt="#,##0")

    # type totals
    r = r + 2 + len(COND) + 2
    box(ws, f"A{r}:H{r}", "DNIT defect types found in this survey", F(11, True, "FFFFFF"), BLUE, LEFT)
    for i, h in enumerate(("Type", "Codes", "Found in main sheet as", "fp")):
        box(ws, f"{get_column_letter(i + 1)}{r + 1}", h, F(9, True, "FFFFFF"), NAVY)
    found = {}
    for nm, _ in info.names:
        c = classify(nm)
        if c and c[1] <= 8:
            found.setdefault(c[1], []).append(f"{nm} ({c[0]})")
    for t in range(1, 9):
        rr = r + 1 + t
        box(ws, f"A{rr}", t, F(10, True), LIGHT[t])
        box(ws, f"B{rr}", TYPES[t][0], F(9, False, "008000"), al=LEFT)
        box(ws, f"C{rr}", ", ".join(found.get(t, [])) or "not found", F(9), al=LEFT)
        box(ws, f"D{rr}", TYPES[t][2], F(10, False, "008000"), fmt="0.0")
    ws.sheet_properties.tabColor = NAVY
    ws.sheet_view.showGridLines = False
    return ws, r0, rl


def build_search(wb, info, secs, r0, rl, section_len=DEFAULT_SECTION_LEN):
    ws = wb.create_sheet("Search", 1)
    for c, w in zip("ABCDEFGH", (4, 34, 44, 4, 26, 30, 4, 4)):
        ws.column_dimensions[c].width = w
    box(ws, "A1:H1", f"SEARCH  -  type a chainage and jump straight to its {int(section_len)} m sheet", F(15, True, NAVY), "FFFFFF", CEN,
        Border(bottom=Side(style="thick", color=NAVY)))
    ws.row_dimensions[1].height = 54
    add_logos(ws, info, "F", 62)
    box(ws, "A2:H2", info.segment, F(11, True), PALE, LEFT)
    box(ws, "B4:C4", "HOW TO USE", F(10, True, "FFFFFF"), NAVY)
    box(ws, "B5:C5", "Type a chainage in the yellow cell:  km+m  (example 200+350)  or metres  (example 200350).\n"
        "Or type a sheet number in the second yellow cell. The matching sheet is shown below - click the blue link to open it.",
        F(9), None, LEFT)
    ws.row_dimensions[5].height = 52
    box(ws, "B7", "Chainage (km+m or metres)", F(10, True), GREY, LEFT)
    box(ws, "C7", None, F(14, True, "0000FF"), "FFFF99", CEN, Border(*(Side(style="medium", color="C00000"),) * 4))
    box(ws, "B8", "...or sheet number (1 - %d)" % len(secs), F(10, True), GREY, LEFT)
    box(ws, "C8", None, F(14, True, "0000FF"), "FFFF99", CEN, Border(*(Side(style="medium", color="C00000"),) * 4))
    ws.row_dimensions[7].height = ws.row_dimensions[8].height = 30
    box(ws, "B10", "Survey range", F(9, True), GREY, LEFT)
    box(ws, "C10", f'=INT({info.start}/1000)&"+"&TEXT(MOD({info.start},1000),"000")&"   \u2192   "&INT({info.end}/1000)&"+"&TEXT(MOD({info.end},1000),"000")&"   ({info.side}, chainage {"decreases" if info.sign < 0 else "increases"})"', F(10, True), al=LEFT)
    # helper cells
    box(ws, "E4:F4", "calculation (do not edit)", F(8, True, "7F7F7F"), None, CEN, Border())
    box(ws, "E5", "chainage in metres", F(8, False, "7F7F7F"), None, RIGHT, Border())
    box(ws, "F5", '=IF(C7="","",IF(ISNUMBER(C7),C7,IFERROR(VALUE(LEFT(C7,FIND("+",C7)-1))*1000+VALUE(MID(C7,FIND("+",C7)+1,10)),"")))', F(9), border=BOX, fmt="#,##0.###")
    box(ws, "E6", "metres along the survey", F(8, False, "7F7F7F"), None, RIGHT, Border())
    box(ws, "F6", f'=IF(F5="","",{info.sign}*(F5-{info.start}))', F(9), border=BOX, fmt="#,##0.###")
    box(ws, "E7", "sheet number", F(8, False, "7F7F7F"), None, RIGHT, Border())
    box(ws, "F7", f'=IF(AND(ISNUMBER(C8),C8>=1,C8<={len(secs)}),INT(C8),IF(AND(ISNUMBER(F6),F6>=0,F6<={info.total}),MIN({len(secs)},INT(F6/{section_len:g})+1),""))',
        F(9, True), border=BOX)
    box(ws, "B12:C12", '=IF(AND(C7="",C8=""),"\u2192  Enter a chainage in the yellow cell",IF(F7="","\u26A0  This chainage is outside the survey range","\u2714  Sheet found"))',
        F(12, True, "FFFFFF"), NAVY)
    ws.row_dimensions[12].height = 28
    idx = "$F$7"
    S = lambda col: f"INDEX(Summary!${col}${r0}:${col}${rl},{idx})"
    rows = [("Sheet", f'=IF({idx}="","",HYPERLINK("#\'"&{S("B")}&"\'!A1","\u25B6  Open  "&{S("B")}))', F(12, True, "0563C1", u="single")),
            ("Chainage of the sheet", f'=IF({idx}="","",{S("C")})', F(11, True)),
            ("Detected anomalies", f'=IF({idx}="","",{S("D")})', F(11, True)),
            ("Stations affected", f'=IF({idx}="","",{S("E")})', F(11, True)),
            ("Stations affected %", f'=IF({idx}="","",{S("F")})', F(11, True)),
            ("IGG (sum IGI)", f'=IF({idx}="","",{S("G")})', F(11, True)),
            ("Pavement condition", f'=IF({idx}="","",{S("H")})', F(12, True))]
    for i, (lab, fm, fo) in enumerate(rows):
        r = 14 + i
        box(ws, f"B{r}", lab, F(10, True), GREY, LEFT)
        box(ws, f"C{r}", fm, fo, None, CEN, fmt="0.0%" if "%" in lab else ("0.00" if "IGG" in lab else None))
        ws.row_dimensions[r].height = 26
    cond_cf(ws, "C20")
    ws.sheet_properties.tabColor = "FFC000"
    ws.sheet_view.showGridLines = False
    return ws


# --------------------------------------------------------------------------------------
# 9.  BUILD THE WORKBOOK
# --------------------------------------------------------------------------------------
def generate(info, out_path, lane_w=LANE_WIDTH, section_len=DEFAULT_SECTION_LEN, progress=None, cancel=lambda: False, recalc=False):
    P = progress or (lambda p, t: None)
    P(7, "Step 1/3  Mapping every anomaly to DNIT types ...")
    secs = sections(info, section_len)
    assign(info, secs, section_len)
    for s in secs:
        analyse(s)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    wb.calculation.fullCalcOnLoad = True
    n = len(secs)
    for s in secs:
        if cancel():
            raise RuntimeError("Cancelled")
        s["png"] = draw_diagram(info, s, lane_w, section_len)
        build_sheet(wb, info, s, n, lane_w, section_len)
        s.pop("png")
        if s["k"] % 5 == 0 or s["k"] == n - 1:
            P(10 + int(80 * (s["k"] + 1) / n), f"Step 2/3  Building {int(section_len)} m sheets + diagrams  ({s['k'] + 1} / {n})  {s['name']}")
    P(91, "Step 3/3  Summary sheet ...")
    ws_sum, r0, rl = build_summary(wb, info, secs, section_len)
    P(94, "Step 3/3  Search sheet ...")
    build_search(wb, info, secs, r0, rl, section_len)
    order = ["Summary", "Search"] + [s["name"] for s in secs]
    wb._sheets = [wb[nm] for nm in order]
    wb.active = 0
    P(96, "Step 3/3  Saving workbook ...")
    wb.save(out_path)
    if recalc:
        P(98, "Recalculating with LibreOffice ...")
        recalc_libreoffice(out_path)
    P(100, "Done.")
    stats = dict(sheets=n, anoms=len(info.anoms), classes={c[0]: sum(1 for s in secs if s["cond"] == c[0]) for c in COND},
                 mean_igg=sum(s["igg"] for s in secs) / n, secs=secs)
    return stats


def find_soffice():
    for c in ("soffice", "libreoffice", r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe", "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        p = shutil.which(c) or (c if os.path.exists(c) else None)
        if p:
            return p
    return None


def recalc_libreoffice(path):
    so = find_soffice()
    if not so:
        return False
    tmp = tempfile.mkdtemp()
    subprocess.run([so, "--headless", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", tmp, path],
                   check=True, timeout=1800, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    out = os.path.join(tmp, os.path.basename(path))
    if os.path.exists(out):
        shutil.move(out, path)
    return True


# --------------------------------------------------------------------------------------
# 10.  BATCH PROCESSOR FOR DIRECTORIES OR SINGLE FILES
# --------------------------------------------------------------------------------------
def process_single_or_batch(input_path, output_dir_or_file=None, side_override="Auto", section_len=DEFAULT_SECTION_LEN, lane_w=LANE_WIDTH, recalc=False, progress=None):
    P = progress or (lambda p, t: None)
    if os.path.isdir(input_path):
        excel_files = [os.path.join(input_path, f) for f in os.listdir(input_path) if f.endswith(".xlsx") and not f.startswith("~$") and not f.endswith("_DNIT_report.xlsx")]
        if not excel_files:
            raise ValueError(f"No valid .xlsx report files found in directory '{input_path}'")
        out_dir = output_dir_or_file or os.path.join(input_path, "DNIT_Output_Reports")
        os.makedirs(out_dir, exist_ok=True)

        results = []
        total_files = len(excel_files)
        for idx, fpath in enumerate(excel_files):
            fname = os.path.splitext(os.path.basename(fpath))[0]
            out_path = os.path.join(out_dir, f"{fname}_DNIT_{int(section_len)}m_report.xlsx")
            P(int(100 * idx / total_files), f"Processing file {idx+1}/{total_files}: {os.path.basename(fpath)}")
            inf = parse_report(fpath, side_override, section_len)
            st = generate(inf, out_path, lane_w, section_len, recalc=recalc)
            results.append((fpath, out_path, st))
        P(100, f"Batch processing completed! {len(results)} files processed.")
        return results
    else:
        out_path = output_dir_or_file
        if not out_path or os.path.isdir(out_path):
            base_dir = output_dir_or_file if (out_path and os.path.isdir(out_path)) else os.path.dirname(input_path)
            fname = os.path.splitext(os.path.basename(input_path))[0]
            out_path = os.path.join(base_dir, f"{fname}_DNIT_{int(section_len)}m_report.xlsx")
        inf = parse_report(input_path, side_override, section_len, progress=P)
        st = generate(inf, out_path, lane_w, section_len, recalc=recalc, progress=P)
        return [(input_path, out_path, st)]


# --------------------------------------------------------------------------------------
# 11.  GUI WITH SECTION LENGTH & BATCH SUPPORT
# --------------------------------------------------------------------------------------
def run_gui():
    from PyQt5.QtCore import Qt, QThread, pyqtSignal
    from PyQt5.QtWidgets import (QApplication, QCheckBox, QComboBox, QDoubleSpinBox, QSpinBox, QFileDialog, QGroupBox, QHBoxLayout,
                                 QLabel, QLineEdit, QMainWindow, QMessageBox, QPlainTextEdit, QProgressBar, QPushButton,
                                 QVBoxLayout, QWidget)

    class BatchWorker(QThread):
        prog = pyqtSignal(int, str)
        done = pyqtSignal(object)
        fail = pyqtSignal(str)

        def __init__(self, in_path, out_path, side, sec_len, lane_w, recalc):
            super().__init__()
            self.in_path, self.out_path, self.side = in_path, out_path, side
            self.sec_len, self.lane_w, self.recalc = sec_len, lane_w, recalc

        def run(self):
            try:
                res = process_single_or_batch(
                    self.in_path, self.out_path, self.side, self.sec_len, self.lane_w, self.recalc,
                    progress=lambda p, t: self.prog.emit(int(p), t)
                )
                self.done.emit(res)
            except Exception as e:
                self.fail.emit(str(e))

    class Win(QMainWindow):
        def __init__(self):
            super().__init__()
            self.setWindowTitle("Road Athena - DNIT Report Builder")
            self.resize(980, 860)
            self.input_path = ""
            self.out_path = ""
            self.setStyleSheet("""
                QMainWindow, QWidget {background:#F4F6FA; font-family:Segoe UI, Arial; font-size:13px;}
                QGroupBox {font-weight:bold; border:1px solid #C9D2E3; border-radius:8px; margin-top:12px; padding:10px; background:white;}
                QGroupBox::title {subcontrol-origin: margin; left:12px; padding:0 6px; color:#1F2A44;}
                QPushButton {background:#2F5597; color:white; border-radius:6px; padding:8px 16px; font-weight:bold;}
                QPushButton:hover {background:#1F4E9C;} QPushButton:disabled {background:#A9B4CC;}
                QProgressBar {border:1px solid #C9D2E3; border-radius:6px; text-align:center; height:24px; background:white;}
                QProgressBar::chunk {background:#2E9D5B; border-radius:5px;}
                QLineEdit, QComboBox, QDoubleSpinBox, QSpinBox {padding:5px; border:1px solid #C9D2E3; border-radius:5px; background:white;}
                QPlainTextEdit {background:#101828; color:#D6E2FF; font-family:Consolas, monospace; font-size:12px; border-radius:6px;}
            """)
            root = QWidget(); self.setCentralWidget(root)
            v = QVBoxLayout(root); v.setSpacing(8)
            t = QLabel("ROAD ATHENA  \u2022  DNIT 006/2003-PRO REPORT BUILDER")
            t.setStyleSheet("background:#1F2A44;color:white;padding:14px;font-size:18px;font-weight:bold;border-radius:8px;")
            v.addWidget(t)

            g1 = QGroupBox("STEP 1  -  Choose Single Excel File OR Folder for Batch Processing")
            h = QHBoxLayout(g1)
            self.pathEdit = QLineEdit(); self.pathEdit.setReadOnly(True); self.pathEdit.setPlaceholderText("No file or folder selected")
            b_file = QPushButton("Browse File ..."); b_file.clicked.connect(self.pick_file)
            b_dir = QPushButton("Browse Folder ..."); b_dir.clicked.connect(self.pick_dir)
            h.addWidget(self.pathEdit, 1); h.addWidget(b_file); h.addWidget(b_dir)
            v.addWidget(g1)

            g2 = QGroupBox("STEP 2  -  Report Configuration & Parameters")
            vv = QVBoxLayout(g2)
            hh = QHBoxLayout()
            hh.addWidget(QLabel("Section Length (m):"))
            self.secLenBox = QSpinBox(); self.secLenBox.setRange(20, 1000); self.secLenBox.setSingleStep(20); self.secLenBox.setValue(100)
            hh.addWidget(self.secLenBox)
            hh.addSpacing(20); hh.addWidget(QLabel("Road Side:"))
            self.sideBox = QComboBox(); self.sideBox.addItems(["Auto", "RHS", "LHS"])
            hh.addWidget(self.sideBox)
            hh.addSpacing(20); hh.addWidget(QLabel("Lane Width (m):"))
            self.laneW = QDoubleSpinBox(); self.laneW.setRange(2.0, 8.0); self.laneW.setSingleStep(0.1); self.laneW.setValue(LANE_WIDTH)
            hh.addWidget(self.laneW); hh.addStretch(1)
            vv.addLayout(hh)
            v.addWidget(g2)

            g3 = QGroupBox("STEP 3  -  Generate Report(s)")
            vv = QVBoxLayout(g3)
            h = QHBoxLayout()
            self.outEdit = QLineEdit(); self.outEdit.setPlaceholderText("Output file or folder path...")
            b2 = QPushButton("Select Output ..."); b2.clicked.connect(self.pick_out)
            h.addWidget(self.outEdit, 1); h.addWidget(b2); vv.addLayout(h)
            self.chkRecalc = QCheckBox("After saving, recalculate formulas with LibreOffice")
            self.chkRecalc.setEnabled(find_soffice() is not None)
            vv.addWidget(self.chkRecalc)
            h = QHBoxLayout()
            self.btnGo = QPushButton("\u25B6  GENERATE REPORT(S)"); self.btnGo.clicked.connect(self.go)
            h.addWidget(self.btnGo); vv.addLayout(h)
            self.bar = QProgressBar(); self.bar.setRange(0, 100)
            self.step = QLabel("Ready ..."); self.step.setStyleSheet("font-weight:bold;color:#1F2A44;")
            vv.addWidget(self.step); vv.addWidget(self.bar)
            v.addWidget(g3)

            self.log = QPlainTextEdit(); self.log.setReadOnly(True); self.log.setMaximumBlockCount(500)
            v.addWidget(self.log, 1)

        def say(self, text):
            self.log.appendPlainText(datetime.now().strftime("%H:%M:%S  ") + text)

        def pick_file(self):
            p, _ = QFileDialog.getOpenFileName(self, "Select Pavement Report File", "", "Excel (*.xlsx *.xlsm)")
            if p:
                self.input_path = p
                self.pathEdit.setText(p)
                slen = self.secLenBox.value()
                self.outEdit.setText(os.path.splitext(p)[0] + f"_DNIT_{slen}m_report.xlsx")
                self.say(f"Selected file: {p}")

        def pick_dir(self):
            d = QFileDialog.getExistingDirectory(self, "Select Folder Containing Excel Reports")
            if d:
                self.input_path = d
                self.pathEdit.setText(d)
                self.outEdit.setText(os.path.join(d, "DNIT_Output_Reports"))
                self.say(f"Selected folder for batch processing: {d}")

        def pick_out(self):
            if os.path.isdir(self.input_path):
                d = QFileDialog.getExistingDirectory(self, "Select Output Folder")
                if d: self.outEdit.setText(d)
            else:
                p, _ = QFileDialog.getSaveFileName(self, "Save Output Excel", self.outEdit.text(), "Excel (*.xlsx)")
                if p: self.outEdit.setText(p)

        def go(self):
            if not self.pathEdit.text():
                QMessageBox.warning(self, "No Input", "Please select a file or folder first.")
                return
            self.btnGo.setEnabled(False)
            side = self.sideBox.currentText()
            sec_len = float(self.secLenBox.value())
            lane_w = self.laneW.value()
            recalc = self.chkRecalc.isChecked()

            self.worker = BatchWorker(self.pathEdit.text(), self.outEdit.text(), side, sec_len, lane_w, recalc)
            self.worker.prog.connect(lambda p, t: (self.bar.setValue(p), self.step.setText(t), self.say(t)))
            self.worker.done.connect(self.finished)
            self.worker.fail.connect(lambda e: (self.say("ERROR: " + e), QMessageBox.critical(self, "Error", e), self.btnGo.setEnabled(True)))
            self.worker.start()

        def finished(self, results):
            self.btnGo.setEnabled(True)
            self.step.setText("Completed successfully!")
            self.say(f"Successfully processed {len(results)} report(s).")
            QMessageBox.information(self, "Success", f"Processing completed successfully!\nGenerated {len(results)} file(s).")

    app = QApplication(sys.argv)
    w = Win()
    w.show()
    sys.exit(app.exec_())


# --------------------------------------------------------------------------------------
# 12.  MAIN ENTRY POINT (CLI / GUI)
# --------------------------------------------------------------------------------------
if __name__ == "__main__":
    if len(sys.argv) >= 2 and not sys.argv[1].startswith("-"):
        in_p = sys.argv[1]
        out_p = sys.argv[2] if len(sys.argv) >= 3 else None
        side = sys.argv[3] if len(sys.argv) >= 4 else "Auto"
        slen = float(sys.argv[4]) if len(sys.argv) >= 5 else DEFAULT_SECTION_LEN
        print(f"Running processing on: {in_p} (Section Length={slen}m)")
        res = process_single_or_batch(in_p, out_p, side_override=side, section_len=slen, progress=lambda p, t: print(f"[{p:3d}%] {t}"))
        print(f"Completed! {len(res)} file(s) generated.")
    else:
        run_gui()
