"""
build_excel.py
Generates SX_Steady_State_Model_UThHfZr.xlsx
Primene JMT / dual-acid SX model: H2SO4 loading, HNO3 stripping
Elements: U, Th, Hf, Zr   |   openpyxl only
"""

from openpyxl import Workbook
from openpyxl.styles import (PatternFill, Font, Alignment, Border, Side,
                              GradientFill)
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.chart.series import SeriesLabel
from openpyxl.formatting.rule import ColorScaleRule
import re

# ── Colour palette ────────────────────────────────────────────────────────────
C = {
    "INPUT_BG":  "FFFDE7", "INPUT_FG":  "1A237E",
    "CALC_BG":   "FFFFFF", "CALC_FG":   "212121",
    "STEP_BG":   "E8EAF6", "STEP_FG":   "212121",
    "OUTPUT_BG": "F1F8E9", "OUTPUT_FG": "2E7D32",
    "REF_BG":    "E3F2FD", "REF_FG":    "0D47A1",
    "TITLE":     "1B3A6B",
    "SECHDR":    "2E4057",
    "COLHDR":    "4A6FA5",
    "EXT_TINT":  "FFFDE7",
    "SCR_TINT":  "F3E5F5",
    "STR_TINT":  "E8F5E9",
    "HDR_FG":    "FFFFFF",
    "WARN":      "FF5722",
}

def fill(hex_bg):
    return PatternFill("solid", fgColor=hex_bg)

def font(hex_fg, bold=False, sz=10):
    return Font(color=hex_fg, bold=bold, size=sz, name="Calibri")

def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def left():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)

def thin_border():
    s = Side(style="thin", color="BDBDBD")
    return Border(left=s, right=s, top=s, bottom=s)

def style_cell(ws, row, col, bg, fg, bold=False, align="center", sz=10,
               border=True, val=None):
    c = ws.cell(row=row, column=col)
    if val is not None:
        c.value = val
    c.fill = fill(bg)
    c.font = font(fg, bold=bold, sz=sz)
    c.alignment = center() if align == "center" else left()
    if border:
        c.border = thin_border()
    return c

def hdr(ws, row, col, text, span=1, bg=None, fg=None, sz=10, bold=True):
    bg = bg or C["COLHDR"]
    fg = fg or C["HDR_FG"]
    c = ws.cell(row=row, column=col, value=text)
    c.fill = fill(bg)
    c.font = font(fg, bold=bold, sz=sz)
    c.alignment = center()
    c.border = thin_border()
    if span > 1:
        ws.merge_cells(start_row=row, start_column=col,
                       end_row=row, end_column=col+span-1)
    return c

def inp(ws, row, col, val=None):
    c = ws.cell(row=row, column=col)
    if val is not None:
        c.value = val
    c.fill = fill(C["INPUT_BG"])
    c.font = font(C["INPUT_FG"])
    c.alignment = center()
    c.border = thin_border()
    return c

def out(ws, row, col, formula=None):
    c = ws.cell(row=row, column=col)
    if formula is not None:
        c.value = formula
    c.fill = fill(C["OUTPUT_BG"])
    c.font = font(C["OUTPUT_FG"])
    c.alignment = center()
    c.border = thin_border()
    return c

def calc(ws, row, col, formula=None):
    c = ws.cell(row=row, column=col)
    if formula is not None:
        c.value = formula
    c.fill = fill(C["CALC_BG"])
    c.font = font(C["CALC_FG"])
    c.alignment = center()
    c.border = thin_border()
    return c

def step(ws, row, col, formula=None):
    c = ws.cell(row=row, column=col)
    if formula is not None:
        c.value = formula
    c.fill = fill(C["STEP_BG"])
    c.font = font(C["STEP_FG"])
    c.alignment = center()
    c.border = thin_border()
    return c

def ref_cell(ws, row, col, formula=None):
    c = ws.cell(row=row, column=col)
    if formula is not None:
        c.value = formula
    c.fill = fill(C["REF_BG"])
    c.font = font(C["REF_FG"])
    c.alignment = center()
    c.border = thin_border()
    return c

def cr(row, col):
    """Return absolute cell address string e.g. $B$5"""
    return f"${get_column_letter(col)}${row}"

def addr(row, col):
    return f"{get_column_letter(col)}{row}"

# ── Workbook setup ─────────────────────────────────────────────────────────────
wb = Workbook()
ws = wb.active
ws.title = "SX Model"
ws.sheet_properties.tabColor = C["TITLE"]

ELEMENTS = ["U", "Th", "Hf", "Zr"]
N_STAGES = 25

# Column widths
for col in range(1, 120):
    ws.column_dimensions[get_column_letter(col)].width = 13
ws.column_dimensions["A"].width = 18
ws.column_dimensions["B"].width = 16

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ROW 1 — Main title banner
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ws.row_dimensions[1].height = 30
ws.merge_cells("A1:AZ1")
c = ws["A1"]
c.value = "STEADY-STATE SX SIMULATION — Primene JMT / H₂SO₄ Loading / HNO₃ Stripping — U, Th, Hf, Zr"
c.fill = fill(C["TITLE"])
c.font = Font(color="FFFFFF", bold=True, size=14, name="Calibri")
c.alignment = center()

# ROW 2 — Legend
ws.row_dimensions[2].height = 18
legend = [
    ("INPUT (yellow)", "FFFDE7", C["INPUT_FG"]),
    ("CALC (white)", "FFFFFF", C["CALC_FG"]),
    ("STEP (blue-grey)", "E8EAF6", C["STEP_FG"]),
    ("OUTPUT (mint)", "F1F8E9", C["OUTPUT_FG"]),
    ("REF (light blue)", "E3F2FD", C["REF_FG"]),
]
col = 1
for label, bg, fg in legend:
    ws.merge_cells(start_row=2, start_column=col, end_row=2, end_column=col+3)
    c = ws.cell(row=2, column=col, value=label)
    c.fill = fill(bg)
    c.font = font(fg, bold=True, sz=9)
    c.alignment = center()
    col += 4

# ROW 3 — spacer
ws.row_dimensions[3].height = 6

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SECTION 0 — P/Q REGRESSION  (rows 4–55)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
REG_START = 4

# Section header
ws.row_dimensions[REG_START].height = 20
ws.merge_cells(f"A{REG_START}:Z{REG_START}")
c = ws.cell(row=REG_START, column=1,
            value="SECTION 0 — P/Q REGRESSION  (D = P × [acid]^Q  via log-log linear regression)")
c.fill = fill(C["SECHDR"])
c.font = font("FFFFFF", bold=True, sz=12)
c.alignment = center()

# Sub-panel headers
# Left panel: LOADING (H2SO4), Right panel: STRIPPING (HNO3)
# LOADING starts col 1, STRIPPING starts col 14
LOAD_COL = 1
STRIP_COL = 14

PANEL_COLS = ["Stage", "[acid] N", "Flow_aq\n(mL/min)", "Flow_org\n(mL/min)",
              "U_aq\n(ppm)", "Th_aq\n(ppm)", "Hf_aq\n(ppm)", "Zr_aq\n(ppm)",
              "U_org\n(ppm)", "Th_org\n(ppm)", "Hf_org\n(ppm)", "Zr_org\n(ppm)"]

STAGE_LABELS_LOAD = ["L1", "L2", "L3"]
STAGE_LABELS_STRIP = ["S1", "S2", "S3"]

# Panel sub-headers row
phdr_row = REG_START + 1
ws.row_dimensions[phdr_row].height = 18
ws.merge_cells(start_row=phdr_row, start_column=LOAD_COL,
               end_row=phdr_row, end_column=LOAD_COL+11)
hdr(ws, phdr_row, LOAD_COL, "LOADING DATA  (H₂SO₄ system)", span=12,
    bg=C["EXT_TINT"], fg=C["SECHDR"])

ws.merge_cells(start_row=phdr_row, start_column=STRIP_COL,
               end_row=phdr_row, end_column=STRIP_COL+11)
hdr(ws, phdr_row, STRIP_COL, "STRIPPING DATA  (HNO₃ system)", span=12,
    bg=C["STR_TINT"], fg=C["SECHDR"])

col_hdr_row = REG_START + 2
ws.row_dimensions[col_hdr_row].height = 34
for i, label in enumerate(PANEL_COLS):
    hdr(ws, col_hdr_row, LOAD_COL+i, label)
    hdr(ws, col_hdr_row, STRIP_COL+i, label)

# Data rows: 3 loading stages, 3 strip stages (S3 empty, only 2 used)
DATA_ROW_START = REG_START + 3  # row 7
for i, stage_lbl in enumerate(STAGE_LABELS_LOAD):
    r = DATA_ROW_START + i
    ws.row_dimensions[r].height = 16
    style_cell(ws, r, LOAD_COL, C["STEP_BG"], C["STEP_FG"], val=stage_lbl)
    for c_off in range(1, 12):
        inp(ws, r, LOAD_COL + c_off)

for i, stage_lbl in enumerate(STAGE_LABELS_STRIP):
    r = DATA_ROW_START + i
    style_cell(ws, r, STRIP_COL, C["STEP_BG"], C["STEP_FG"], val=stage_lbl)
    for c_off in range(1, 12):
        inp(ws, r, STRIP_COL + c_off)

# ── Calculated D values ───────────────────────────────────────────────────────
D_HDR_ROW = DATA_ROW_START + 3  # row 10
ws.row_dimensions[D_HDR_ROW].height = 18
ws.merge_cells(f"A{D_HDR_ROW}:M{D_HDR_ROW}")
hdr(ws, D_HDR_ROW, 1, "CALCULATED D VALUES  (D = C_org / C_aq)", span=13,
    bg=C["COLHDR"])
ws.merge_cells(f"N{D_HDR_ROW}:Z{D_HDR_ROW}")
hdr(ws, D_HDR_ROW, 14, "CALCULATED D VALUES  (D = C_org / C_aq)", span=13,
    bg=C["COLHDR"])

D_COL_HDR = D_HDR_ROW + 1
ws.row_dimensions[D_COL_HDR].height = 16
d_cols = ["Stage", "D_U", "D_Th", "D_Hf", "D_Zr"]
for i, lbl in enumerate(d_cols):
    hdr(ws, D_COL_HDR, LOAD_COL+i, lbl)
    hdr(ws, D_COL_HDR, STRIP_COL+i, lbl)

# D data rows (Load)
D_DATA_START = D_COL_HDR + 1  # row 12
# Column mapping for input table: col offset 4=U_aq,5=Th,6=Hf,7=Zr (aq), 8=U_org...
for i in range(3):
    r = D_DATA_START + i
    ws.row_dimensions[r].height = 16
    inp_r = DATA_ROW_START + i
    style_cell(ws, r, LOAD_COL, C["STEP_BG"], C["STEP_FG"],
               val=STAGE_LABELS_LOAD[i])
    # D for each element
    for e_idx, e in enumerate(ELEMENTS):
        aq_col = LOAD_COL + 4 + e_idx   # U_aq=col5, Th=6, Hf=7, Zr=8
        org_col = LOAD_COL + 8 + e_idx  # U_org=col9, Th=10, Hf=11, Zr=12
        aq_addr = addr(inp_r, aq_col)
        org_addr = addr(inp_r, org_col)
        formula = (f'=IFERROR(IF(OR({aq_addr}="",{org_addr}=""),"—",'
                   f'{org_addr}/{aq_addr}),"—")')
        calc(ws, r, LOAD_COL+1+e_idx, formula)

    style_cell(ws, r, STRIP_COL, C["STEP_BG"], C["STEP_FG"],
               val=STAGE_LABELS_STRIP[i])
    for e_idx, e in enumerate(ELEMENTS):
        aq_col = STRIP_COL + 4 + e_idx
        org_col = STRIP_COL + 8 + e_idx
        aq_addr = addr(inp_r, aq_col)
        org_addr = addr(inp_r, org_col)
        formula = (f'=IFERROR(IF(OR({aq_addr}="",{org_addr}=""),"—",'
                   f'{org_addr}/{aq_addr}),"—")')
        calc(ws, r, STRIP_COL+1+e_idx, formula)

# ── Log-log regression ────────────────────────────────────────────────────────
REG_SUM_ROW = D_DATA_START + 3 + 1  # row 17
ws.row_dimensions[REG_SUM_ROW].height = 18
ws.merge_cells(f"A{REG_SUM_ROW}:Z{REG_SUM_ROW}")
hdr(ws, REG_SUM_ROW, 1,
    "LOG-LOG REGRESSION SUMMARY  (D = P × [acid]^Q  →  feeds TDMA Panel B)",
    span=26, bg=C["SECHDR"])

REG_COL_HDR = REG_SUM_ROW + 1
ws.row_dimensions[REG_COL_HDR].height = 34
reg_cols = ["Element", "P_loading", "Q_loading", "R²_loading",
            "P_strip", "Q_strip", "R²_strip"]
for i, lbl in enumerate(reg_cols):
    hdr(ws, REG_COL_HDR, LOAD_COL+i, lbl)

# Regression data rows — one per element
REG_DATA_START = REG_COL_HDR + 1  # row 19
# D value cell addresses (loading): rows D_DATA_START to D_DATA_START+2, cols 2-5
# acid N: DATA_ROW_START to +2, col 2 (LOAD_COL+1)

# Build ranges
load_acid_range = (f"{addr(DATA_ROW_START, LOAD_COL+1)}:"
                   f"{addr(DATA_ROW_START+2, LOAD_COL+1)}")
strip_acid_range = (f"{addr(DATA_ROW_START, STRIP_COL+1)}:"
                    f"{addr(DATA_ROW_START+2, STRIP_COL+1)}")

# We store regression cell addresses for later linking to TDMA Panel B
REG_CELLS = {}  # element -> {P_load, Q_load, P_strip, Q_strip}

for e_idx, elem in enumerate(ELEMENTS):
    r = REG_DATA_START + e_idx
    ws.row_dimensions[r].height = 16
    style_cell(ws, r, 1, C["STEP_BG"], C["STEP_FG"], val=elem)

    load_d_col = LOAD_COL + 1 + e_idx   # D_U is col2, etc.
    strip_d_col = STRIP_COL + 1 + e_idx

    load_d_range = (f"{addr(D_DATA_START, load_d_col)}:"
                    f"{addr(D_DATA_START+2, load_d_col)}")
    strip_d_range = (f"{addr(D_DATA_START, strip_d_col)}:"
                     f"{addr(D_DATA_START+2, strip_d_col)}")

    # LN arrays for SLOPE/INTERCEPT: We need numeric only.
    # Since D might be "—", use IFERROR with VALUE... but in Excel arrays
    # we rely on IF checks. We'll use helper approach with IFERROR on SLOPE.
    def reg_formula_Q(d_range, acid_range):
        return (f'=IFERROR(IF(COUNTA({d_range})<2,"Need ≥2 pts",'
                f'SLOPE(LN(IF({d_range}<>"—",IF(ISNUMBER({d_range}),{d_range},1),1)),'
                f'LN(IF({acid_range}<>"",{acid_range},1)))),"N/A")')

    def reg_formula_P(d_range, acid_range):
        return (f'=IFERROR(IF(COUNTA({d_range})<2,"Need ≥2 pts",'
                f'EXP(INTERCEPT(LN(IF({d_range}<>"—",IF(ISNUMBER({d_range}),{d_range},1),1)),'
                f'LN(IF({acid_range}<>"",{acid_range},1))))),"N/A")')

    def reg_formula_R2(d_range, acid_range):
        return (f'=IFERROR(IF(COUNTA({d_range})<2,"Need ≥2 pts",'
                f'RSQ(LN(IF({d_range}<>"—",IF(ISNUMBER({d_range}),{d_range},1),1)),'
                f'LN(IF({acid_range}<>"",{acid_range},1)))),"N/A")')

    # P_loading (col 2)
    P_load_cell = out(ws, r, 2)
    P_load_cell.value = reg_formula_P(load_d_range, load_acid_range)

    # Q_loading (col 3)
    Q_load_cell = out(ws, r, 3)
    Q_load_cell.value = reg_formula_Q(load_d_range, load_acid_range)

    # R2_loading (col 4)
    R2_load = calc(ws, r, 4)
    R2_load.value = reg_formula_R2(load_d_range, load_acid_range)

    # P_strip (col 5)
    P_strip_cell = out(ws, r, 5)
    P_strip_cell.value = reg_formula_P(strip_d_range, strip_acid_range)

    # Q_strip (col 6)
    Q_strip_cell = out(ws, r, 6)
    Q_strip_cell.value = reg_formula_Q(strip_d_range, strip_acid_range)

    # R2_strip (col 7)
    R2_strip = calc(ws, r, 7)
    R2_strip.value = reg_formula_R2(strip_d_range, strip_acid_range)

    REG_CELLS[elem] = {
        "P_load": (r, 2),
        "Q_load": (r, 3),
        "P_strip": (r, 5),
        "Q_strip": (r, 6),
    }

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TDMA INPUT PANELS  (rows 57–68)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
TDMA_START = 57
ws.row_dimensions[TDMA_START].height = 20
ws.merge_cells(f"A{TDMA_START}:AZ{TDMA_START}")
c = ws.cell(row=TDMA_START, column=1,
            value="TDMA SOLVER — INPUT PANELS")
c.fill = fill(C["TITLE"])
c.font = font("FFFFFF", bold=True, sz=12)
c.alignment = center()

# Panel A — General Inputs (cols A-B = 1-2)
PA_ROW = TDMA_START + 1
ws.row_dimensions[PA_ROW].height = 16
ws.merge_cells(f"A{PA_ROW}:B{PA_ROW}")
hdr(ws, PA_ROW, 1, "PANEL A — General Inputs", span=2, bg=C["SECHDR"])

panel_a = [
    ("n_ext (extraction stages)", 3),
    ("n_scr (scrub stages)", 0),
    ("n_str (strip stages)", 2),
    ("[H₂SO₄] Extraction (N)", 0.5),
    ("[H₂SO₄] Scrub (N)", 1.0),
    ("[HNO₃] Strip (N)", 4.0),
    ("F — Feed flow aq (L/min)", 1.0),
    ("S — Strip flow aq (L/min)", 1.0),
    ("O — Org flow (L/min)", 1.0),
    ("R — Recycle ratio", 0.0),
]

# Absolute cell references for Panel A inputs
PA_DATA_START = PA_ROW + 1
PA_CELLS = {}  # key -> (row, col=2)
for i, (label, default) in enumerate(panel_a):
    r = PA_DATA_START + i
    ws.row_dimensions[r].height = 16
    style_cell(ws, r, 1, C["REF_BG"], C["REF_FG"], val=label, align="left")
    inp(ws, r, 2, default)
    PA_CELLS[label] = (r, 2)

# Shorthand absolute refs
def pa(key):
    r, c_ = PA_CELLS[key]
    return cr(r, c_)

N_EXT_REF = pa("n_ext (extraction stages)")
N_SCR_REF = pa("n_scr (scrub stages)")
N_STR_REF = pa("n_str (strip stages)")
ACID_EXT_REF = pa("[H₂SO₄] Extraction (N)")
ACID_SCR_REF = pa("[H₂SO₄] Scrub (N)")
ACID_STR_REF = pa("[HNO₃] Strip (N)")
F_REF = pa("F — Feed flow aq (L/min)")
S_REF = pa("S — Strip flow aq (L/min)")
O_REF = pa("O — Org flow (L/min)")
R_REF = pa("R — Recycle ratio")

# Panel B — Distribution Coefficients (cols F-K = 6-11)
PB_COL = 6
PB_ROW = TDMA_START + 1
ws.merge_cells(start_row=PB_ROW, start_column=PB_COL,
               end_row=PB_ROW, end_column=PB_COL+5)
hdr(ws, PB_ROW, PB_COL,
    "PANEL B — Distribution Coefficients (linked from Regression)",
    span=6, bg=C["SECHDR"])

pb_col_hdr = PB_ROW + 1
pb_labels = ["Element", "P_load", "Q_load", "P_strip", "Q_strip", "Note"]
for i, lbl in enumerate(pb_labels):
    hdr(ws, pb_col_hdr, PB_COL+i, lbl)

# Defaults in case regression returns N/A
DEFAULTS = {
    "U":  {"P_load": 0.80, "Q_load": -1.2, "P_strip": 0.80, "Q_strip": -1.2},
    "Th": {"P_load": 1.50, "Q_load": -1.5, "P_strip": 1.50, "Q_strip": -1.5},
    "Hf": {"P_load": 0.30, "Q_load": -0.8, "P_strip": 0.30, "Q_strip": -0.8},
    "Zr": {"P_load": 2.00, "Q_load": -1.0, "P_strip": 2.00, "Q_strip": -1.0},
}

PB_CELLS = {}  # element -> {P_load, Q_load, P_strip, Q_strip} absolute refs
pb_data_start = pb_col_hdr + 1
for e_idx, elem in enumerate(ELEMENTS):
    r = pb_data_start + e_idx
    ws.row_dimensions[r].height = 16
    style_cell(ws, r, PB_COL, C["STEP_BG"], C["STEP_FG"], val=elem)

    reg = REG_CELLS[elem]
    defs = DEFAULTS[elem]

    def linked(reg_r, reg_c, default_val):
        ref = cr(reg_r, reg_c)
        return (f'=IFERROR(IF(ISNUMBER({ref}),{ref},{default_val}),{default_val})')

    P_load_c = out(ws, r, PB_COL+1)
    P_load_c.value = linked(*reg["P_load"], defs["P_load"])

    Q_load_c = out(ws, r, PB_COL+2)
    Q_load_c.value = linked(*reg["Q_load"], defs["Q_load"])

    P_strip_c = out(ws, r, PB_COL+3)
    P_strip_c.value = linked(*reg["P_strip"], defs["P_strip"])

    Q_strip_c = out(ws, r, PB_COL+4)
    Q_strip_c.value = linked(*reg["Q_strip"], defs["Q_strip"])

    note_c = ws.cell(row=r, column=PB_COL+5,
                     value="Live-linked from regression")
    note_c.fill = fill(C["CALC_BG"])
    note_c.font = font(C["CALC_FG"], sz=8)
    note_c.alignment = left()

    PB_CELLS[elem] = {
        "P_load":  cr(r, PB_COL+1),
        "Q_load":  cr(r, PB_COL+2),
        "P_strip": cr(r, PB_COL+3),
        "Q_strip": cr(r, PB_COL+4),
    }

# Panel C — Feed Concentrations (cols M-O = 13-15)
PC_COL = 13
PC_ROW = TDMA_START + 1
ws.merge_cells(start_row=PC_ROW, start_column=PC_COL,
               end_row=PC_ROW, end_column=PC_COL+2)
hdr(ws, PC_ROW, PC_COL, "PANEL C — Feed Conc. (g/L)", span=3, bg=C["SECHDR"])
hdr(ws, PC_ROW+1, PC_COL, "Element")
hdr(ws, PC_ROW+1, PC_COL+1, "X_feed (g/L)")
hdr(ws, PC_ROW+1, PC_COL+2, "MW (g/mol)")

FEED_DEFAULTS = {"U": 5.0, "Th": 2.0, "Hf": 0.5, "Zr": 1.0}
MW = {"U": 238.03, "Th": 232.04, "Hf": 178.49, "Zr": 91.22}
PC_CELLS = {}
pc_data_start = PC_ROW + 2
for e_idx, elem in enumerate(ELEMENTS):
    r = pc_data_start + e_idx
    ws.row_dimensions[r].height = 16
    style_cell(ws, r, PC_COL, C["STEP_BG"], C["STEP_FG"], val=elem)
    inp(ws, r, PC_COL+1, FEED_DEFAULTS[elem])
    style_cell(ws, r, PC_COL+2, C["CALC_BG"], C["CALC_FG"], val=MW[elem])
    PC_CELLS[elem] = {"XF": cr(r, PC_COL+1)}

# Panel D — Results Summary (cols Q-R = 17-18)
PD_COL = 17
PD_ROW = TDMA_START + 1

# Panel E — O/A Ratio Tracker (cols T-V = 20-22)
PE_COL = 20
PE_ROW = TDMA_START + 1

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TDMA SOLVER  (rows 70+)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SOLVER_START = 80
ws.row_dimensions[SOLVER_START].height = 20
ws.merge_cells(f"A{SOLVER_START}:BZ{SOLVER_START}")
c = ws.cell(row=SOLVER_START, column=1, value="TDMA SOLVER — 25 STAGES")
c.fill = fill(C["TITLE"])
c.font = font("FFFFFF", bold=True, sz=12)
c.alignment = center()

# Column layout:
# Per element: D, a, b, c, d, cp(c'), dp(d'), y_org  = 8 cols each
# After 4 elements: x_aq per element = 4 cols
# Then: total_org, total_aq
COL_STAGE = 1
ELEM_BASE = 2  # starts at col 2
COLS_PER_ELEM = 8
X_AQ_BASE = ELEM_BASE + 4 * COLS_PER_ELEM  # col 34
TOTAL_ORG_COL = X_AQ_BASE + 4   # col 38
TOTAL_AQ_COL = TOTAL_ORG_COL + 1  # col 39

ELEM_COLS = {}
for e_idx, elem in enumerate(ELEMENTS):
    base = ELEM_BASE + e_idx * COLS_PER_ELEM
    ELEM_COLS[elem] = {
        "D": base,
        "a": base+1,
        "b": base+2,
        "c": base+3,
        "d": base+4,
        "cp": base+5,
        "dp": base+6,
        "y_org": base+7,
    }
    x_col = X_AQ_BASE + e_idx
    ELEM_COLS[elem]["x_aq"] = x_col

# Header rows
HDR1 = SOLVER_START + 1
HDR2 = SOLVER_START + 2
ws.row_dimensions[HDR1].height = 16
ws.row_dimensions[HDR2].height = 16

# Stage col
hdr(ws, HDR1, COL_STAGE, "Stage", bg=C["SECHDR"])
ws.merge_cells(start_row=HDR1, start_column=COL_STAGE,
               end_row=HDR2, end_column=COL_STAGE)

for elem in ELEMENTS:
    ec = ELEM_COLS[elem]
    base = ec["D"]
    ws.merge_cells(start_row=HDR1, start_column=base,
                   end_row=HDR1, end_column=base+7)
    hdr(ws, HDR1, base, f"{elem}  (8-column TDMA block)", span=8,
        bg=C["COLHDR"])
    sub_labels = ["D_n", "a_n", "b_n", "c_n", "d_n", "c'_n", "d'_n", "y_org"]
    for i, lbl in enumerate(sub_labels):
        hdr(ws, HDR2, base+i, lbl)

    x_col = ec["x_aq"]
    ws.merge_cells(start_row=HDR1, start_column=x_col,
                   end_row=HDR2, end_column=x_col)
    hdr(ws, HDR1, x_col, f"x_aq\n{elem}", bg=C["OUTPUT_BG"], fg=C["OUTPUT_FG"])

ws.merge_cells(start_row=HDR1, start_column=TOTAL_ORG_COL,
               end_row=HDR2, end_column=TOTAL_ORG_COL)
hdr(ws, HDR1, TOTAL_ORG_COL, "Total\nOrg", bg=C["COLHDR"])
ws.merge_cells(start_row=HDR1, start_column=TOTAL_AQ_COL,
               end_row=HDR2, end_column=TOTAL_AQ_COL)
hdr(ws, HDR1, TOTAL_AQ_COL, "Total\nAq", bg=C["COLHDR"])

# Stage data rows
STAGE_ROW_START = HDR2 + 1  # row 73
for n in range(1, N_STAGES+1):
    r = STAGE_ROW_START + n - 1
    ws.row_dimensions[r].height = 16

    # Stage label with section color
    # Stage section: extraction if n <= n_ext, scrub if n <= n_ext+n_scr, else strip
    # We use conditional references; for label just put n
    lbl_cell = ws.cell(row=r, column=COL_STAGE, value=n)
    lbl_cell.font = font(C["STEP_FG"], bold=True)
    lbl_cell.alignment = center()
    lbl_cell.border = thin_border()
    # Tint based on position (static since section sizes are user-defined;
    # we tint by default config: ext=1-3, scr=none, strip=4-5)
    if n <= 3:
        lbl_cell.fill = fill(C["EXT_TINT"])
    elif n <= 5:
        lbl_cell.fill = fill(C["STR_TINT"])
    else:
        lbl_cell.fill = fill(C["CALC_BG"])

    for elem in ELEMENTS:
        ec = ELEM_COLS[elem]
        pb = PB_CELLS[elem]

        # Cell addresses for this row
        D_cell_addr = addr(r, ec["D"])
        a_addr = addr(r, ec["a"])
        b_addr = addr(r, ec["b"])
        c_addr = addr(r, ec["c"])
        d_addr = addr(r, ec["d"])
        cp_addr = addr(r, ec["cp"])
        dp_addr = addr(r, ec["dp"])
        y_addr = addr(r, ec["y_org"])

        # ── D_n ──
        # IF strip stage: use P_strip, Q_strip, [HNO3]; else P_load, Q_load, [H2SO4]
        # Strip stage: n > n_ext + n_scr
        # n_ext + n_scr is the last extraction/scrub stage
        d_formula = (
            f'=IF({n}>{N_EXT_REF}+{N_SCR_REF},'
            f'{pb["P_strip"]}*{ACID_STR_REF}^{pb["Q_strip"]},'
            f'{pb["P_load"]}*{ACID_EXT_REF}^{pb["Q_load"]})'
        )
        step(ws, r, ec["D"], d_formula)

        # ── a_n = O; a_1 = 0 ──
        a_formula = f'=IF({n}=1,0,{O_REF})'
        step(ws, r, ec["a"], a_formula)

        # ── b_n = -(O + Q_aq/D_n) ──
        # Q_aq:
        #   extraction: Q_aq = S*R + F   (=S*R+F)
        #   scrub:      Q_aq = S*R
        #   strip:      Q_aq = S
        # SR = S*R
        Q_aq = (
            f'IF({n}<={N_EXT_REF},{S_REF}*{R_REF}+{F_REF},'
            f'IF({n}<={N_EXT_REF}+{N_SCR_REF},{S_REF}*{R_REF},{S_REF}))'
        )
        b_formula = f'=-({O_REF}+({Q_aq})/{D_cell_addr})'
        step(ws, r, ec["b"], b_formula)

        # ── c_n ──
        # n > n_ext+n_scr AND n != N  →  S/D_{n+1}
        # n = N (=25)                 →  0  (boundary)
        # n = n_ext+n_scr (last scrub) →  0  (boundary)
        # scrub (n_ext <= n < n_ext+n_scr) → SR/D_{n+1}
        # extraction (n < n_ext)           → (SR+F)/D_{n+1}
        if n < N_STAGES:
            D_next_addr = addr(r+1, ec["D"])
            c_formula = (
                f'=IF({n}={N_EXT_REF}+{N_SCR_REF},0,'
                f'IF({n}>{N_EXT_REF}+{N_SCR_REF},{S_REF}/{D_next_addr},'
                f'IF({n}>={N_EXT_REF},{S_REF}*{R_REF}/{D_next_addr},'
                f'({S_REF}*{R_REF}+{F_REF})/{D_next_addr})))'
            )
        else:
            c_formula = '=0'
        step(ws, r, ec["c"], c_formula)

        # ── d_n = -XF*F at n=n_ext; else 0 ──
        XF = PC_CELLS[elem]["XF"]
        d_formula_coeff = (
            f'=IF({n}={N_EXT_REF},-{XF}*{F_REF},0)'
        )
        step(ws, r, ec["d"], d_formula_coeff)

        # ── Forward sweep ──
        # Row 1 (n=1):
        #   c'_1 = c_1 / b_1
        #   d'_1 = d_1 / b_1
        # Row n>1:
        #   c'_n = c_n / (b_n - a_n * c'_{n-1})
        #   d'_n = (d_n - a_n * d'_{n-1}) / (b_n - a_n * c'_{n-1})
        if n == 1:
            cp_formula = f'={c_addr}/{b_addr}'
            dp_formula = f'={d_addr}/{b_addr}'
        else:
            prev_cp = addr(r-1, ec["cp"])
            prev_dp = addr(r-1, ec["dp"])
            denom = f'({b_addr}-{a_addr}*{prev_cp})'
            cp_formula = f'={c_addr}/{denom}'
            dp_formula = f'=({d_addr}-{a_addr}*{prev_dp})/{denom}'
        step(ws, r, ec["cp"], cp_formula)
        step(ws, r, ec["dp"], dp_formula)

# ── Back substitution (y_org) — must do in reverse order ──
for n in range(N_STAGES, 0, -1):
    r = STAGE_ROW_START + n - 1
    for elem in ELEMENTS:
        ec = ELEM_COLS[elem]
        cp_addr = addr(r, ec["cp"])
        dp_addr = addr(r, ec["dp"])
        y_addr = addr(r, ec["y_org"])

        if n == N_STAGES:
            y_formula = f'={dp_addr}'
        else:
            y_next = addr(r+1, ec["y_org"])
            y_formula = f'={dp_addr}-{cp_addr}*{y_next}'
        out(ws, r, ec["y_org"], y_formula)

# ── x_aq and totals ──
for n in range(1, N_STAGES+1):
    r = STAGE_ROW_START + n - 1
    total_org = []
    total_aq = []
    for elem in ELEMENTS:
        ec = ELEM_COLS[elem]
        y_addr = addr(r, ec["y_org"])
        D_addr = addr(r, ec["D"])
        x_formula = f'=IFERROR({y_addr}/{D_addr},0)'
        calc(ws, r, ec["x_aq"], x_formula)
        total_org.append(addr(r, ec["y_org"]))
        total_aq.append(addr(r, ec["x_aq"]))

    calc(ws, r, TOTAL_ORG_COL,
         "="+"+".join(total_org))
    calc(ws, r, TOTAL_AQ_COL,
         "="+"+".join(total_aq))

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# PANEL D — Results Summary
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ws.merge_cells(start_row=PD_ROW, start_column=PD_COL,
               end_row=PD_ROW, end_column=PD_COL+1)
hdr(ws, PD_ROW, PD_COL, "PANEL D — Results Summary", span=2, bg=C["SECHDR"])
hdr(ws, PD_ROW+1, PD_COL, "Metric")
hdr(ws, PD_ROW+1, PD_COL+1, "Value (g/L)")

pd_data_start = PD_ROW + 2
result_rows = []
for e_idx, elem in enumerate(ELEMENTS):
    ec = ELEM_COLS[elem]
    # Raffinate = x_aq at stage 1
    r_raff = pd_data_start + e_idx * 3
    ws.row_dimensions[r_raff].height = 16
    style_cell(ws, r_raff, PD_COL, C["REF_BG"], C["REF_FG"],
               val=f"Raffinate {elem} (g/L)")
    raff_addr = addr(STAGE_ROW_START, ec["x_aq"])
    out(ws, r_raff, PD_COL+1, f"={raff_addr}")

    # Loaded org at last extraction stage
    r_load = r_raff + 1
    ws.row_dimensions[r_load].height = 16
    style_cell(ws, r_load, PD_COL, C["REF_BG"], C["REF_FG"],
               val=f"Loaded Org {elem} (g/L)")
    # INDEX on y_org column at row n_ext
    y_col_letter = get_column_letter(ec["y_org"])
    y_range = f"{y_col_letter}{STAGE_ROW_START}:{y_col_letter}{STAGE_ROW_START+N_STAGES-1}"
    out(ws, r_load, PD_COL+1,
        f'=IFERROR(INDEX({y_range},{N_EXT_REF}),0)')

    # Pregnant solution = x_aq at strip exit stage
    r_preg = r_raff + 2
    ws.row_dimensions[r_preg].height = 16
    style_cell(ws, r_preg, PD_COL, C["REF_BG"], C["REF_FG"],
               val=f"Pregnant {elem} (g/L)")
    x_col_letter = get_column_letter(ec["x_aq"])
    x_range = f"{x_col_letter}{STAGE_ROW_START}:{x_col_letter}{STAGE_ROW_START+N_STAGES-1}"
    # Preg = x_aq at stage n_ext+n_scr+n_str (last strip stage)
    out(ws, r_preg, PD_COL+1,
        f'=IFERROR(INDEX({x_range},{N_EXT_REF}+{N_SCR_REF}+{N_STR_REF}),0)')

# Panel E — O/A tracker
ws.merge_cells(start_row=PE_ROW, start_column=PE_COL,
               end_row=PE_ROW, end_column=PE_COL+2)
hdr(ws, PE_ROW, PE_COL, "PANEL E — O/A Ratio Tracker", span=3, bg=C["SECHDR"])
pe_labels = ["Section", "O/A (org/aq)", "Note"]
for i, lbl in enumerate(pe_labels):
    hdr(ws, PE_ROW+1, PE_COL+i, lbl)

pe_data = [
    ("Extraction", f"={O_REF}/{F_REF}", "O/F"),
    ("Scrub",      f"={O_REF}/({S_REF}*{R_REF}+0.001)", "O/(S*R)"),
    ("Strip",      f"={O_REF}/{S_REF}", "O/S"),
]
for i, (sec, formula, note) in enumerate(pe_data):
    r = PE_ROW + 2 + i
    ws.row_dimensions[r].height = 16
    style_cell(ws, r, PE_COL, C["STEP_BG"], C["STEP_FG"], val=sec)
    out(ws, r, PE_COL+1, formula)
    style_cell(ws, r, PE_COL+2, C["CALC_BG"], C["CALC_FG"], val=note)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# STAGE PROFILE TABLE (for charts) — after solver
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PROF_START = STAGE_ROW_START + N_STAGES + 2
ws.row_dimensions[PROF_START].height = 18
ws.merge_cells(f"A{PROF_START}:J{PROF_START}")
hdr(ws, PROF_START, 1, "STAGE AQUEOUS CONCENTRATION PROFILES  (g/L)",
    span=10, bg=C["SECHDR"])

PROF_HDR = PROF_START + 1
ws.row_dimensions[PROF_HDR].height = 16
hdr(ws, PROF_HDR, 1, "Stage")
for e_idx, elem in enumerate(ELEMENTS):
    hdr(ws, PROF_HDR, 2+e_idx, f"x_aq {elem}")

for n in range(1, N_STAGES+1):
    r = PROF_HDR + n
    ws.row_dimensions[r].height = 14
    style_cell(ws, r, 1, C["STEP_BG"], C["STEP_FG"], val=n)
    for e_idx, elem in enumerate(ELEMENTS):
        ec = ELEM_COLS[elem]
        solver_row = STAGE_ROW_START + n - 1
        ref = addr(solver_row, ec["x_aq"])
        calc(ws, r, 2+e_idx, f"={ref}")

# ── Line charts for aqueous profiles ──
for e_idx, elem in enumerate(ELEMENTS):
    from openpyxl.chart import LineChart
    chart = LineChart()
    chart.title = f"{elem} — Aqueous Profile vs Stage"
    chart.style = 10
    chart.y_axis.title = "x_aq (g/L)"
    chart.x_axis.title = "Stage"
    chart.height = 10
    chart.width = 18

    stage_col = 1
    data_col = 2 + e_idx
    data_ref = Reference(ws,
                         min_col=data_col,
                         min_row=PROF_HDR,
                         max_row=PROF_HDR+N_STAGES)
    cats = Reference(ws,
                     min_col=stage_col,
                     min_row=PROF_HDR+1,
                     max_row=PROF_HDR+N_STAGES)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    chart.series[0].graphicalProperties.line.solidFill = "1B3A6B"

    anchor_row = PROF_HDR + N_STAGES + 2 + e_idx * 22
    anchor_col = get_column_letter(1)
    ws.add_chart(chart, f"{anchor_col}{anchor_row}")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ORGANIC LOADING SECTION
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ORG_START = PROF_HDR + N_STAGES + 100
ws.row_dimensions[ORG_START].height = 20
ws.merge_cells(f"A{ORG_START}:N{ORG_START}")
c = ws.cell(row=ORG_START, column=1,
            value="ORGANIC LOADING SECTION — Primene JMT Amine Capacity")
c.fill = fill(C["SECHDR"])
c.font = font("FFFFFF", bold=True, sz=12)
c.alignment = center()

org_rows = [
    ("Extractant", "Primene JMT"),
    ("MW Extractant (g/mol)", 353.7),
    ("Density extractant (g/mL)", 0.812),
    ("Vol% in diluent", 10.0),
    ("Diluent", "Orfom"),
    ("Density diluent (g/mL)", 0.785),
    ("Note on saponification", "N/A — primary amine (not quaternary)"),
    ("Experimental loading (%)", 80.0),
]
org_data_start = ORG_START + 1
for i, (label, val) in enumerate(org_rows):
    r = org_data_start + i
    ws.row_dimensions[r].height = 16
    style_cell(ws, r, 1, C["REF_BG"], C["REF_FG"], val=label, align="left")
    if isinstance(val, str):
        style_cell(ws, r, 2, C["CALC_BG"], C["CALC_FG"], val=val, align="left")
    else:
        inp(ws, r, 2, val)

# vol% row index
VOL_R = org_data_start + 3  # 10.0
MW_EXT_R = org_data_start + 1
DENS_EXT_R = org_data_start + 2
EXP_LOAD_R = org_data_start + 7

calc_rows = [
    ("Amine conc (mol/L)",
     f"=({cr(VOL_R,2)}/100)*{cr(DENS_EXT_R,2)}*1000/{cr(MW_EXT_R,2)}"),
    ("Avg MW metals (g/mol)",
     f"=AVERAGE({cr(pc_data_start,PC_COL+2)}:{cr(pc_data_start+3,PC_COL+2)})"),
]
calc_r = org_data_start + len(org_rows)
for label, formula in calc_rows:
    ws.row_dimensions[calc_r].height = 16
    style_cell(ws, calc_r, 1, C["REF_BG"], C["REF_FG"], val=label, align="left")
    calc(ws, calc_r, 2, formula)
    calc_r += 1

amine_conc_r = org_data_start + len(org_rows)
avg_mw_r = amine_conc_r + 1

ws.row_dimensions[calc_r].height = 16
style_cell(ws, calc_r, 1, C["REF_BG"], C["REF_FG"],
           val="Max loading (g metal/L org)", align="left")
calc(ws, calc_r, 2,
     f"={cr(amine_conc_r,2)}*{cr(avg_mw_r,2)}")
calc_r += 1

ws.row_dimensions[calc_r].height = 16
style_cell(ws, calc_r, 1, C["REF_BG"], C["REF_FG"],
           val="Actual loading (g/L)", align="left")
out(ws, calc_r, 2,
    f"={cr(calc_r-1,2)}*{cr(EXP_LOAD_R,2)}/100")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# RESIDENCE TIME
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
RT_START = calc_r + 3
ws.row_dimensions[RT_START].height = 20
ws.merge_cells(f"A{RT_START}:N{RT_START}")
c = ws.cell(row=RT_START, column=1, value="RESIDENCE TIME ANALYSIS")
c.fill = fill(C["SECHDR"])
c.font = font("FFFFFF", bold=True, sz=12)
c.alignment = center()

rt_inp_start = RT_START + 1
rt_inputs = [
    ("Mixer volume (m³)", 0.001),
    ("Settler length (m)", 1.0),
    ("Settler width (m)", 0.5),
    ("Settler depth (m)", 0.2),
    ("n settlers/stage", 1),
]
RT_CELLS = {}
for i, (label, val) in enumerate(rt_inputs):
    r = rt_inp_start + i
    ws.row_dimensions[r].height = 16
    style_cell(ws, r, 1, C["REF_BG"], C["REF_FG"], val=label, align="left")
    inp(ws, r, 2, val)
    RT_CELLS[label] = cr(r, 2)

MIXER_V = RT_CELLS["Mixer volume (m³)"]
SETT_L = RT_CELLS["Settler length (m)"]
SETT_W = RT_CELLS["Settler width (m)"]
SETT_D = RT_CELLS["Settler depth (m)"]
N_SETT = RT_CELLS["n settlers/stage"]

rt_hdr_row = rt_inp_start + len(rt_inputs) + 1
ws.row_dimensions[rt_hdr_row].height = 34
rt_cols = ["Section", "n stages", "Q_aq\n(mL/min)", "Q_tot\n(mL/min)",
           "τ_mixer\n(min)", "V_settler\n(mL)", "τ_settler\n(min)",
           "τ_stage\n(min)", "τ_section\n(min)"]
for i, lbl in enumerate(rt_cols):
    hdr(ws, rt_hdr_row, 1+i, lbl)

rt_sections = [
    ("Extraction", N_EXT_REF,
     f"={F_REF}*1000",
     f"=({F_REF}+{O_REF})*1000"),
    ("Scrub", N_SCR_REF,
     f"={S_REF}*{R_REF}*1000",
     f"=({S_REF}*{R_REF}+{O_REF})*1000"),
    ("Strip", N_STR_REF,
     f"={S_REF}*1000",
     f"=({S_REF}+{O_REF})*1000"),
]
rt_data_start = rt_hdr_row + 1
for i, (sec, n_stages, q_aq, q_tot) in enumerate(rt_sections):
    r = rt_data_start + i
    ws.row_dimensions[r].height = 16
    style_cell(ws, r, 1, C["STEP_BG"], C["STEP_FG"], val=sec)
    ref_cell(ws, r, 2, f"={n_stages}")
    calc(ws, r, 3, q_aq)
    calc(ws, r, 4, q_tot)
    # τ_mixer = V_mixer(mL) / Q_tot  (V_mixer in m3 → *1e6 to mL)
    calc(ws, r, 5, f"=({MIXER_V}*1000000)/{addr(r,4)}")
    # V_settler = L*W*D * n_sett * 1e6 (m3→mL)
    calc(ws, r, 6, f"={SETT_L}*{SETT_W}*{SETT_D}*{N_SETT}*1000000")
    # τ_settler
    calc(ws, r, 7, f"={addr(r,6)}/{addr(r,4)}")
    # τ_stage = τ_mixer + τ_settler
    calc(ws, r, 8, f"={addr(r,5)}+{addr(r,7)}")
    # τ_section = n_stages * τ_stage
    calc(ws, r, 9, f"={addr(r,2)}*{addr(r,8)}")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# STEADY-STATE TIME ESTIMATOR
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SS_START = rt_data_start + len(rt_sections) + 2
ws.row_dimensions[SS_START].height = 18
ws.merge_cells(f"A{SS_START}:N{SS_START}")
hdr(ws, SS_START, 1, "STEADY-STATE TIME ESTIMATOR  (Ritcey & Ashbrook 1979; Perry's 8th Ed. §15-42)",
    span=14, bg=C["SECHDR"])

ss_data = SS_START + 1
ws.row_dimensions[ss_data].height = 16
style_cell(ws, ss_data, 1, C["REF_BG"], C["REF_FG"], val="Safety factor", align="left")
inp(ws, ss_data, 2, 3)
SF_REF = cr(ss_data, 2)

ss_data += 1
ws.row_dimensions[ss_data].height = 16
style_cell(ws, ss_data, 1, C["REF_BG"], C["REF_FG"],
           val="τ_max (bottleneck, min)", align="left")
tau_range = (f"{addr(rt_data_start,9)}:{addr(rt_data_start+2,9)}")
calc(ws, ss_data, 2, f"=MAX({tau_range})")
TAU_MAX = cr(ss_data, 2)

ss_data += 1
ws.row_dimensions[ss_data].height = 16
style_cell(ws, ss_data, 1, C["REF_BG"], C["REF_FG"],
           val="t_SS = factor × τ_max (min)", align="left")
out(ws, ss_data, 2, f"={SF_REF}*{TAU_MAX}")

ss_data += 1
ws.row_dimensions[ss_data].height = 16
style_cell(ws, ss_data, 1, C["REF_BG"], C["REF_FG"],
           val="Bottleneck section", align="left")
sec_range_col = get_column_letter(1)
sec_range = f"{sec_range_col}{rt_data_start}:{sec_range_col}{rt_data_start+2}"
tau_col_range = f"{get_column_letter(9)}{rt_data_start}:{get_column_letter(9)}{rt_data_start+2}"
calc(ws, ss_data, 2,
     f'=IFERROR(INDEX({sec_range},MATCH(MAX({tau_col_range}),{tau_col_range},0)),"—")')

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# FEED VARIABILITY SCENARIO
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
FV_START = ss_data + 3
ws.row_dimensions[FV_START].height = 18
ws.merge_cells(f"A{FV_START}:N{FV_START}")
hdr(ws, FV_START, 1, "FEED VARIABILITY SCENARIO  (Zr feed multipliers)",
    span=14, bg=C["SECHDR"])

fv_inp = FV_START + 1
style_cell(ws, fv_inp, 1, C["REF_BG"], C["REF_FG"], val="k1 (Zr × high)", align="left")
inp(ws, fv_inp, 2, 1.20)
K1 = cr(fv_inp, 2)

fv_inp += 1
style_cell(ws, fv_inp, 1, C["REF_BG"], C["REF_FG"], val="k2 (Zr × low)", align="left")
inp(ws, fv_inp, 2, 0.80)
K2 = cr(fv_inp, 2)

fv_hdr = fv_inp + 2
ws.row_dimensions[fv_hdr].height = 34
fv_col_labels = ["Metric", "Normal", f"Zr×k1", f"Zr×k2"]
for i, lbl in enumerate(fv_col_labels):
    hdr(ws, fv_hdr, 1+i, lbl)

# Zr feed ref
ZR_FEED_REF = PC_CELLS["Zr"]["XF"]

# We reference the regression-linked P/Q for Zr indirectly through the TDMA results.
# For the scenario table we just show the key results and note that
# re-running with modified feed requires Solver/VBA.
# Instead, we add a simple analytical approximation row.
fv_metrics = [
    ("Note", "Scenario requires re-solve; values below use stored TDMA output", "", ""),
    ("Loaded Org Zr (g/L)", "See Panel D", "Zr feed × k1", "Zr feed × k2"),
    ("Raff U (g/L)", "See Panel D row U", "—", "—"),
    ("Preg Zr (g/L)", "See Panel D row Zr", "—", "—"),
]
fv_data = fv_hdr + 1
for label, n_val, k1_val, k2_val in fv_metrics:
    ws.row_dimensions[fv_data].height = 16
    style_cell(ws, fv_data, 1, C["REF_BG"], C["REF_FG"], val=label, align="left")
    style_cell(ws, fv_data, 2, C["CALC_BG"], C["CALC_FG"], val=n_val, align="left")
    style_cell(ws, fv_data, 3, C["CALC_BG"], C["CALC_FG"], val=k1_val, align="left")
    style_cell(ws, fv_data, 4, C["CALC_BG"], C["CALC_FG"], val=k2_val, align="left")
    fv_data += 1

# Effective Zr feed row
ws.row_dimensions[fv_data].height = 16
style_cell(ws, fv_data, 1, C["REF_BG"], C["REF_FG"],
           val="Effective Zr feed (g/L)", align="left")
out(ws, fv_data, 2, f"={ZR_FEED_REF}")
out(ws, fv_data, 3, f"={ZR_FEED_REF}*{K1}")
out(ws, fv_data, 4, f"={ZR_FEED_REF}*{K2}")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Save
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
OUT_FILE = "SX_Steady_State_Model_UThHfZr.xlsx"
wb.save(OUT_FILE)
print(f"Saved: {OUT_FILE}")
