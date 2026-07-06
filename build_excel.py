"""
build_excel.py  —  SX_Steady_State_Model_26elem.xlsx
Primene JMT / dual-acid SX  |  H2SO4 loading, HNO3 stripping
26 elements: Al Sc Fe Co Zn Ga Rb Y Zr La Ce Pr Nd Sm Eu Gd Tb Dy Ho Er Tm Yb Lu Hf Th U

Sheets:
  1. "Regression"  — 7-point log-log regression per element (loading + stripping)
                     Outputs: P, Q, R², R²_adj, SE_Q, n per element × acid system
  2. "SX Model"    — TDMA solver; Panel B P/Q linked from Regression sheet
openpyxl only
"""

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, ScatterChart, Reference, Series

# ── Element definitions ────────────────────────────────────────────────────────
ELEMENTS = [
    "Al","Sc","Fe","Co","Zn","Ga","Rb","Y","Zr",
    "La","Ce","Pr","Nd","Sm","Eu","Gd","Tb","Dy",
    "Ho","Er","Tm","Yb","Lu","Hf","Th","U",
]
N_EL = len(ELEMENTS)  # 26

MW = {
    "Al":26.98, "Sc":44.96, "Fe":55.85, "Co":58.93,
    "Zn":65.38, "Ga":69.72, "Rb":85.47, "Y":88.91,
    "Zr":91.22, "La":138.91,"Ce":140.12,"Pr":140.91,
    "Nd":144.24,"Sm":150.36,"Eu":151.96,"Gd":157.25,
    "Tb":158.93,"Dy":162.50,"Ho":164.93,"Er":167.26,
    "Tm":168.93,"Yb":173.04,"Lu":174.97,"Hf":178.49,
    "Th":232.04,"U":238.03,
}

FEED_DEF = {e: 0.1 for e in ELEMENTS}
FEED_DEF.update({"Zr":1.0,"Hf":0.5,"Th":2.0,"U":5.0,
                  "La":0.5,"Ce":0.5,"Nd":0.3,"Y":0.2})

# Default P/Q when regression has < 2 points
P_DEF = {e: 1.0 for e in ELEMENTS}
Q_DEF_LOAD = {e: -1.0 for e in ELEMENTS}
Q_DEF_STRIP= {e: -1.2 for e in ELEMENTS}

N_STAGES = 25

# ── Colour palette ─────────────────────────────────────────────────────────────
C = {
    "INPUT_BG":"FFFDE7","INPUT_FG":"1A237E",
    "CALC_BG" :"FFFFFF","CALC_FG" :"212121",
    "STEP_BG" :"E8EAF6","STEP_FG" :"212121",
    "OUT_BG"  :"F1F8E9","OUT_FG"  :"2E7D32",
    "REF_BG"  :"E3F2FD","REF_FG"  :"0D47A1",
    "TITLE"   :"1B3A6B","SECHDR"  :"2E4057","COLHDR":"4A6FA5",
    "EXT_T"   :"FFFDE7","SCR_T"   :"F3E5F5","STR_T" :"E8F5E9",
    "HDR_FG"  :"FFFFFF",
}

def fl(hex_bg): return PatternFill("solid",fgColor=hex_bg)
def fn(hex_fg,bold=False,sz=9): return Font(color=hex_fg,bold=bold,size=sz,name="Calibri")
def al(h="center"): return Alignment(horizontal=h,vertical="center",wrap_text=True)
def bd():
    s=Side(style="thin",color="BDBDBD")
    return Border(left=s,right=s,top=s,bottom=s)

def sc(ws,row,col,bg,fg,bold=False,val=None,h="center",sz=9):
    c=ws.cell(row=row,column=col)
    if val is not None: c.value=val
    c.fill=fl(bg); c.font=fn(fg,bold,sz)
    c.alignment=al(h); c.border=bd()
    return c

def hdr(ws,row,col,text,span=1,bg=None,fg=None,sz=9,bold=True):
    bg=bg or C["COLHDR"]; fg=fg or C["HDR_FG"]
    if span>1:
        ws.merge_cells(start_row=row,start_column=col,
                       end_row=row,end_column=col+span-1)
    c=ws.cell(row=row,column=col,value=text)
    c.fill=fl(bg);c.font=fn(fg,bold,sz)
    c.alignment=al();c.border=bd()
    return c

def inp(ws,row,col,val=None):
    c=ws.cell(row=row,column=col)
    if val is not None: c.value=val
    c.fill=fl(C["INPUT_BG"]);c.font=fn(C["INPUT_FG"])
    c.alignment=al();c.border=bd()
    return c

def out(ws,row,col,formula=None):
    c=ws.cell(row=row,column=col)
    if formula is not None: c.value=formula
    c.fill=fl(C["OUT_BG"]);c.font=fn(C["OUT_FG"])
    c.alignment=al();c.border=bd()
    return c

def calc(ws,row,col,formula=None):
    c=ws.cell(row=row,column=col)
    if formula is not None: c.value=formula
    c.fill=fl(C["CALC_BG"]);c.font=fn(C["CALC_FG"])
    c.alignment=al();c.border=bd()
    return c

def step(ws,row,col,formula=None):
    c=ws.cell(row=row,column=col)
    if formula is not None: c.value=formula
    c.fill=fl(C["STEP_BG"]);c.font=fn(C["STEP_FG"])
    c.alignment=al();c.border=bd()
    return c

def ref(ws,row,col,formula=None):
    c=ws.cell(row=row,column=col)
    if formula is not None: c.value=formula
    c.fill=fl(C["REF_BG"]);c.font=fn(C["REF_FG"])
    c.alignment=al("left");c.border=bd()
    return c

def cr(row,col): return f"${get_column_letter(col)}${row}"
def ad(row,col): return f"{get_column_letter(col)}{row}"

# ── Workbook ───────────────────────────────────────────────────────────────────
wb = Workbook()

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 1 — REGRESSION
# Layout:
#   Panel L  (Loading  H2SO4): rows 4-onwards, cols 1 to 2+2*N_EL
#   Panel S  (Stripping HNO3): same row range, starts at col 2+2*N_EL+2
#   D-value table (7 rows × 26 elem) per panel — auto-calculated
#   Regression summary table (26 rows × 8 cols):
#     Element | n | P_load | Q_load | R²_load | R²adj_load | SE_Q_load |
#             | P_strip | Q_strip | R²_strip | R²adj_strip | SE_Q_strip
# ══════════════════════════════════════════════════════════════════════════════
wr = wb.active
wr.title = "Regression"
wr.sheet_properties.tabColor = "2E7D32"

N_PTS = 7   # experimental data points per element per acid system

for col in range(1, 300):
    wr.column_dimensions[get_column_letter(col)].width = 11
wr.column_dimensions["A"].width = 14

# ── Title ──────────────────────────────────────────────────────────────────────
wr.row_dimensions[1].height = 26
total_cols = 2 + 2 * N_EL
wr.merge_cells(f"A1:{get_column_letter(total_cols * 2 + 3)}1")
c = wr.cell(row=1, column=1,
    value=("P/Q LOG-LOG REGRESSION  |  D = P × [acid]^Q  |  "
           "7 experimental points × 26 elements × 2 acid systems  "
           "(H₂SO₄ loading  /  HNO₃ stripping)"))
c.fill = fl("2E7D32"); c.font = fn("FFFFFF", bold=True, sz=12); c.alignment = al()

# Colour legend row 2
wr.row_dimensions[2].height = 14
for i, (lbl, bg, fg) in enumerate([
    ("INPUT", C["INPUT_BG"], C["INPUT_FG"]),
    ("CALC",  C["CALC_BG"],  C["CALC_FG"]),
    ("OUTPUT",C["OUT_BG"],   C["OUT_FG"]),
]):
    col_s = 1 + i * 4
    wr.merge_cells(start_row=2, start_column=col_s, end_row=2, end_column=col_s+3)
    c = wr.cell(row=2, column=col_s, value=lbl)
    c.fill = fl(bg); c.font = fn(fg, bold=True, sz=9); c.alignment = al()

wr.row_dimensions[3].height = 5

# ── Input table builder (shared for Loading and Stripping panels) ───────────
# Each panel:
#   col base+0       : Point#  (1-7)
#   col base+1       : [acid] N  ← INPUT yellow
#   col base+2 to base+1+N_EL   : C_aq per element  ← INPUT
#   col base+2+N_EL  to base+1+2*N_EL : C_org per element ← INPUT
PANEL_W_R = 2 + 2 * N_EL   # 54 cols

R_LOAD_COL  = 1
R_STRIP_COL = R_LOAD_COL + PANEL_W_R + 2

def build_reg_panel(ws_r, base_col, title, bg_title, acid_label, pt_row_start):
    """Build input + D-calc block for one acid system. Returns first data row."""
    ws_r.row_dimensions[pt_row_start - 2].height = 18
    hdr(ws_r, pt_row_start - 2, base_col, title, span=PANEL_W_R,
        bg=bg_title, fg="FFFFFF", sz=11)

    # Column headers
    ws_r.row_dimensions[pt_row_start - 1].height = 34
    hdr(ws_r, pt_row_start - 1, base_col,     "Pt#", sz=9)
    hdr(ws_r, pt_row_start - 1, base_col + 1, f"{acid_label}\n(N)", sz=9)
    for i, e in enumerate(ELEMENTS):
        hdr(ws_r, pt_row_start - 1, base_col + 2 + i,        f"C_aq\n{e}\n(ppm)", sz=8)
        hdr(ws_r, pt_row_start - 1, base_col + 2 + N_EL + i, f"C_org\n{e}\n(ppm)", sz=8)

    # 7 data rows
    for pt in range(1, N_PTS + 1):
        r = pt_row_start + pt - 1
        ws_r.row_dimensions[r].height = 15
        sc(ws_r, r, base_col, C["STEP_BG"], C["STEP_FG"], val=pt)
        inp(ws_r, r, base_col + 1)   # [acid]
        for i in range(N_EL):
            inp(ws_r, r, base_col + 2 + i)           # C_aq
            inp(ws_r, r, base_col + 2 + N_EL + i)    # C_org

    return pt_row_start

PT_ROW_BASE = 6   # first actual data row is 6 (headers at 4-5)

R_L_DATA = build_reg_panel(wr, R_LOAD_COL,  "LOADING DATA  (H₂SO₄ system)",
                            C["COLHDR"], "H₂SO₄", PT_ROW_BASE)
R_S_DATA = build_reg_panel(wr, R_STRIP_COL, "STRIPPING DATA  (HNO₃ system)",
                            "1B5E20",    "HNO₃",  PT_ROW_BASE)

# ── D-value table (auto-calculated from input rows) ───────────────────────────
D_HDR_R = PT_ROW_BASE + N_PTS + 1   # row 14

wr.row_dimensions[D_HDR_R].height = 18
D_PANEL_W_R = 1 + N_EL
hdr(wr, D_HDR_R, R_LOAD_COL,  "D VALUES  (C_org / C_aq)  — Loading",
    span=D_PANEL_W_R, bg=C["COLHDR"])
hdr(wr, D_HDR_R, R_STRIP_COL, "D VALUES  (C_org / C_aq)  — Stripping",
    span=D_PANEL_W_R, bg="1B5E20", fg="FFFFFF")

D_COL_HDR_R = D_HDR_R + 1
wr.row_dimensions[D_COL_HDR_R].height = 16
hdr(wr, D_COL_HDR_R, R_LOAD_COL,  "Pt#")
hdr(wr, D_COL_HDR_R, R_STRIP_COL, "Pt#")
for i, e in enumerate(ELEMENTS):
    hdr(wr, D_COL_HDR_R, R_LOAD_COL  + 1 + i, f"D_{e}", sz=8)
    hdr(wr, D_COL_HDR_R, R_STRIP_COL + 1 + i, f"D_{e}", sz=8)

D_DATA_R = D_COL_HDR_R + 1   # row 16
for pt in range(1, N_PTS + 1):
    r = D_DATA_R + pt - 1
    wr.row_dimensions[r].height = 15
    inp_r = R_L_DATA + pt - 1
    sc(wr, r, R_LOAD_COL,  C["STEP_BG"], C["STEP_FG"], val=pt)
    sc(wr, r, R_STRIP_COL, C["STEP_BG"], C["STEP_FG"], val=pt)
    for e_idx in range(N_EL):
        # Loading D
        aq_col  = R_LOAD_COL + 2 + e_idx
        org_col = R_LOAD_COL + 2 + N_EL + e_idx
        aqa  = ad(inp_r, aq_col);  orga = ad(inp_r, org_col)
        calc(wr, r, R_LOAD_COL + 1 + e_idx,
             f'=IFERROR(IF(OR({aqa}="",{orga}="",VALUE({aqa})=0),"",{orga}/{aqa}),"N/A")')
        # Stripping D
        aq_col  = R_STRIP_COL + 2 + e_idx
        org_col = R_STRIP_COL + 2 + N_EL + e_idx
        inp_r_s = R_S_DATA + pt - 1
        aqa  = ad(inp_r_s, aq_col); orga = ad(inp_r_s, org_col)
        calc(wr, r, R_STRIP_COL + 1 + e_idx,
             f'=IFERROR(IF(OR({aqa}="",{orga}="",VALUE({aqa})=0),"",{orga}/{aqa}),"N/A")')

# ── Regression summary table ───────────────────────────────────────────────────
RS_R_HDR = D_DATA_R + N_PTS + 1   # row 24

wr.row_dimensions[RS_R_HDR].height = 18
sum_cols = 13
wr.merge_cells(f"A{RS_R_HDR}:{get_column_letter(sum_cols)}{RS_R_HDR}")
c = wr.cell(row=RS_R_HDR, column=1,
    value="REGRESSION SUMMARY  (n=7 points per element per acid system)")
c.fill = fl(C["SECHDR"]); c.font = fn("FFFFFF", bold=True, sz=11); c.alignment = al()

RS_R_COL_HDR = RS_R_HDR + 1
wr.row_dimensions[RS_R_COL_HDR].height = 42
for i, lbl in enumerate([
    "Element", "n_valid",
    "P_load", "Q_load", "R²_load", "R²adj_load", "SE_Q_load",
    "P_strip","Q_strip","R²_strip","R²adj_strip","SE_Q_strip",
    "Notes",
]):
    hdr(wr, RS_R_COL_HDR, 1 + i, lbl, sz=9,
        bg=C["EXT_T"] if i < 7 else C["STR_T"],
        fg=C["SECHDR"])

RS_R_DATA = RS_R_COL_HDR + 1   # row 26  ← Panel B in SX Model references here

# Helper: range for D values of one element across N_PTS rows
def d_range_r(base_col, e_idx):
    col = base_col + 1 + e_idx
    return f"{ad(D_DATA_R, col)}:{ad(D_DATA_R + N_PTS - 1, col)}"

def acid_range_r(base_col):
    col = base_col + 1   # [acid] column
    return f"{ad(R_L_DATA if base_col==R_LOAD_COL else R_S_DATA, col)}:{ad((R_L_DATA if base_col==R_LOAD_COL else R_S_DATA)+N_PTS-1, col)}"

# Regression formulas using ISNUMBER to filter blank / "N/A" D cells
def reg_n(d_rng):
    return f'=SUMPRODUCT((ISNUMBER({d_rng})*1))'

def reg_P(d_rng, acid_rng):
    return (f'=IFERROR(IF(SUMPRODUCT(ISNUMBER({d_rng})*1)<2,"Need≥2",'
            f'EXP(INTERCEPT('
            f'IF(ISNUMBER({d_rng}),LN({d_rng}),FALSE),'
            f'IF(ISNUMBER({d_rng}),LN({acid_rng}),FALSE)))),"N/A")')

def reg_Q(d_rng, acid_rng):
    return (f'=IFERROR(IF(SUMPRODUCT(ISNUMBER({d_rng})*1)<2,"Need≥2",'
            f'SLOPE('
            f'IF(ISNUMBER({d_rng}),LN({d_rng}),FALSE),'
            f'IF(ISNUMBER({d_rng}),LN({acid_rng}),FALSE))),"N/A")')

def reg_R2(d_rng, acid_rng):
    return (f'=IFERROR(IF(SUMPRODUCT(ISNUMBER({d_rng})*1)<2,"Need≥2",'
            f'RSQ('
            f'IF(ISNUMBER({d_rng}),LN({d_rng}),FALSE),'
            f'IF(ISNUMBER({d_rng}),LN({acid_rng}),FALSE))),"N/A")')

def reg_R2adj(d_rng, acid_rng, n_cell):
    """R²_adj = 1 - (1-R²)*(n-1)/(n-2)  for simple linear regression (1 predictor)"""
    r2 = reg_R2(d_rng, acid_rng).lstrip("=")
    return (f'=IFERROR(IF(SUMPRODUCT(ISNUMBER({d_rng})*1)<3,"Need≥3",'
            f'1-(1-({r2}))*({n_cell}-1)/({n_cell}-2)),"N/A")')

def reg_SEQ(d_rng, acid_rng, n_cell):
    """SE(Q) = sqrt( (1-R²) / (R²*(n-2)) ) * STDEV(ln D) / STDEV(ln acid)
       Simplified: SE_b from OLS = sqrt(MSE / Sxx)
       = sqrt( (1-R²)*VAR(lnD) / ((n-2)*VAR(lnAcid)) )  — valid for LN-LN"""
    r2 = reg_R2(d_rng, acid_rng).lstrip("=")
    return (f'=IFERROR(IF(SUMPRODUCT(ISNUMBER({d_rng})*1)<3,"Need≥3",'
            f'SQRT((1-({r2}))'
            f'*VAR(IF(ISNUMBER({d_rng}),LN({d_rng}),FALSE))'
            f'/(({n_cell}-2)'
            f'*VAR(IF(ISNUMBER({d_rng}),LN({acid_rng}),FALSE))))),"N/A")')

# Store references for SX Model to use
REG_SHEET_CELLS = {}   # elem -> {"P_load":(row,col), ...}  on Regression sheet

load_acid_r = acid_range_r(R_LOAD_COL)
strip_acid_r= acid_range_r(R_STRIP_COL)

for e_idx, elem in enumerate(ELEMENTS):
    r = RS_R_DATA + e_idx
    wr.row_dimensions[r].height = 15
    sc(wr, r, 1, C["STEP_BG"], C["STEP_FG"], val=elem)

    ld_d = d_range_r(R_LOAD_COL,  e_idx)
    st_d = d_range_r(R_STRIP_COL, e_idx)
    n_cell_ld = f"B{r}"   # n_valid col

    # n_valid (loading — same for both since same 7 pts)
    calc(wr, r, 2, reg_n(ld_d))
    nc = f"B{r}"

    out(wr,  r,  3, reg_P(ld_d,  load_acid_r))
    out(wr,  r,  4, reg_Q(ld_d,  load_acid_r))
    calc(wr, r,  5, reg_R2(ld_d, load_acid_r))
    calc(wr, r,  6, reg_R2adj(ld_d, load_acid_r, nc))
    calc(wr, r,  7, reg_SEQ(ld_d,  load_acid_r, nc))
    out(wr,  r,  8, reg_P(st_d,  strip_acid_r))
    out(wr,  r,  9, reg_Q(st_d,  strip_acid_r))
    calc(wr, r, 10, reg_R2(st_d, strip_acid_r))
    calc(wr, r, 11, reg_R2adj(st_d, strip_acid_r, nc))
    calc(wr, r, 12, reg_SEQ(st_d,  strip_acid_r, nc))

    # Quality flag: R²_adj < 0.85 → "Check fit"
    r2adj_addr = ad(r, 6)
    calc(wr, r, 13,
         f'=IF(ISNUMBER({r2adj_addr}),IF({r2adj_addr}<0.85,"⚠ Check fit","OK"),"")')

    REG_SHEET_CELLS[elem] = {
        "P_load" : (r, 3), "Q_load" : (r, 4),
        "P_strip": (r, 8), "Q_strip": (r, 9),
    }

# ── Scatter charts: one per element (log D vs log acid, loading + stripping) ──
# Only generate for the 4 key elements to keep file light; add note for others.
CHART_ELEMS_R = ["U", "Th", "Hf", "Zr"]
for i, elem in enumerate(CHART_ELEMS_R):
    e_idx = ELEMENTS.index(elem)
    chart = ScatterChart()
    chart.title = f"{elem} — log D vs log [acid]"
    chart.style = 10
    chart.x_axis.title = "LN([acid])"
    chart.y_axis.title = "LN(D)"
    chart.height = 10; chart.width = 16

    for panel_col, panel_label, color in [
        (R_LOAD_COL,  "Loading",   "1B3A6B"),
        (R_STRIP_COL, "Stripping", "2E7D32"),
    ]:
        d_col   = panel_col + 1 + e_idx
        acid_col= panel_col + 1
        # We can't plot LN() directly from chart references; reference raw D rows
        # and note that D is on linear scale. Use D values as proxy.
        d_ref   = Reference(wr, min_col=d_col,   min_row=D_DATA_R, max_row=D_DATA_R+N_PTS-1)
        acid_ref= Reference(wr, min_col=acid_col, min_row=R_L_DATA if panel_col==R_LOAD_COL else R_S_DATA,
                            max_row=(R_L_DATA if panel_col==R_LOAD_COL else R_S_DATA)+N_PTS-1)
        s = Series(d_ref, xvalues=acid_ref, title=panel_label)
        s.marker.symbol = "circle" if panel_col == R_LOAD_COL else "triangle"
        s.marker.size = 5
        s.graphicalProperties.line.noFill = True
        s.graphicalProperties.solidFill = color
        chart.series.append(s)

    anchor_row = RS_R_DATA + N_EL + 3 + i * 22
    wr.add_chart(chart, f"A{anchor_row}")

# Note for remaining elements
note_r = RS_R_DATA + N_EL + 2
wr.merge_cells(f"A{note_r}:N{note_r}")
c = wr.cell(row=note_r, column=1,
    value="Charts above: U, Th, Hf, Zr (key elements). "
          "All 26 regressions computed in the summary table above.")
c.fill = fl(C["REF_BG"]); c.font = fn(C["REF_FG"], sz=9)

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 2 — SX MODEL  (add after Regression sheet)
# ══════════════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("SX Model")
ws.sheet_properties.tabColor = C["TITLE"]

# REG_CELLS now points to Regression sheet — cross-sheet references
REG_CELLS = {}
for elem in ELEMENTS:
    rc = REG_SHEET_CELLS[elem]
    REG_CELLS[elem] = {
        "P_load" : rc["P_load"],
        "Q_load" : rc["Q_load"],
        "P_strip": rc["P_strip"],
        "Q_strip": rc["Q_strip"],
    }

# Default column width
for col in range(1,600):
    ws.column_dimensions[get_column_letter(col)].width = 11
ws.column_dimensions["A"].width = 20

# ──────────────────────────────────────────────────────────────────────────────
# ROW 1  Title
# ──────────────────────────────────────────────────────────────────────────────
ws.row_dimensions[1].height = 28
ws.merge_cells("A1:BZ1")
c=ws["A1"]
c.value=("STEADY-STATE SX SIMULATION — Primene JMT / H₂SO₄ Loading / HNO₃ Stripping — "
         "26 Elements: Al Sc Fe Co Zn Ga Rb Y Zr La Ce Pr Nd Sm Eu Gd Tb Dy Ho Er Tm Yb Lu Hf Th U")
c.fill=fl(C["TITLE"]); c.font=fn("FFFFFF",bold=True,sz=13); c.alignment=al()

# ROW 2  Legend
ws.row_dimensions[2].height = 16
legend=[("INPUT","FFFDE7",C["INPUT_FG"]),("CALC","FFFFFF",C["CALC_FG"]),
        ("STEP","E8EAF6",C["STEP_FG"]),("OUTPUT","F1F8E9",C["OUT_FG"]),
        ("REF","E3F2FD",C["REF_FG"])]
col=1
for lbl,bg,fg in legend:
    ws.merge_cells(start_row=2,start_column=col,end_row=2,end_column=col+2)
    c=ws.cell(row=2,column=col,value=lbl)
    c.fill=fl(bg);c.font=fn(fg,bold=True,sz=9);c.alignment=al()
    col+=3

ws.row_dimensions[3].height=5

# Note: Regression is on the "Regression" sheet. Panel B P/Q values are
# cross-sheet linked from Regression!cols C-D (P_load,Q_load) and H-I (P_strip,Q_strip).

# ══════════════════════════════════════════════════════════════════════════════
# TDMA INPUT PANELS  (start at row 4 — no Section 0 needed here)
# ══════════════════════════════════════════════════════════════════════════════
TDMA_START = 4

ws.row_dimensions[TDMA_START].height=20
ws.merge_cells(f"A{TDMA_START}:{get_column_letter(30)}{TDMA_START}")
c=ws.cell(row=TDMA_START,column=1,value="TDMA SOLVER — INPUT PANELS")
c.fill=fl(C["TITLE"]);c.font=fn("FFFFFF",bold=True,sz=12);c.alignment=al()

# Panel A — cols 1-2
PA_ROW = TDMA_START+1
ws.merge_cells(f"A{PA_ROW}:B{PA_ROW}")
hdr(ws,PA_ROW,1,"PANEL A — General Inputs",span=2,bg=C["SECHDR"])

panel_a=[
    ("n_ext (extraction stages)",3),
    ("n_scr (scrub stages)",0),
    ("n_str (strip stages)",2),
    ("[H₂SO₄] Extraction (N)",0.5),
    ("[H₂SO₄] Scrub (N)",1.0),
    ("[HNO₃] Strip (N)",4.0),
    ("F — Feed flow aq (L/min)",1.0),
    ("S — Strip flow aq (L/min)",1.0),
    ("O — Org flow (L/min)",1.0),
    ("R — Recycle ratio",0.0),
]
PA_DATA = PA_ROW+1
PA_CELLS={}
for i,(lbl,val) in enumerate(panel_a):
    r=PA_DATA+i
    ws.row_dimensions[r].height=15
    ref(ws,r,1,lbl)
    inp(ws,r,2,val)
    PA_CELLS[lbl]=(r,2)

def pa(key): r,c=PA_CELLS[key]; return cr(r,c)
N_EXT_REF  = pa("n_ext (extraction stages)")
N_SCR_REF  = pa("n_scr (scrub stages)")
N_STR_REF  = pa("n_str (strip stages)")
AE_REF     = pa("[H₂SO₄] Extraction (N)")
AS_REF     = pa("[H₂SO₄] Scrub (N)")
AN_REF     = pa("[HNO₃] Strip (N)")
F_REF      = pa("F — Feed flow aq (L/min)")
S_REF      = pa("S — Strip flow aq (L/min)")
O_REF      = pa("O — Org flow (L/min)")
R_REF      = pa("R — Recycle ratio")

# Panel B — Distribution Coefficients (cols 4-8)
PB_COL=4
PB_ROW=TDMA_START+1
ws.merge_cells(start_row=PB_ROW,start_column=PB_COL,
               end_row=PB_ROW,end_column=PB_COL+5)
hdr(ws,PB_ROW,PB_COL,"PANEL B — Distribution Coefficients (live-linked from Regression)",
    span=6,bg=C["SECHDR"])
for i,lbl in enumerate(["Element","P_load","Q_load","P_strip","Q_strip","Note"]):
    hdr(ws,PB_ROW+1,PB_COL+i,lbl)

PB_CELLS={}
PB_DATA=PB_ROW+2
for e_idx,elem in enumerate(ELEMENTS):
    r=PB_DATA+e_idx
    ws.row_dimensions[r].height=15
    sc(ws,r,PB_COL,C["STEP_BG"],C["STEP_FG"],val=elem)
    reg=REG_CELLS[elem]
    def linked(rr,cc,defval):
        ref_=f"Regression!${get_column_letter(cc)}${rr}"
        return f'=IFERROR(IF(ISNUMBER({ref_}),{ref_},{defval}),{defval})'
    out(ws,r,PB_COL+1, linked(*reg["P_load"], P_DEF[elem]))
    out(ws,r,PB_COL+2, linked(*reg["Q_load"], Q_DEF_LOAD[elem]))
    out(ws,r,PB_COL+3, linked(*reg["P_strip"],P_DEF[elem]))
    out(ws,r,PB_COL+4, linked(*reg["Q_strip"],Q_DEF_STRIP[elem]))
    c_=ws.cell(row=r,column=PB_COL+5,value="→ regression")
    c_.fill=fl(C["CALC_BG"]);c_.font=fn(C["CALC_FG"],sz=8);c_.alignment=al()
    PB_CELLS[elem]={
        "P_load" :cr(r,PB_COL+1), "Q_load" :cr(r,PB_COL+2),
        "P_strip":cr(r,PB_COL+3), "Q_strip":cr(r,PB_COL+4),
    }

# Panel C — Feed Concentrations (cols 10-12)
PC_COL=10
PC_ROW=TDMA_START+1
ws.merge_cells(start_row=PC_ROW,start_column=PC_COL,
               end_row=PC_ROW,end_column=PC_COL+2)
hdr(ws,PC_ROW,PC_COL,"PANEL C — Feed Conc. (g/L)",span=3,bg=C["SECHDR"])
hdr(ws,PC_ROW+1,PC_COL,"Element")
hdr(ws,PC_ROW+1,PC_COL+1,"X_feed (g/L)")
hdr(ws,PC_ROW+1,PC_COL+2,"MW (g/mol)")

PC_CELLS={}
PC_DATA=PC_ROW+2
for e_idx,elem in enumerate(ELEMENTS):
    r=PC_DATA+e_idx
    ws.row_dimensions[r].height=15
    sc(ws,r,PC_COL,  C["STEP_BG"],C["STEP_FG"],val=elem)
    inp(ws,r,PC_COL+1, FEED_DEF[elem])
    sc(ws,r,PC_COL+2, C["CALC_BG"],C["CALC_FG"],val=MW[elem])
    PC_CELLS[elem]={"XF":cr(r,PC_COL+1)}

# Panel D — Results (cols 14-16)
PD_COL=14
PD_ROW=TDMA_START+1
ws.merge_cells(start_row=PD_ROW,start_column=PD_COL,
               end_row=PD_ROW,end_column=PD_COL+1)
hdr(ws,PD_ROW,PD_COL,"PANEL D — Results Summary",span=2,bg=C["SECHDR"])
hdr(ws,PD_ROW+1,PD_COL,"Metric")
hdr(ws,PD_ROW+1,PD_COL+1,"Value (g/L)")

# Panel E — O/A Ratios (cols 17-19)
PE_COL=17
PE_ROW=TDMA_START+1
ws.merge_cells(start_row=PE_ROW,start_column=PE_COL,
               end_row=PE_ROW,end_column=PE_COL+2)
hdr(ws,PE_ROW,PE_COL,"PANEL E — O/A Ratio Tracker",span=3,bg=C["SECHDR"])
for i,lbl in enumerate(["Section","O/A","Note"]):
    hdr(ws,PE_ROW+1,PE_COL+i,lbl)
pe_data=[("Extraction",f"={O_REF}/{F_REF}","O/F"),
         ("Scrub",f"={O_REF}/MAX({S_REF}*{R_REF},0.0001)","O/(S·R)"),
         ("Strip",f"={O_REF}/{S_REF}","O/S")]
PE_DATA=PE_ROW+2
for i,(sec,formula,note) in enumerate(pe_data):
    r=PE_DATA+i
    ws.row_dimensions[r].height=15
    sc(ws,r,PE_COL,C["STEP_BG"],C["STEP_FG"],val=sec)
    out(ws,r,PE_COL+1,formula)
    sc(ws,r,PE_COL+2,C["CALC_BG"],C["CALC_FG"],val=note)

# ══════════════════════════════════════════════════════════════════════════════
# TDMA SOLVER  (25 stages × 26 elements)
# Column layout per element: D a b c d c' d' y_org  =8 cols
# After all elements: x_aq per element (26 cols), total_org, total_aq
# ══════════════════════════════════════════════════════════════════════════════
# Ensure solver starts well below all panels
PANEL_LAST_ROW = max(TDMA_START+1+len(panel_a)+1,
                     PB_DATA+N_EL,
                     PC_DATA+N_EL,
                     PD_ROW+2+N_EL*3,
                     PE_DATA+3) + 3

SOLVER_START = PANEL_LAST_ROW

ws.row_dimensions[SOLVER_START].height=20
COLS_PER_ELEM = 8
ELEM_BASE = 2   # col B
X_AQ_BASE = ELEM_BASE + N_EL*COLS_PER_ELEM
TOT_ORG_COL = X_AQ_BASE + N_EL
TOT_AQ_COL  = TOT_ORG_COL + 1

end_solver_col = get_column_letter(TOT_AQ_COL)
ws.merge_cells(f"A{SOLVER_START}:{end_solver_col}{SOLVER_START}")
c=ws.cell(row=SOLVER_START,column=1,value="TDMA SOLVER — 25 STAGES × 26 ELEMENTS")
c.fill=fl(C["TITLE"]);c.font=fn("FFFFFF",bold=True,sz=12);c.alignment=al()

HDR1=SOLVER_START+1
HDR2=SOLVER_START+2
ws.row_dimensions[HDR1].height=16; ws.row_dimensions[HDR2].height=28

hdr(ws,HDR1,1,"Stage",bg=C["SECHDR"])
ws.merge_cells(start_row=HDR1,start_column=1,end_row=HDR2,end_column=1)

ELEM_COLS={}
for e_idx,elem in enumerate(ELEMENTS):
    base=ELEM_BASE+e_idx*COLS_PER_ELEM
    ws.merge_cells(start_row=HDR1,start_column=base,
                   end_row=HDR1,end_column=base+7)
    hdr(ws,HDR1,base,f"{elem}",span=8,bg=C["COLHDR"])
    for i,lbl in enumerate(["D_n","a_n","b_n","c_n","d_n","c'_n","d'_n","y_org"]):
        hdr(ws,HDR2,base+i,lbl,sz=8)
    x_col=X_AQ_BASE+e_idx
    ws.merge_cells(start_row=HDR1,start_column=x_col,
                   end_row=HDR2,end_column=x_col)
    hdr(ws,HDR1,x_col,f"x_aq\n{elem}",bg=C["OUT_BG"],fg=C["OUT_FG"],sz=8)
    ELEM_COLS[elem]={
        "D":base,"a":base+1,"b":base+2,"c":base+3,
        "d":base+4,"cp":base+5,"dp":base+6,"y_org":base+7,
        "x_aq":x_col
    }

ws.merge_cells(start_row=HDR1,start_column=TOT_ORG_COL,
               end_row=HDR2,end_column=TOT_ORG_COL)
hdr(ws,HDR1,TOT_ORG_COL,"ΣOrg",sz=8)
ws.merge_cells(start_row=HDR1,start_column=TOT_AQ_COL,
               end_row=HDR2,end_column=TOT_AQ_COL)
hdr(ws,HDR1,TOT_AQ_COL,"ΣAq",sz=8)

STAGE_ROW_START=HDR2+1

# ── Stage data rows ─────────────────────────────────────────────────────────
for n in range(1,N_STAGES+1):
    r=STAGE_ROW_START+n-1
    ws.row_dimensions[r].height=14
    lbl_c=ws.cell(row=r,column=1,value=n)
    lbl_c.font=fn(C["STEP_FG"],bold=True,sz=9)
    lbl_c.alignment=al()
    lbl_c.border=bd()
    if n<=3:   lbl_c.fill=fl(C["EXT_T"])
    elif n<=5: lbl_c.fill=fl(C["STR_T"])
    else:      lbl_c.fill=fl(C["CALC_BG"])

    for elem in ELEMENTS:
        ec=ELEM_COLS[elem]
        pb=PB_CELLS[elem]
        D_a=ad(r,ec["D"]); b_a=ad(r,ec["b"])
        c_a=ad(r,ec["c"]); d_a=ad(r,ec["d"])
        cp_a=ad(r,ec["cp"]); dp_a=ad(r,ec["dp"])

        # D_n: branch on section
        d_f=(f'=IF({n}>{N_EXT_REF}+{N_SCR_REF},'
             f'{pb["P_strip"]}*{AN_REF}^{pb["Q_strip"]},'
             f'{pb["P_load"]}*{AE_REF}^{pb["Q_load"]})')
        step(ws,r,ec["D"],d_f)

        # a_n
        step(ws,r,ec["a"],f'=IF({n}=1,0,{O_REF})')

        # b_n
        Qaq=(f'IF({n}<={N_EXT_REF},{S_REF}*{R_REF}+{F_REF},'
             f'IF({n}<={N_EXT_REF}+{N_SCR_REF},{S_REF}*{R_REF},{S_REF}))')
        step(ws,r,ec["b"],f'=-({O_REF}+({Qaq})/{D_a})')

        # c_n
        if n<N_STAGES:
            D_next=ad(r+1,ec["D"])
            c_f=(f'=IF({n}={N_EXT_REF}+{N_SCR_REF},0,'
                 f'IF({n}>{N_EXT_REF}+{N_SCR_REF},{S_REF}/{D_next},'
                 f'IF({n}>={N_EXT_REF},{S_REF}*{R_REF}/{D_next},'
                 f'({S_REF}*{R_REF}+{F_REF})/{D_next})))')
        else:
            c_f='=0'
        step(ws,r,ec["c"],c_f)

        # d_n
        XF=PC_CELLS[elem]["XF"]
        step(ws,r,ec["d"],f'=IF({n}={N_EXT_REF},-{XF}*{F_REF},0)')

        # Forward sweep
        if n==1:
            step(ws,r,ec["cp"],f'={c_a}/{b_a}')
            step(ws,r,ec["dp"],f'={d_a}/{b_a}')
        else:
            pcp=ad(r-1,ec["cp"]); pdp=ad(r-1,ec["dp"])
            denom=f'({b_a}-{ad(r,ec["a"])}*{pcp})'
            step(ws,r,ec["cp"],f'={c_a}/{denom}')
            step(ws,r,ec["dp"],f'=({d_a}-{ad(r,ec["a"])}*{pdp})/{denom}')

# Back substitution (must go in reverse)
for n in range(N_STAGES,0,-1):
    r=STAGE_ROW_START+n-1
    for elem in ELEMENTS:
        ec=ELEM_COLS[elem]
        if n==N_STAGES:
            out(ws,r,ec["y_org"],f'={ad(r,ec["dp"])}')
        else:
            out(ws,r,ec["y_org"],
                f'={ad(r,ec["dp"])}-{ad(r,ec["cp"])}*{ad(r+1,ec["y_org"])}')

# x_aq and totals
for n in range(1,N_STAGES+1):
    r=STAGE_ROW_START+n-1
    y_parts=[]; x_parts=[]
    for elem in ELEMENTS:
        ec=ELEM_COLS[elem]
        calc(ws,r,ec["x_aq"],
             f'=IFERROR({ad(r,ec["y_org"])}/{ad(r,ec["D"])},0)')
        y_parts.append(ad(r,ec["y_org"]))
        x_parts.append(ad(r,ec["x_aq"]))
    calc(ws,r,TOT_ORG_COL,"="+"+".join(y_parts))
    calc(ws,r,TOT_AQ_COL, "="+"+".join(x_parts))

# ── Panel D — Results ──────────────────────────────────────────────────────
PD_DATA=PD_ROW+2
for e_idx,elem in enumerate(ELEMENTS):
    ec=ELEM_COLS[elem]
    r_raff=PD_DATA+e_idx*3
    r_load=r_raff+1
    r_preg=r_raff+2
    for r_ in [r_raff,r_load,r_preg]:
        ws.row_dimensions[r_].height=15

    ref(ws,r_raff,PD_COL,f"Raffinate {elem}")
    out(ws,r_raff,PD_COL+1,f'={ad(STAGE_ROW_START,ec["x_aq"])}')

    ref(ws,r_load,PD_COL,f"Loaded Org {elem}")
    ycol=get_column_letter(ec["y_org"])
    yrange=f'{ycol}{STAGE_ROW_START}:{ycol}{STAGE_ROW_START+N_STAGES-1}'
    out(ws,r_load,PD_COL+1,f'=IFERROR(INDEX({yrange},{N_EXT_REF}),0)')

    ref(ws,r_preg,PD_COL,f"Pregnant {elem}")
    xcol=get_column_letter(ec["x_aq"])
    xrange=f'{xcol}{STAGE_ROW_START}:{xcol}{STAGE_ROW_START+N_STAGES-1}'
    out(ws,r_preg,PD_COL+1,
        f'=IFERROR(INDEX({xrange},{N_EXT_REF}+{N_SCR_REF}+{N_STR_REF}),0)')

# ══════════════════════════════════════════════════════════════════════════════
# STAGE PROFILE TABLE  (for charts)
# ══════════════════════════════════════════════════════════════════════════════
PROF_START=STAGE_ROW_START+N_STAGES+3
ws.row_dimensions[PROF_START].height=18
prof_end_col=1+N_EL
ws.merge_cells(f"A{PROF_START}:{get_column_letter(prof_end_col)}{PROF_START}")
hdr(ws,PROF_START,1,"STAGE AQUEOUS PROFILES  (g/L)",span=prof_end_col,bg=C["SECHDR"])

PROF_HDR=PROF_START+1
ws.row_dimensions[PROF_HDR].height=16
hdr(ws,PROF_HDR,1,"Stage")
for e_idx,elem in enumerate(ELEMENTS):
    hdr(ws,PROF_HDR,2+e_idx,f"x_aq {elem}",sz=8)

for n in range(1,N_STAGES+1):
    r=PROF_HDR+n
    ws.row_dimensions[r].height=13
    sc(ws,r,1,C["STEP_BG"],C["STEP_FG"],val=n)
    for e_idx,elem in enumerate(ELEMENTS):
        ec=ELEM_COLS[elem]
        src=ad(STAGE_ROW_START+n-1,ec["x_aq"])
        calc(ws,r,2+e_idx,f'={src}')

# ── 4 selected line charts (U Th Hf Zr) to keep file manageable ──────────────
chart_elems=["U","Th","Hf","Zr"]
for i,elem in enumerate(chart_elems):
    e_idx=ELEMENTS.index(elem)
    chart=LineChart()
    chart.title=f"{elem} — Aqueous Profile vs Stage"
    chart.style=10
    chart.y_axis.title="x_aq (g/L)"
    chart.x_axis.title="Stage"
    chart.height=10; chart.width=18
    data_ref=Reference(ws,min_col=2+e_idx,
                       min_row=PROF_HDR,max_row=PROF_HDR+N_STAGES)
    cats=Reference(ws,min_col=1,
                   min_row=PROF_HDR+1,max_row=PROF_HDR+N_STAGES)
    chart.add_data(data_ref,titles_from_data=True)
    chart.set_categories(cats)
    anchor_row=PROF_HDR+N_STAGES+3+i*22
    ws.add_chart(chart,f"A{anchor_row}")

# ══════════════════════════════════════════════════════════════════════════════
# ORGANIC LOADING SECTION
# ══════════════════════════════════════════════════════════════════════════════
ORG_START=PROF_HDR+N_STAGES+100
ws.row_dimensions[ORG_START].height=20
ws.merge_cells(f"A{ORG_START}:N{ORG_START}")
c=ws.cell(row=ORG_START,column=1,
          value="ORGANIC LOADING SECTION — Primene JMT Amine Capacity")
c.fill=fl(C["SECHDR"]);c.font=fn("FFFFFF",bold=True,sz=12);c.alignment=al()

org_inp=[
    ("Extractant","Primene JMT"),
    ("MW Extractant (g/mol)",353.7),
    ("Density extractant (g/mL)",0.812),
    ("Vol% in diluent",10.0),
    ("Diluent","Orfom"),
    ("Density diluent (g/mL)",0.785),
    ("Note — saponification","N/A — primary amine"),
    ("Experimental loading (%)",80.0),
]
ORG_DATA=ORG_START+1
for i,(lbl,val) in enumerate(org_inp):
    r=ORG_DATA+i
    ws.row_dimensions[r].height=15
    ref(ws,r,1,lbl)
    if isinstance(val,str):
        sc(ws,r,2,C["CALC_BG"],C["CALC_FG"],val=val,h="left")
    else:
        inp(ws,r,2,val)

VOL_R  =ORG_DATA+3
MW_R   =ORG_DATA+1
DENS_R =ORG_DATA+2
EXP_R  =ORG_DATA+7

calc_r=ORG_DATA+len(org_inp)
ws.row_dimensions[calc_r].height=15
ref(ws,calc_r,1,"Amine conc (mol/L)")
AMINE_R=calc_r
calc(ws,calc_r,2,
     f'=({cr(VOL_R,2)}/100)*{cr(DENS_R,2)}*1000/{cr(MW_R,2)}')

calc_r+=1
ws.row_dimensions[calc_r].height=15
ref(ws,calc_r,1,"Avg MW metals (g/mol)")
AVGMW_R=calc_r
# Average MW of all 26 elements
mw_vals=",".join([str(MW[e]) for e in ELEMENTS])
calc(ws,calc_r,2,f'=AVERAGE({mw_vals})')

calc_r+=1
ws.row_dimensions[calc_r].height=15
ref(ws,calc_r,1,"Max loading (g metal/L org)")
MAXLOAD_R=calc_r
calc(ws,calc_r,2,f'={cr(AMINE_R,2)}*{cr(AVGMW_R,2)}')

calc_r+=1
ws.row_dimensions[calc_r].height=15
ref(ws,calc_r,1,"Actual loading (g/L)")
out(ws,calc_r,2,f'={cr(MAXLOAD_R,2)}*{cr(EXP_R,2)}/100')

# ══════════════════════════════════════════════════════════════════════════════
# RESIDENCE TIME
# ══════════════════════════════════════════════════════════════════════════════
RT_START=calc_r+3
ws.row_dimensions[RT_START].height=20
ws.merge_cells(f"A{RT_START}:N{RT_START}")
c=ws.cell(row=RT_START,column=1,value="RESIDENCE TIME ANALYSIS")
c.fill=fl(C["SECHDR"]);c.font=fn("FFFFFF",bold=True,sz=12);c.alignment=al()

rt_inps=[
    ("Mixer volume (m³)",0.001),
    ("Settler length (m)",1.0),
    ("Settler width (m)",0.5),
    ("Settler depth (m)",0.2),
    ("n settlers/stage",1),
]
RT_DATA=RT_START+1
RT_CELLS={}
for i,(lbl,val) in enumerate(rt_inps):
    r=RT_DATA+i
    ws.row_dimensions[r].height=15
    ref(ws,r,1,lbl); inp(ws,r,2,val)
    RT_CELLS[lbl]=cr(r,2)

MV=RT_CELLS["Mixer volume (m³)"]; SL=RT_CELLS["Settler length (m)"]
SW=RT_CELLS["Settler width (m)"]; SD=RT_CELLS["Settler depth (m)"]
NS=RT_CELLS["n settlers/stage"]

rt_hdr=RT_DATA+len(rt_inps)+1
ws.row_dimensions[rt_hdr].height=34
for i,lbl in enumerate(["Section","n stages","Q_aq\n(mL/min)","Q_tot\n(mL/min)",
                         "τ_mixer\n(min)","V_settler\n(mL)","τ_settler\n(min)",
                         "τ_stage\n(min)","τ_section\n(min)"]):
    hdr(ws,rt_hdr,1+i,lbl)

rt_secs=[
    ("Extraction",N_EXT_REF, f"={F_REF}*1000",       f"=({F_REF}+{O_REF})*1000"),
    ("Scrub",     N_SCR_REF, f"={S_REF}*{R_REF}*1000",f"=({S_REF}*{R_REF}+{O_REF})*1000"),
    ("Strip",     N_STR_REF, f"={S_REF}*1000",        f"=({S_REF}+{O_REF})*1000"),
]
RT_ROW_BASE=rt_hdr+1
for i,(sec,ns_ref,qaq,qtot) in enumerate(rt_secs):
    r=RT_ROW_BASE+i
    ws.row_dimensions[r].height=15
    sc(ws,r,1,C["STEP_BG"],C["STEP_FG"],val=sec)
    ref(ws,r,2,f'={ns_ref}')
    calc(ws,r,3,qaq); calc(ws,r,4,qtot)
    calc(ws,r,5,f'=({MV}*1000000)/{ad(r,4)}')
    calc(ws,r,6,f'={SL}*{SW}*{SD}*{NS}*1000000')
    calc(ws,r,7,f'={ad(r,6)}/{ad(r,4)}')
    calc(ws,r,8,f'={ad(r,5)}+{ad(r,7)}')
    calc(ws,r,9,f'={ad(r,2)}*{ad(r,8)}')

# Steady-state estimator
SS_START=RT_ROW_BASE+3+2
ws.row_dimensions[SS_START].height=18
ws.merge_cells(f"A{SS_START}:N{SS_START}")
hdr(ws,SS_START,1,"STEADY-STATE TIME ESTIMATOR  (Ritcey & Ashbrook 1979; Perry's 8th §15-42)",
    span=14,bg=C["SECHDR"])

ss=SS_START+1
ref(ws,ss,1,"Safety factor"); inp(ws,ss,2,3); SF=cr(ss,2)
ss+=1
tau_range=f'{ad(RT_ROW_BASE,9)}:{ad(RT_ROW_BASE+2,9)}'
ref(ws,ss,1,"τ_max (min)"); calc(ws,ss,2,f'=MAX({tau_range})'); TAU=cr(ss,2)
ss+=1
ref(ws,ss,1,"t_SS = factor × τ_max (min)"); out(ws,ss,2,f'={SF}*{TAU}')
ss+=1
sec_range=f'A{RT_ROW_BASE}:A{RT_ROW_BASE+2}'
ref(ws,ss,1,"Bottleneck section")
calc(ws,ss,2,f'=IFERROR(INDEX({sec_range},MATCH(MAX({tau_range}),{tau_range},0)),"—")')

# Feed variability
FV_START=ss+3
ws.row_dimensions[FV_START].height=18
ws.merge_cells(f"A{FV_START}:N{FV_START}")
hdr(ws,FV_START,1,"FEED VARIABILITY SCENARIO  (Zr × k1/k2)",
    span=14,bg=C["SECHDR"])

fv=FV_START+1
ref(ws,fv,1,"k1 (Zr × high)"); inp(ws,fv,2,1.20); K1=cr(fv,2)
fv+=1
ref(ws,fv,1,"k2 (Zr × low)");  inp(ws,fv,2,0.80); K2=cr(fv,2)
fv+=1
ZR_FEED=PC_CELLS["Zr"]["XF"]
for i,lbl in enumerate(["Metric","Normal","Zr×k1","Zr×k2"]):
    hdr(ws,fv,1+i,lbl)
fv+=1
ref(ws,fv,1,"Effective Zr feed (g/L)")
out(ws,fv,2,f'={ZR_FEED}')
out(ws,fv,3,f'={ZR_FEED}*{K1}')
out(ws,fv,4,f'={ZR_FEED}*{K2}')

# ══════════════════════════════════════════════════════════════════════════════
# ELEMENT DISTRIBUTION CHARTS — Bar charts (Feed / Raffinate / Pregnant)
# Three grouped bar charts showing concentration of all 26 elements per stream.
# Also a data table with element groups for reference.
# ══════════════════════════════════════════════════════════════════════════════

TREE_START = fv + 4
ws.row_dimensions[TREE_START].height = 18
ws.merge_cells(f"A{TREE_START}:{get_column_letter(N_EL+2)}{TREE_START}")
c = ws.cell(row=TREE_START, column=1,
            value="ELEMENT CONCENTRATION DATA — Feed / Raffinate / Pregnant  (g/L)")
c.fill = fl(C["SECHDR"]); c.font = fn("FFFFFF", bold=True, sz=11); c.alignment = al()

# Column headers
TM_HDR = TREE_START + 1
ws.row_dimensions[TM_HDR].height = 16
hdr(ws, TM_HDR, 1, "Element")
hdr(ws, TM_HDR, 2, "Feed (g/L)",      bg=C["EXT_T"], fg=C["SECHDR"])
hdr(ws, TM_HDR, 3, "Raffinate (g/L)", bg=C["SCR_T"], fg=C["SECHDR"])
hdr(ws, TM_HDR, 4, "Pregnant (g/L)",  bg=C["STR_T"], fg=C["SECHDR"])
hdr(ws, TM_HDR, 5, "Group")

TM_DATA = TM_HDR + 1
GROUPS = {
    "Al":"Light metals","Sc":"Transition","Fe":"Transition","Co":"Transition",
    "Zn":"Transition","Ga":"Post-transition","Rb":"Alkali","Y":"REE",
    "Zr":"HFSEs","La":"REE","Ce":"REE","Pr":"REE","Nd":"REE","Sm":"REE",
    "Eu":"REE","Gd":"REE","Tb":"REE","Dy":"REE","Ho":"REE","Er":"REE",
    "Tm":"REE","Yb":"REE","Lu":"REE","Hf":"HFSEs","Th":"Actinides","U":"Actinides",
}

for e_idx, elem in enumerate(ELEMENTS):
    r = TM_DATA + e_idx
    ws.row_dimensions[r].height = 14
    sc(ws, r, 1, C["STEP_BG"], C["STEP_FG"], val=elem)
    feed_ref = PC_CELLS[elem]["XF"]
    calc(ws, r, 2, f"={feed_ref}")
    ec = ELEM_COLS[elem]
    calc(ws, r, 3, f"={ad(STAGE_ROW_START, ec['x_aq'])}")
    xcol = get_column_letter(ec["x_aq"])
    xrange = f'{xcol}{STAGE_ROW_START}:{xcol}{STAGE_ROW_START+N_STAGES-1}'
    calc(ws, r, 4,
         f'=IFERROR(INDEX({xrange},{N_EXT_REF}+{N_SCR_REF}+{N_STR_REF}),0)')
    sc(ws, r, 5, C["REF_BG"], C["REF_FG"], val=GROUPS.get(elem,"Other"), h="left")

# ── Bar charts: Feed / Raffinate / Pregnant for all 26 elements ───────────────
from openpyxl.chart import BarChart
CHART_DEFS = [
    ("Feed Concentration — 26 Elements (g/L)",      2, "1B3A6B"),
    ("Raffinate Concentration — 26 Elements (g/L)", 3, "B71C1C"),
    ("Pregnant Solution — 26 Elements (g/L)",       4, "1B5E20"),
]
for i, (title, data_col, color) in enumerate(CHART_DEFS):
    chart = BarChart()
    chart.type = "col"
    chart.title = title
    chart.style = 10
    chart.y_axis.title = "Concentration (g/L)"
    chart.x_axis.title = "Element"
    chart.height = 14
    chart.width = 28

    data_ref = Reference(ws, min_col=data_col,
                         min_row=TM_HDR, max_row=TM_DATA + N_EL - 1)
    cats = Reference(ws, min_col=1,
                     min_row=TM_DATA, max_row=TM_DATA + N_EL - 1)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    chart.series[0].graphicalProperties.solidFill = color
    chart.series[0].graphicalProperties.line.solidFill = color

    anchor_row = TM_DATA + N_EL + 3 + i * 32
    ws.add_chart(chart, f"A{anchor_row}")

# ── Save ───────────────────────────────────────────────────────────────────────
OUT = "SX_Steady_State_Model_26elem.xlsx"
wb.save(OUT)
print(f"Saved: {OUT}")
