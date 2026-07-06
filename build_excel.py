"""
build_excel.py  —  SX_Steady_State_Model_26elem.xlsx
Primene JMT / dual-acid SX  |  H2SO4 loading, HNO3 stripping
26 elements: Al Sc Fe Co Zn Ga Rb Y Zr La Ce Pr Nd Sm Eu Gd Tb Dy Ho Er Tm Yb Lu Hf Th U
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
ws = wb.active
ws.title = "SX Model"
ws.sheet_properties.tabColor = C["TITLE"]

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

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 0 — REGRESSION  (rows 4+)
# Layout of each panel (Loading / Stripping):
#   col 1          : Stage
#   col 2          : [acid] N
#   col 3          : Flow_aq
#   col 4          : Flow_org
#   cols 5 to 5+N_EL-1  : C_aq per element
#   cols 5+N_EL to 5+2*N_EL-1 : C_org per element
# Panel width = 4 + 2*N_EL
# ══════════════════════════════════════════════════════════════════════════════
REG_START = 4
PANEL_W = 4 + 2*N_EL          # 4 + 52 = 56

LOAD_COL  = 1
STRIP_COL = LOAD_COL + PANEL_W + 1   # 58

ws.row_dimensions[REG_START].height=20
ws.merge_cells(f"A{REG_START}:{get_column_letter(STRIP_COL+PANEL_W)}{REG_START}")
c=ws.cell(row=REG_START,column=1,
          value="SECTION 0 — P/Q REGRESSION  (D = P × [acid]^Q  via log-log linear regression)")
c.fill=fl(C["SECHDR"]);c.font=fn("FFFFFF",bold=True,sz=11);c.alignment=al()

STAGE_LBLS_L = ["L1","L2","L3"]
STAGE_LBLS_S = ["S1","S2","S3"]
N_STAGES_REG = 3

def build_panel_header(ws, base_row, base_col, title, bg_title):
    """Build 3-row header for a regression input panel."""
    ws.row_dimensions[base_row].height=16
    hdr(ws, base_row, base_col, title, span=PANEL_W, bg=bg_title, fg=C["SECHDR"], sz=10)
    ws.row_dimensions[base_row+1].height=30
    hdr(ws,base_row+1,base_col,   "Stage")
    hdr(ws,base_row+1,base_col+1, "[acid]\nN")
    hdr(ws,base_row+1,base_col+2, "Flow_aq\n(mL/min)")
    hdr(ws,base_row+1,base_col+3, "Flow_org\n(mL/min)")
    for i,e in enumerate(ELEMENTS):
        hdr(ws,base_row+1,base_col+4+i,     f"{e}_aq\n(ppm)")
        hdr(ws,base_row+1,base_col+4+N_EL+i,f"{e}_org\n(ppm)")

build_panel_header(ws, REG_START+1, LOAD_COL,  "LOADING DATA  (H₂SO₄ system)", C["EXT_T"])
build_panel_header(ws, REG_START+1, STRIP_COL, "STRIPPING DATA  (HNO₃ system)", C["STR_T"])

# Input data rows
INP_ROW_BASE = REG_START + 3   # first data row = row 7
for i,(ll,sl) in enumerate(zip(STAGE_LBLS_L, STAGE_LBLS_S)):
    r = INP_ROW_BASE + i
    ws.row_dimensions[r].height=15
    sc(ws,r,LOAD_COL, C["STEP_BG"],C["STEP_FG"],val=ll)
    sc(ws,r,STRIP_COL,C["STEP_BG"],C["STEP_FG"],val=sl)
    for c_off in range(1, PANEL_W):
        inp(ws,r,LOAD_COL +c_off)
        inp(ws,r,STRIP_COL+c_off)

# ── Calculated D values ────────────────────────────────────────────────────────
D_HDR_ROW = INP_ROW_BASE + N_STAGES_REG + 1  # row 11
D_PANEL_W = 1 + N_EL   # stage + N_EL D cols

ws.row_dimensions[D_HDR_ROW].height=18
hdr(ws,D_HDR_ROW,LOAD_COL,  "CALCULATED D VALUES  (C_org/C_aq)", span=D_PANEL_W, bg=C["COLHDR"])
hdr(ws,D_HDR_ROW,STRIP_COL, "CALCULATED D VALUES  (C_org/C_aq)", span=D_PANEL_W, bg=C["COLHDR"])

D_COL_HDR = D_HDR_ROW+1
ws.row_dimensions[D_COL_HDR].height=16
hdr(ws,D_COL_HDR,LOAD_COL,"Stage")
hdr(ws,D_COL_HDR,STRIP_COL,"Stage")
for i,e in enumerate(ELEMENTS):
    hdr(ws,D_COL_HDR,LOAD_COL +1+i,f"D_{e}")
    hdr(ws,D_COL_HDR,STRIP_COL+1+i,f"D_{e}")

D_DATA_ROW = D_COL_HDR+1  # row 13
for i,(ll,sl) in enumerate(zip(STAGE_LBLS_L,STAGE_LBLS_S)):
    r=D_DATA_ROW+i
    ws.row_dimensions[r].height=15
    inp_r=INP_ROW_BASE+i
    sc(ws,r,LOAD_COL, C["STEP_BG"],C["STEP_FG"],val=ll)
    sc(ws,r,STRIP_COL,C["STEP_BG"],C["STEP_FG"],val=sl)
    for e_idx in range(N_EL):
        # Loading
        aq_c  = LOAD_COL+4+e_idx
        org_c = LOAD_COL+4+N_EL+e_idx
        aqa=ad(inp_r,aq_c); orga=ad(inp_r,org_c)
        calc(ws,r,LOAD_COL+1+e_idx,
             f'=IFERROR(IF(OR({aqa}="",{orga}="",{aqa}=0),"—",{orga}/{aqa}),"—")')
        # Stripping
        aq_c  = STRIP_COL+4+e_idx
        org_c = STRIP_COL+4+N_EL+e_idx
        aqa=ad(inp_r,aq_c); orga=ad(inp_r,org_c)
        calc(ws,r,STRIP_COL+1+e_idx,
             f'=IFERROR(IF(OR({aqa}="",{orga}="",{aqa}=0),"—",{orga}/{aqa}),"—")')

# ── Regression summary ─────────────────────────────────────────────────────────
REG_SUM_ROW = D_DATA_ROW + N_STAGES_REG + 1  # row 17
ws.row_dimensions[REG_SUM_ROW].height=18
hdr(ws,REG_SUM_ROW,1,
    "LOG-LOG REGRESSION SUMMARY  (D = P × [acid]^Q)  — OUTPUT cells feed TDMA Panel B",
    span=7*N_EL, bg=C["SECHDR"], sz=10)

RS_COL_HDR = REG_SUM_ROW+1
ws.row_dimensions[RS_COL_HDR].height=30
reg_labels=["Element","P_loading","Q_loading","R²_load","P_strip","Q_strip","R²_strip"]
for i,lbl in enumerate(reg_labels):
    hdr(ws,RS_COL_HDR,1+i,lbl)

# acid N column: LOAD col 2, rows INP_ROW_BASE to INP_ROW_BASE+2
load_acid_range = f"{ad(INP_ROW_BASE,LOAD_COL+1)}:{ad(INP_ROW_BASE+2,LOAD_COL+1)}"
strip_acid_range= f"{ad(INP_ROW_BASE,STRIP_COL+1)}:{ad(INP_ROW_BASE+2,STRIP_COL+1)}"

REG_CELLS = {}   # element -> dict of (row,col) for P_load,Q_load,P_strip,Q_strip
RS_DATA_ROW = RS_COL_HDR+1  # row 19

def rq(d_range, acid_range):
    return (f'=IFERROR(IF(COUNTA({d_range})<2,"Need≥2",'
            f'SLOPE(LN(IF(ISNUMBER({d_range}),{d_range},1)),'
            f'LN(IF(ISNUMBER({acid_range}),{acid_range},1)))),"N/A")')

def rp(d_range, acid_range):
    return (f'=IFERROR(IF(COUNTA({d_range})<2,"Need≥2",'
            f'EXP(INTERCEPT(LN(IF(ISNUMBER({d_range}),{d_range},1)),'
            f'LN(IF(ISNUMBER({acid_range}),{acid_range},1))))),"N/A")')

def rr2(d_range, acid_range):
    return (f'=IFERROR(IF(COUNTA({d_range})<2,"Need≥2",'
            f'RSQ(LN(IF(ISNUMBER({d_range}),{d_range},1)),'
            f'LN(IF(ISNUMBER({acid_range}),{acid_range},1)))),"N/A")')

for e_idx,elem in enumerate(ELEMENTS):
    r = RS_DATA_ROW + e_idx
    ws.row_dimensions[r].height=15
    sc(ws,r,1,C["STEP_BG"],C["STEP_FG"],val=elem)

    ld_col = LOAD_COL+1+e_idx   # D column in D table
    st_col = STRIP_COL+1+e_idx
    ld_d_range = f"{ad(D_DATA_ROW,ld_col)}:{ad(D_DATA_ROW+2,ld_col)}"
    st_d_range = f"{ad(D_DATA_ROW,st_col)}:{ad(D_DATA_ROW+2,st_col)}"

    out(ws,r,2, rp(ld_d_range, load_acid_range))
    out(ws,r,3, rq(ld_d_range, load_acid_range))
    calc(ws,r,4,rr2(ld_d_range, load_acid_range))
    out(ws,r,5, rp(st_d_range, strip_acid_range))
    out(ws,r,6, rq(st_d_range, strip_acid_range))
    calc(ws,r,7,rr2(st_d_range, strip_acid_range))

    REG_CELLS[elem]={"P_load":(r,2),"Q_load":(r,3),
                     "P_strip":(r,5),"Q_strip":(r,6)}

# ══════════════════════════════════════════════════════════════════════════════
# TDMA INPUT PANELS  (start after regression + spacer)
# ══════════════════════════════════════════════════════════════════════════════
TDMA_START = RS_DATA_ROW + N_EL + 2   # row ~47

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
        ref_=cr(rr,cc)
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

# ── Save ───────────────────────────────────────────────────────────────────────
OUT="SX_Steady_State_Model_26elem.xlsx"
wb.save(OUT)
print(f"Saved: {OUT}")
