"""Builds template/OAR - Tilt Panel Design.xlsx (OAR house calc style).

Run from the repo root:  python template/build_panel_template.py
One calc tab = one tilt panel (or panel zone). The analysis is done in any FE package: the second-order design
actions (per metre, or totals for an isolated leg / head of width b: N*, M*y in the span and at the support,
M*x) for each ULS combination and the SLS (working)
N and M are entered on the sheet, and screenshots of the plate results are pasted into the frames on the right
(printed as extra pages with the calc).

The sheet then does the AS 3600-2018 section checks on b × tw (b = 1000 for unit width):
  - vertical capacity at N* by strain compatibility (OAR.CONCRETE.COLUMN), M*y span (≥ 0.05 tw N*) and at the
    support, horizontal M*x against the horizontal bars (OAR.CONCRETE.ULT); moments are taken as second order
    (no moment magnifier);
  - SLS cracked bar stresses at the working moments, inward and outward (OAR.CONCRETE.CRACKED);
  - wall reinforcement minimums and spacing (Cl 11.7).
In-service only (lifting and bracing by the contractor's lifting designer). No new engine functions.
"""
import sys, os, re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.formula import ArrayFormula

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from oar_style import *  # house-style helpers
from oar_style import BARS, BAR_NOTES

LOGO = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "assets", "oar-logo.png")
OUT = sys.argv[2] if len(sys.argv) > 2 else os.path.join(HERE, "OAR - Tilt Panel Design.xlsx")
_fsproj = open(os.path.join(HERE, "..", "src", "OAR.Concrete.Excel", "OAR.Concrete.Excel.fsproj"), encoding="utf-8").read()
TEMPLATE_VERSION = re.search(r"<Version>([0-9]+\.[0-9]+\.[0-9]+)</Version>", _fsproj).group(1)
_maj, _min, _pat = (int(v) for v in TEMPLATE_VERSION.split("."))

wb = Workbook()
wb.remove(wb.active)

# =========================== DATA TAB: Bars ===========================
D_ = wb.create_sheet("Bars")
setup_data_tab(D_, "Reinforcement – bar properties",
               "AS/NZS 4671 nominal values. S = 250 MPa deformed (D250N), R = plain round, L = mesh wire (Class L: SL92 = L9-200, SL82 = L8-200, "
               "SL102 = L10-200). Edit or add rows; the calc tabs read this list.",
               {"B": 10, "C": 10, "D": 10, "E": 10, "F": 10, "G": 10, "H": 44})
for j, t in enumerate(["Bar", "db (mm)", "Ab (mm²)", "fsy (MPa)", "Class", "Es (MPa)", "Note"]):
    c = D_.cell(4, 2 + j); c.value = t; c.font = font(8, True); c.fill = fill(HDR_FILL); c.border = BOX; c.alignment = CEN
P_BARS = [r + [200000] for r in BARS]
BAR0 = 5; BAR_LAST = BAR0 + len(P_BARS) + 5
for i in range(BAR_LAST - BAR0 + 1):
    row = P_BARS[i] + [BAR_NOTES.get(P_BARS[i][0], "")] if i < len(P_BARS) else [None] * 7
    for j, v in enumerate(row):
        c = D_.cell(BAR0 + i, 2 + j); c.value = v; c.font = font(8, color="000000"); c.fill = fill(IN_FILL); c.border = BOX
        c.alignment = LEFT if j in (0, 6) else CEN
EC_C = 11
for j, t in enumerate(["f'c (MPa)", "Ec (MPa)"]):
    c = D_.cell(4, EC_C + j); c.value = t; c.font = font(8, True); c.fill = fill(HDR_FILL); c.border = BOX; c.alignment = CEN
for i, (fc_, ec_) in enumerate(zip([20, 25, 32, 40, 50, 65, 80, 100], [24000, 26700, 30100, 32800, 34800, 37400, 39600, 42200])):
    for j, v in enumerate((fc_, ec_)):
        c = D_.cell(5 + i, EC_C + j); c.value = v; c.font = font(8, color="000000"); c.fill = fill(IN_FILL); c.border = BOX; c.alignment = CEN
c = D_.cell(13, EC_C); c.value = "AS 3600 Table 3.1.2 (ρ = 2400 kg/m³); calc tabs interpolate linearly."; c.font = font(7, i=True, color=GREY)
for k_, w_ in (("I", 3), ("J", 3), ("K", 10), ("L", 10)): D_.column_dimensions[k_].width = w_
wb.defined_names["BarList"] = DefinedName("BarList", attr_text=f"Bars!$B${BAR0}:$B${BAR_LAST}")
wb.defined_names["EcFc"] = DefinedName("EcFc", attr_text="Bars!$K$5:$K$12")
wb.defined_names["EcVal"] = DefinedName("EcVal", attr_text="Bars!$L$5:$L$12")
BARTAB = f"Bars!$B${BAR0}:$G${BAR_LAST}"
def barv(ref, col): return f"VLOOKUP({ref},{BARTAB},{col},FALSE)"

NLC = 6
XPOS = ("Both faces / with mesh", "Inside face", "Outside face")
STRIP = ("Unit width (per metre)", "Isolated leg / head")
COLS = "EFGHIJ"


def shot_frame(ws, r0, r1, title, hint):
    """Screenshot frame in M:T (printed as its own page)."""
    c = ws[f"M{r0 - 2}"]; c.value = title; c.font = font(10, True, color=GREEN); c.alignment = Alignment(vertical="bottom")
    for x in ws[f"M{r0 - 2}:T{r0 - 2}"][0]: x.border = UNDER
    dash = Side(style="dashed", color=RULE)
    for row in ws[f"M{r0}:T{r1}"]:
        for x in row:
            x.fill = fill("FAF8F6")
            x.border = Border(left=dash if x.column == 13 else None, right=dash if x.column == 20 else None,
                              top=dash if x.row == r0 else None, bottom=dash if x.row == r1 else None)
    c = ws[f"M{r0}"]; c.value = hint + "\n(right-click > Paste as Picture, then resize to fit this frame)"
    c.font = font(9, i=True, color="A6A6A6"); c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.merge_cells(f"M{r0}:T{r1}")


def build_tab(name, d):
    ws = wb.create_sheet(name)
    N = {}
    pend = []
    arrays = []
    def nm(n, ref):
        col = "".join(ch for ch in ref if ch.isalpha()); row = "".join(ch for ch in ref if ch.isdigit())
        N[n] = f"${col}${row}"
    def Fx(s):
        if not isinstance(s, str): return s
        s = s.replace("«TWO»", '«lay»="Two layers"').replace("«ISO»", f'«strip»="{STRIP[1]}"')
        for k in sorted(N, key=len, reverse=True): s = s.replace(f"«{k}»", N[k])
        assert "«" not in s, s
        return s
    def lab(ref, t, **k): return label(ws, ref, t, **k)
    def F(ref, f): pend.append((ref, f))
    def IN(ref, v=None, nf=None, al=CEN, name=None):
        c = inp(ws, ref, v, nf, al)
        if name: nm(name, ref.split(":")[0])
        return c
    def CA(ref, f, nf=None, al=CEN, name=None):
        c = calc(ws, ref, None, nf, al)
        if f is not None: pend.append((ref.split(":")[0], f))
        if name: nm(name, ref.split(":")[0])
        return c
    def U(ref, t): return unit(ws, ref, t)
    def sym(r, t):
        ws.merge_cells(f"B{r}:C{r}"); lab(f"B{r}", t, al=CEN)
    def desc(r, t, col="D"):
        if isinstance(t, str) and t.startswith("="): pend.append((f"{col}{r}", t))
        else: ws[f"{col}{r}"].value = t
        ws[f"{col}{r}"].font = font(8, color=GREY); ws[f"{col}{r}"].alignment = LEFT
    def small(ref, t): lab(ref, t, sz=7, color=GREY)
    def case_hdr(r):
        table_header(ws, r, [(f"{c}{r}", "") for c in COLS])
        for i, c in enumerate(COLS): F(f"{c}{r}", f"=$B${LC0 + i}")

    setup_calc_tab(ws, LOGO, '="TILT PANEL (AS 3600-2018)  –  "&IF(G3="","(ref)",G3)&IF(I3="","","  ·  "&I3)', "PANEL", name, "")
    ws["I3"].value = d["loc"]
    for col, w in zip("LMNOPQRST", (4, 8, 7, 9, 8, 9, 26, 6.5, 6.5)): ws.column_dimensions[col].width = w
    shot_frame(ws, 8, 50, "Analysis – ULS plate moments (second order, my and mx) and N", "Paste the ULS plate moment screenshots here")
    shot_frame(ws, 54, 96, "Analysis – SLS (working) moments", "Paste the SLS moment screenshots here")

    # ---- add-in version check (V3:V7, not printed)
    ws.column_dimensions["U"].width = 3; ws.column_dimensions["V"].width = 12
    lab("V2", "Add-in check", sz=7, color=GREY, i=True)
    for rr, t, v in ((3, "template", TEMPLATE_VERSION), (4, "add-in", '=IFERROR(SUBSTITUTE(OAR.CONCRETE.VERSION(),"OAR Concrete ",""),"")'),
                     (6, "template no.", _maj * 10000 + _min * 100 + _pat)):
        lab(f"W{rr}", t, sz=7, color=GREY); c = ws[f"V{rr}"]; c.value = v; c.font = font(7, color=GREY); c.alignment = LEFT
    p1 = 'FIND(".",$V$4)'; p2 = f'FIND(".",$V$4,{p1}+1)'
    ws["V5"].value = f'=IFERROR(LEFT($V$4,{p1}-1)*10000+MID($V$4,{p1}+1,{p2}-{p1}-1)*100+MID($V$4,{p2}+1,9),0)'
    ws["V7"].value = '=AND($V$4<>"",$V$5<$V$6)'
    lab("W5", "add-in no.", sz=7, color=GREY); lab("W7", "add-in older?", sz=7, color=GREY)
    for rr in (5, 7): ws[f"V{rr}"].font = font(7, color=GREY); ws[f"V{rr}"].alignment = LEFT
    ws["B7"].value = ('=IF($V$4="","⚠ OAR Concrete add-in not loaded – results below are not calculated",IF($V$7,"⚠ Add-in "&$V$4&" is older than this template ("&$V$3&") – update the add-in and restart Excel",'
                      '"Copy this tab for additional panels.   ·   Calculated with OAR Concrete "&$V$4))')
    ws.conditional_formatting.add("B7", FormulaRule(formula=['OR($V$7,$V$4="")'], font=Font(color=RED, bold=True, italic=False)))

    # ================================================================== INPUTS
    heading(ws, 33, "Panel and Concrete")
    lab("B34", "Panel thickness tw"); IN("E34", d["tw"], "0", name="tw"); U("F34", "mm")
    lab("B35", "Concrete strength f'c"); IN("E35", d["fc"], "0", name="fc"); U("F35", "MPa")
    lab("B36", "Stress block (ULS)"); IN("E36:F36", d.get("model", "Rectangular"), name="model")
    lab("B37", "Ec override"); IN("E37", None, "0", name="ecin"); U("F37", "MPa"); small("H37", "blank = Table 3.1.2")
    lab("B38", "φ0 for compression, Table 2.2.2"); IN("E38", d.get("phi0", 0.6), "0.00", name="phi0"); small("H38", "Amdt 2: φ0 = 0.65 kφ – enter the value")
    lab("B39", "Design strip | width b"); IN("E39:F39", d.get("strip", STRIP[0]), name="strip"); IN("G39", d.get("bw", 1000), "0", name="bw"); U("H39", "mm")
    small("I39", "unit width: b = 1000")
    F("B40", '=IF(«ISO»,"Isolated leg / head: actions are strip totals over b, extra vertical bars as a number. ","Unit width: actions per metre, extras at spacing. ")&"ULS moments are second order from the analysis – no magnifier."')
    ws["B40"].font = font(7, i=True, color=GREY)

    heading(ws, 41, "Reinforcement")
    lab("B42", "Layout"); IN("E42:F42", d.get("lay", "Central"), name="lay")
    lab("B43", "Cover to vertical bars (two layers)"); IN("E43", d.get("cover", 40), "0", name="cover"); U("F43", "mm")
    lab("B44", "Vertical bars: bar | spacing (each layer)"); IN("E44", d["vbar"], name="vbar"); IN("F44", d["vs"], "0", name="vs"); U("G44", "mm")
    lab("B45", "Horizontal bars: bar | spacing (each layer)"); IN("E45", d["hbar"], name="hbar"); IN("F45", d["hs"], "0", name="hs"); U("G45", "mm")
    F("B46", '=IF(«ISO»,"Extra vertical bars: bar | number in b | face","Extra vertical bars: bar | spacing | face")'); ws["B46"].font = font(9)
    IN("E46", d.get("xbar", "None"), name="xbar"); IN("F46", d.get("xs"), "0", name="xs"); U("G46", None); F("G46", '=IF(«ISO»,"no.","mm")')
    IN("H46:J46", d.get("xpos", XPOS[0]), name="xpos")
    lab("B47", "Extra horizontal bars: bar | spacing | face"); IN("E47", d.get("hxbar", "None"), name="hxbar"); IN("F47", d.get("hxs"), "0", name="hxs"); U("G47", "mm")
    IN("H47:J47", d.get("hxpos", XPOS[0]), name="hxpos")
    lab("B48", "Restrained against horizontal shrinkage"); IN("E48", d.get("hrest", "No"), name="hrest"); small("H48", "No = free to move (panel joints)")
    note(ws, 49, "Extras face: central mesh – tied against the inside or outside of the mesh (+M d up, −M d down); two layers – that layer only. Both = with the mesh / each layer.")

    heading(ws, 52, "ULS Design Actions (second order)"); F("B52", '=IF(«ISO»,"ULS Design Actions (second order, strip total over b)","ULS Design Actions (second order, per metre)")')
    note(ws, 53, "N* + compression. M*y = vertical bending (+ inside face in tension): in the span, and at the support (e.g. parapet or bracket). M*x = horizontal bending.")
    table_header(ws, 54, [("B54:C54", "Case"), ("D54", "Description"), ("E54", "N*"), ("F54", "M*y span"), ("G54", "M*y support"),
                          ("H54", "M*x (kNm/m)")])
    F("E54", '=IF(«ISO»,"N* (kN)","N* (kN/m)")')
    LC0 = 55
    for i in range(NLC):
        r = LC0 + i; lc = d["lcs"][i] if i < len(d["lcs"]) else None
        IN(f"B{r}:C{r}", f"LC{i + 1}")
        IN(f"D{r}", lc[0] if lc else None, al=LEFT)
        for j, col in enumerate("EFGH"):
            IN(f"{col}{r}", lc[1 + j] if lc else None, "0.00")
    note(ws, 61, "Blank N* = case not used. Enter the peak second-order moments; M*y span is taken as at least 0.05 tw N* (Cl 10.1.2).")

    heading(ws, 62, "SLS (working) Actions"); F("B62", '=IF(«ISO»,"SLS (working) Actions, strip total over b","SLS (working) Actions, per metre")')
    lab("B63", "Ns, SLS combination (compression +)"); IN("E63", d.get("Ns"), "0.0", name="Ns"); U("F63", None); F("F63", '=IF(«ISO»,"kN","kN/m")')
    lab("B64", "Ms+ inward (inside face in tension)"); IN("E64", d.get("Msp"), "0.00", name="Msp"); U("F64", None); F("F64", '=IF(«ISO»,"kNm","kNm/m")')
    lab("B65", "Ms− outward (outside face in tension)"); IN("E65", d.get("Msn"), "0.00", name="Msn"); U("F65", None); F("F65", '=IF(«ISO»,"kNm","kNm/m")'); small("H65", "magnitude")
    lab("B66", "Bar stress limit σs (optional)"); IN("E66", d.get("slim"), "0", name="slim"); U("F66", "MPa"); small("H66", "blank = report only, e.g. 0.8 fsy")
    note(ws, 67, "Cracked elastic section, no concrete tension, n = Es / Ec (short term), at Ns and Ms. Ns < 0 (net tension) is taken as 0.")


    # ================================================================== CALCULATIONS
    heading(ws, 70, "Section Properties"); F("B70", '=IF(«ISO»,"Section Properties (strip b × tw)","Section Properties (per metre)")')
    def prop(rr, s_, ds, f, nf, u, name=None):
        sym(rr, s_); desc(rr, ds); CA(f"E{rr}", f, nf, name=name); U(f"G{rr}", u)
    prop(71, "b × tw", "section (add-in dims)", "=IF(«ISO»,«bw»,1000)", "0", "mm", name="dimb"); CA("F71", "=«tw»", "0", name="dimt")
    prop(72, "Ag", "b tw", "=«dimb»*«tw»", "0", "mm²", name="Ag")
    I_ = "MIN(MATCH(«fc»,EcFc,1),7)"
    prop(73, "Ec", '=IF(N(«ecin»)>0,"entered","Table 3.1.2, interpolated")',
         f"=IF(N(«ecin»)>0,«ecin»,INDEX(EcVal,{I_})+(«fc»-INDEX(EcFc,{I_}))*(INDEX(EcVal,{I_}+1)-INDEX(EcVal,{I_}))/(INDEX(EcFc,{I_}+1)-INDEX(EcFc,{I_})))",
         "0", "MPa", name="Ec")
    prop(74, "d", '=IF(«TWO»,"tw − cover − db/2 (inside layer)","tw / 2 (central layer)")',
         f'=IF(«TWO»,«tw»-«cover»-IFERROR({barv("«vbar»", 2)},0)/2,«tw»/2)', "0.0", "mm", name="dd")
    prop(75, "ρv", "vertical, main bars all layers / (1000 tw)", f'=IF(«TWO»,2,1)*IFERROR({barv("«vbar»", 3)},0)/(«vs»*«tw»)', "0.000%", "", name="rhov")
    prop(76, "ρh", "horizontal, main bars all layers / (1000 tw)", f'=IF(«TWO»,2,1)*IFERROR({barv("«hbar»", 3)},0)/(«hs»*«tw»)', "0.000%", "", name="rhoh")

    SB0 = 250; BB0 = 256    # add-in bar tables (below the print area)
    BARS_V = f"$L${SB0}:$P${SB0 + 3}"; CLS_V = f"$Q${SB0}:$Q${SB0 + 3}"
    BARS_H = f"$L${BB0}:$P${BB0 + 3}"; CLS_H = f"$Q${BB0}:$Q${BB0 + 3}"
    SEC = '"Rect",«dimb»:«dimt»'
    SEC_H = f'"Rect",$L${BB0 + 5}:$M${BB0 + 5}'   # horizontal bending: always per metre of height
    MOD = '«fc»,«model»'
    def idx(expr, k): return f'IFERROR(INDEX({expr},{k},1),{expr})'

    # ---- design moments (second order from the analysis; minimum moment Cl 10.1.2)
    heading(ws, 81, "Design Moments – vertical (second order)")
    note(ws, 82, "M*d = max(|M*y span|, 0.05 tw N*), sign of M*y span (Cl 10.1.2 minimum). Moments from the analysis are second order: no magnifier.")
    table_header(ws, 83, [("B83:C83", "Case"), ("D83", "N* (kN)"), ("E83", "M*y (kNm)"), ("F83", "0.05 tw N*"), ("G83", "M*d (kNm)")])
    MAG0 = 84
    for i in range(NLC):
        gr = MAG0 + i; lr = LC0 + i; blank = f'$E${lr}=""'
        ws.merge_cells(f"B{gr}:C{gr}"); F(f"B{gr}", f"=$B${lr}"); ws[f"B{gr}"].font = font(8); ws[f"B{gr}"].alignment = CEN
        CA(f"D{gr}", f'=IF({blank},"",$E${lr})', "0.0")
        CA(f"E{gr}", f'=IF({blank},"",N($F${lr}))', "0.00")
        CA(f"F{gr}", f'=IF({blank},"",MAX(0,0.05*«tw»/1000*D{gr}))', "0.00")
        CA(f"G{gr}", f'=IF({blank},"",IF(E{gr}<0,-1,1)*MAX(ABS(E{gr}),F{gr}))', "0.00")

    # ---- vertical capacity
    CP = 91
    heading(ws, CP, "Vertical Strength – strain compatibility at N* (OAR Concrete add-in)")
    note(ws, CP + 1, "OAR.CONCRETE.COLUMN on b × tw at N*, M*d: φ from Tbl 2.2.2 (φ0 … φb by N*/Nub; any Class L bar → φb 0.65).")
    case_hdr(CP + 2)
    CPR = [("φ", "Tbl 2.2.2 at N*", "0.000", ""), ("Mu", "capacity", "0.00", "kNm"), ("φMu", "", "0.00", "kNm"),
           ("Mux", "", "0.00", "kNm"), ("Muy", "", "0.00", "kNm"), ("θ", "compression direction", "0", "deg"),
           ("dn", "neutral axis depth", "0.0", "mm"), ("Nub", "balanced", "0", "kN"), ("Nuo", "squash", "0", "kN"),
           ("εs", "outermost tension bar strain", "0.00%", "")]
    C0 = CP + 3
    for kk, (s_, ds, nf, u) in enumerate(CPR):
        rr = C0 + kk; sym(rr, s_); desc(rr, ds)
        for c in COLS: calc(ws, f"{c}{rr}", None, nf)
        U(f"K{rr}", u)
    C1 = C0 + len(CPR) - 1
    for i, c in enumerate(COLS):
        lr = LC0 + i; gr = MAG0 + i
        F(f"{c}{C0}", f'=IF($E${lr}="","",OAR.CONCRETE.COLUMN({SEC},{BARS_V},{MOD},$D${gr},$G${gr},0,«phi0»,{CLS_V},«Ec»))')
        arrays.append(f"{c}{C0}:{c}{C1}")
    RM = C1 + 1; sym(RM, "|M*d|"); desc(RM, "design moment"); U(f"K{RM}", "kNm")
    for i, c in enumerate(COLS): CA(f"{c}{RM}", f'=IF($E${LC0 + i}="","",ABS($G${MAG0 + i}))', "0.00")
    RSC = RM + 1; sym(RSC, "ratio"); desc(RSC, "|M*d| / φMu")
    for i, c in enumerate(COLS): CA(f"{c}{RSC}", f'=IF($E${LC0 + i}="","",{c}{RM}/{c}{C0 + 2})', "0.00")

    # ---- top support
    TS = RSC + 2
    heading(ws, TS, "Support – M*y at the support")
    case_hdr(TS + 1)
    T0 = TS + 2
    for kk, (s_, ds, nf, u) in enumerate([("N*", "", "0.0", "kN"), ("M*y sup", "", "0.00", "kNm"),
                                          ("φMu", "capacity at N*", "0.00", "kNm"), ("ratio", "|M*y sup| / φMu", "0.00", "")]):
        rr = T0 + kk; sym(rr, s_); desc(rr, ds); U(f"K{rr}", u)
        for i, c in enumerate(COLS):
            lr = LC0 + i; blank = f'$E${lr}=""'
            colf = f"OAR.CONCRETE.COLUMN({SEC},{BARS_V},{MOD},{c}{T0},{c}{T0 + 1},0,«phi0»,{CLS_V},«Ec»)"
            f = {0: f'=IF({blank},"",$E${lr})', 1: f'=IF({blank},"",N($G${lr}))',
                 2: f'=IF(OR({blank},N({c}{T0 + 1})=0),"",{idx(colf, 3)})',
                 3: f'=IF({c}{T0 + 2}="","",ABS({c}{T0 + 1})/{c}{T0 + 2})'}[kk]
            CA(f"{c}{rr}", f, nf)
    TR = T0 + 3

    # ---- horizontal bending
    HZ = TR + 2
    heading(ws, HZ, "Horizontal Bending – M*x against the horizontal bars (OAR Concrete add-in)")
    note(ws, HZ + 1, "OAR.CONCRETE.ULT per metre of height, horizontal bars (main + extras), N = 0.")
    r = HZ + 2; table_header(ws, r, [(f"E{r}", "+ (inside tension)"), (f"F{r}", "− (outside tension)")])
    r += 1; sym(r, "Muo"); desc(r, "OAR.CONCRETE.ULT, N = 0"); r_muo = r
    ult = {s_: f'OAR.CONCRETE.ULT({SEC_H},{BARS_H},{MOD},"{s_}",0,«Ec»,{CLS_H},"Whole section 0.65")' for s_ in ("Sagging", "Hogging")}
    for col, sense in (("E", "Sagging"), ("F", "Hogging")): CA(f"{col}{r}", "=" + idx(ult[sense], 13), "0.00")
    U(f"G{r}", "kNm")
    r += 1; sym(r, "φ"); desc(r, "Tbl 2.2.2 (Class L in tension → 0.65)")
    for col, sense in (("E", "Sagging"), ("F", "Hogging")): CA(f"{col}{r}", "=" + idx(ult[sense], 14), "0.000")
    r += 1; sym(r, "φMuo"); desc(r, "")
    for col in "EF": CA(f"{col}{r}", f'={col}{r_muo}*{col}{r - 1}', "0.00")
    U(f"G{r}", "kNm"); nm("phiMxp", f"E{r}"); nm("phiMxn", f"F{r}")
    r += 1; case_hdr(r)
    X0 = r + 1
    sym(X0, "M*x"); desc(X0, "from the analysis"); U(f"K{X0}", "kNm")
    sym(X0 + 1, "ratio"); desc(X0 + 1, "|M*x| / φMuo")
    for i, c in enumerate(COLS):
        lr = LC0 + i
        CA(f"{c}{X0}", f'=IF(OR($E${lr}="",$H${lr}=""),"",$H${lr})', "0.00")
        CA(f"{c}{X0 + 1}", f'=IF({c}{X0}="","",ABS({c}{X0})/IF({c}{X0}>=0,«phiMxp»,«phiMxn»))', "0.00")
    XR = X0 + 1

    # ---- SLS cracked bar stresses
    SD = XR + 2
    heading(ws, SD, "Serviceability – cracked bar stresses at the working moment")
    note(ws, SD + 1, "OAR.CONCRETE.CRACKED on b × tw at Ns (≥ 0) and Ms: transformed section, n = Es / Ec, no concrete tension.")
    r = SD + 2; table_header(ws, r, [(f"E{r}", "Ms+ inward"), (f"F{r}", "Ms− outward")])
    r += 1; sym(r, "Ms"); desc(r, "working moment"); U(f"G{r}", "kNm")
    CA(f"E{r}", '=IF(N(«Msp»)=0,"",ABS(«Msp»))', "0.00"); CA(f"F{r}", '=IF(N(«Msn»)=0,"",ABS(«Msn»))', "0.00"); RMS = r
    crk = {c: f'OAR.CONCRETE.CRACKED({SEC},{BARS_V},«Ec»,{c}{RMS},"{sense}",MAX(0,N(«Ns»)))' for c, sense in (("E", "Sagging"), ("F", "Hogging"))}
    for k_, (s_, ds, nf, u) in enumerate([("dn", "neutral axis depth", "0.0", "mm"), ("σst", "largest bar tension stress", "0", "MPa"),
                                          ("σc", "extreme fibre compression", "0.0", "MPa")]):
        r += 1; sym(r, s_); desc(r, ds); U(f"G{r}", u)
        for c in "EF": CA(f"{c}{r}", f'=IF({c}{RMS}="","",{idx(crk[c], (1, 3, 5)[k_])})', nf)
        if s_ == "σst": RST = r
    LAST = r + 2
    note(ws, LAST, "Not covered: shear, slenderness (moments entered as second order), lifting and handling, temporary bracing, in-plane lintel design, dowel / cleat capacity, FRL (AS 3600 Sec 5).")

    # ---- add-in bar tables (not printed)
    lab(f"L{SB0 - 2}", "Add-in bar tables on b (vertical) / per metre (horizontal), not printed: As | depth from outside face | fsy | x | Es | class", sz=7, color=GREY, i=True)
    lab(f"S{SB0}", "vertical: main outside/central | main inside | extra outside/central | extra inside", sz=7, color=GREY)
    lab(f"S{BB0}", "horizontal: same order", sz=7, color=GREY)
    def bar_rows(r0, bar, spacing, xbarref, xspacing, d_out, d_in, xd_out, xd_in, xpos, width="1000", xcount="FALSE"):
        ab = f'IFERROR({barv(bar, 3)},0)'; ax_ = f'IFERROR({barv(xbarref, 3)},0)'
        xa = f'IF({xcount},{xspacing}*{ax_},{ax_}*1000/{xspacing})'
        # extra outside/central row: two layers → outside layer unless "inside face"; central → always (depth set by face)
        rows = [(f'=IF(N({spacing})>0,{ab}*{width}/{spacing},0)', d_out, bar),
                (f'=IF(AND(«TWO»,N({spacing})>0),{ab}*{width}/{spacing},0)', d_in, bar),
                (f'=IF(AND(N({xspacing})>0,OR(NOT(«TWO»),{xpos}<>"{XPOS[1]}")),{xa},0)', xd_out, xbarref),
                (f'=IF(AND(«TWO»,N({xspacing})>0,{xpos}<>"{XPOS[2]}"),{xa},0)', xd_in, xbarref)]
        for k_, (fa, fd, bref) in enumerate(rows):
            rr = r0 + k_
            for col, f in (("L", fa), ("M", fd), ("N", f'=IFERROR({barv(bref, 4)},500)'), ("P", f'=IFERROR({barv(bref, 6)},200000)'),
                           ("Q", f'=IFERROR({barv(bref, 5)},"N")')):
                F(f"{col}{rr}", f); ws[f"{col}{rr}"].font = font(7, color=GREY)
    dbv = f'IFERROR({barv("«vbar»", 2)},0)'; dbh = f'IFERROR({barv("«hbar»", 2)},0)'
    dbx = f'IFERROR({barv("«xbar»", 2)},0)'; dbhx = f'IFERROR({barv("«hxbar»", 2)},0)'
    # central mesh: extras tied against it on the inside (deeper from the outside face) or the outside
    def off(xpos, dbm, dbe): return f'IF({xpos}="{XPOS[1]}",1,IF({xpos}="{XPOS[2]}",-1,0))*({dbm}+{dbe})/2'
    bar_rows(SB0, "«vbar»", "«vs»", "«xbar»", "«xs»",
             f"=IF(«TWO»,«cover»+{dbv}/2,«tw»/2)", f"=«tw»-«cover»-{dbv}/2",
             f"=IF(«TWO»,«cover»+{dbx}/2,«tw»/2+{off('«xpos»', dbv, dbx)})", f"=«tw»-«cover»-{dbx}/2", "«xpos»", width="«dimb»", xcount="«ISO»")
    bar_rows(BB0, "«hbar»", "«hs»", "«hxbar»", "«hxs»",
             f"=IF(«TWO»,«cover»+{dbv}+{dbh}/2,«tw»/2)", f"=«tw»-«cover»-{dbv}-{dbh}/2",
             f"=IF(«TWO»,«cover»+{dbv}+{dbhx}/2,«tw»/2+{off('«hxpos»', dbh, dbhx)})", f"=«tw»-«cover»-{dbv}-{dbhx}/2", "«hxpos»")

    for col, f in (("L", "=1000"), ("M", "=«tw»")): F(f"{col}{BB0 + 5}", f); ws[f"{col}{BB0 + 5}"].font = font(7, color=GREY)
    lab(f"N{BB0 + 5}", "horizontal section dims (per metre of height)", sz=7, color=GREY)

    # ================================================================== SUMMARY + CHECKS
    heading(ws, 8, "Design Summary")
    lab("B9", "Adopted panel", sz=7, color=GREY); lab("F9", "Max utilisation", sz=7, color=GREY); lab("I9", "Status", sz=7, color=GREY)
    ws.merge_cells("B10:E10"); ws.merge_cells("F10:G10"); ws.merge_cells("I10:K10")
    reo = ('IF(«TWO»,"2 layers ","central ")&«vbar»&"-"&«vs»&IF(AND(N(«xs»)>0,«xbar»<>"None")," + "&IF(«ISO»,«xs»&«xbar»,«xbar»&"-"&«xs»)&" extra"&IF(«xpos»="Both faces / with mesh",""," ("&IF(«xpos»="Inside face","inside","outside")&")"),"")&IF(«TWO»," EF","")')
    F("B10", f'=«tw»&" thk  ·  "&IF(«ISO»,"b "&«bw»&"  ·  ","")&{reo}')
    ws["B10"].font = font(12, True, color=GREEN); ws["B10"].alignment = LEFT
    CHK0, CHK1 = 15, 29
    c = ws["F10"]; c.value = f'=IF(COUNT(I{CHK0}:I{CHK1})=0,"–",MAX(I{CHK0}:I{CHK1}))'; c.font = font(12, True, color=GREEN); c.number_format = "0%"; c.alignment = LEFT
    ADD = [f"E{C0}:J{C0}", f"E{T0 + 2}:J{T0 + 2}", f"E{r_muo}:F{r_muo + 1}", f"E{RST - 1}:F{RST + 1}"]
    errs = "+".join(f'COUNTIF({a},"#OAR*")' for a in ADD)
    c = ws["I10"]
    pend.append(("I10", f'=IF($V$7,"UPDATE ADD-IN",IF({errs}>0,"INPUT ERROR",IF(COUNT(I{CHK0}:I{CHK1})=0,"PENDING",'
                        f'IF(OR(MAX(I{CHK0}:I{CHK1})>1,COUNTIF(J{CHK0}:J{CHK1},"NG")>0),"FAIL","PASS"))))'))
    c.font = font(10, True, color=GREY); c.alignment = CEN
    for x in ws["I10:K10"][0]: x.border = BOX
    passfail_cf(ws, "I10:K10", "$I$10")
    ws.conditional_formatting.add("I10:K10", FormulaRule(formula=['$I$10="UPDATE ADD-IN"'], fill=PatternFill(bgColor="C55A11"), font=Font(color="FFFFFF", bold=True)))
    F("B11", '=IF(«ISO»,"Isolated leg / head, b = "&«bw»&" mm – strip totals","Unit width – per metre")&"  ·  f\'c "&«fc»&" MPa  ·  second-order design actions from the analysis  ·  "&«model»&" block"')
    ws["B11"].font = font(7, color=GREY)
    table_header(ws, 13, [("B13:C13", "Check"), ("D13", "Basis"), ("E13", "Actual"), ("F13", "Limit"), ("G13", "Unit"),
                          ("H13", "Other"), ("I13", "Utilisation"), ("J13:K13", "Status")])
    heading(ws, 14, "Strength", sz=8)
    def chk(rr, s_, basis, actual, limit, unit_, util, status=None, nf="0.00", other=None, onf="0.00"):
        ws.merge_cells(f"B{rr}:C{rr}")
        if s_.startswith("="): F(f"B{rr}", s_)
        else: ws[f"B{rr}"].value = s_
        ws[f"B{rr}"].font = font(9); ws[f"B{rr}"].alignment = CEN
        desc(rr, basis)
        CA(f"E{rr}", actual, nf); CA(f"F{rr}", limit, nf)
        U(f"G{rr}", unit_); ws[f"G{rr}"].alignment = CEN
        CA(f"H{rr}", other, onf); CA(f"I{rr}", util, "0%")
        ws.merge_cells(f"J{rr}:K{rr}"); F(f"J{rr}", status if status else f'=IF(I{rr}="","–",IF(I{rr}<=1,"OK","NG"))')
        ws[f"J{rr}"].font = font(9, True, color=GREEN); ws[f"J{rr}"].alignment = CEN
    for i in range(NLC):
        rr = 15 + i; c = COLS[i]; lr = LC0 + i; gr = MAG0 + i; blank = f'$E${lr}=""'
        basis = f'=IF({blank},"(not used)","N* "&TEXT($D${gr},"0")&" kN, M*y span")'
        chk(rr, f'=$B${lr}&" N–M"', basis, f'=IF({blank},"",{c}{RM})', f'=IF({blank},"",{c}{C0 + 2})', "kNm",
            f'=IF({blank},"",{c}{RSC})', other=f'=IF({blank},"",{c}{C0})', onf='"φ "0.00')
    def gov(rng, row): return f'INDEX({row},MATCH(MAX({rng}),{rng},0))'
    trow = f"E{TR}:J{TR}"
    chk(21, "Support M*y", f'=IF(COUNT({trow})=0,"(no M*y support entered)","M*y at the support  ·  "&{gov(trow, f"$E${T0 - 1}:$J${T0 - 1}")})',
        f'=IF(COUNT({trow})=0,"",ABS({gov(trow, f"$E${T0 + 1}:$J${T0 + 1}")}))',
        f'=IF(COUNT({trow})=0,"",{gov(trow, f"$E${T0 + 2}:$J${T0 + 2}")})', "kNm", f'=IF(COUNT({trow})=0,"",MAX({trow}))')
    xrow = f"E{XR}:J{XR}"
    chk(22, "Horizontal M*x", f'=IF(COUNT({xrow})=0,"(no M*x entered)","horizontal bars, N = 0  ·  "&{gov(xrow, f"$E${X0 - 1}:$J${X0 - 1}")})',
        f'=IF(COUNT({xrow})=0,"",ABS({gov(xrow, f"$E${X0}:$J${X0}")}))',
        f'=IF(COUNT({xrow})=0,"",ABS({gov(xrow, f"$E${X0}:$J${X0}")})/MAX({xrow}))', "kNm", f'=IF(COUNT({xrow})=0,"",MAX({xrow}))')
    heading(ws, 23, "Serviceability", sz=8)
    for rr, c, lab_, ms in ((24, "E", "σst (Ms+)", "«Msp»"), (25, "F", "σst (Ms−)", "«Msn»")):
        chk(rr, lab_, f'=IF({c}{RMS}="","(no working moment entered)",IF(«slim»="","cracked, report only","cracked, ≤ σs limit")&"  ·  Ms "&TEXT({c}{RMS},"0.00")&", Ns "&TEXT(MAX(0,N(«Ns»)),"0.0"))',
            f'=IF({c}{RMS}="","",{c}{RST})', '=IF(N(«slim»)>0,«slim»,"")', "MPa", f'=IF(OR(E{rr}="",F{rr}=""),"",E{rr}/F{rr})',
            status=f'=IF(E{rr}="","–",IF(F{rr}="","info",IF(I{rr}<=1,"OK","NG")))', nf="0")
    heading(ws, 26, "Detailing", sz=8)
    chk(27, "ρv ≥ 0.0015", "vertical, main bars (Cl 11.7.1)", "=«rhov»", "=0.0015", "", '=IF(«rhov»=0,99,0.0015/«rhov»)', nf="0.000%")
    chk(28, '=IF(«hrest»="Yes","ρh ≥ 0.0025","ρh ≥ 0.0015")',
        '=IF(«hrest»="Yes","horizontal, restrained (Cl 11.7.2)","horizontal, free to move (Cl 11.7.2)")',
        "=«rhoh»", '=IF(«hrest»="Yes",0.0025,0.0015)', "", '=IF(«rhoh»=0,99,F28/«rhoh»)', nf="0.000%")
    chk(29, "spacing", "max(sv, sh) ≤ min(2.5 tw, 350) (Cl 11.7.3)", "=MAX(«vs»,«hs»)", "=MIN(2.5*«tw»,350)", "mm", "=E29/F29", nf="0")
    status_cf(ws, f"J{CHK0}:K{CHK1}"); util_cf(ws, f"I{CHK0}:I{CHK1}", f"I{CHK0}")
    F("B12", f'=IFERROR(IF({errs}>0,"Add-in error – check the add-in cells on the calculation pages",IF(COUNT(I{CHK0}:I{CHK1})=0,"","Governs: "&INDEX(B{CHK0}:B{CHK1},MATCH(MAX(I{CHK0}:I{CHK1}),I{CHK0}:I{CHK1},0))'
             f'&" – "&INDEX(D{CHK0}:D{CHK1},MATCH(MAX(I{CHK0}:I{CHK1}),I{CHK0}:I{CHK1},0)))),"")')
    ws["B12"].font = font(7, color=GREY)

    # ================================================================== resolve, arrays, validation, print
    for ref, f in pend: ws[ref].value = Fx(f)
    for ref in arrays:
        top = ref.split(":")[0]; ws[top].value = ArrayFormula(ref, ws[top].value)
    def dvl(formula, ref):
        v = DataValidation(type="list", formula1=formula, allow_blank=True); ws.add_data_validation(v)
        for part in ref.split(): v.add(part)
    dvl('"Rectangular,Curvilinear,Parabolic"', "E36"); dvl('"No,Yes"', "E48")
    dvl('"Central,Two layers"', "E42"); dvl('"' + ",".join(STRIP) + '"', "E39"); dvl('"' + ",".join(XPOS) + '"', "H46 H47"); dvl("=BarList", "E44:E47");
    f = Fx('NOT(«TWO»)')
    ws.conditional_formatting.add("E43", FormulaRule(formula=[f], fill=PatternFill(bgColor=NA_FILL), font=Font(color=NA_FONT)))
    ws.conditional_formatting.add("B43", FormulaRule(formula=[f], font=Font(color=NA_FONT)))
    ws.conditional_formatting.add("G39", FormulaRule(formula=[Fx('NOT(«ISO»)')], fill=PatternFill(bgColor=NA_FILL), font=Font(color=NA_FONT)))
    for rr in range(15, LAST + 1):
        if ws.row_dimensions[rr].height is None: ws.row_dimensions[rr].height = 14
    ws.print_area = [f"A1:K{LAST}", "M6:T50", "M52:T96"]
    for b_ in (51, CP - 1): ws.row_breaks.append(Break(id=b_))
    return ws


# ---------------- example (illustrative second-order analysis results: 180 panel, wind beam at 6.6 m, 0.8 m parapet) ----------------
p1 = dict(loc="Example – 180 panel, wind beam at 6.6 m, 0.8 m parapet", tw=180, fc=32, lay="Central",
          vbar="N12", vs=200, hbar="N12", hs=250,
          lcs=[("1.2G + 1.5Q", 9.0, -0.30, -0.45, 0.10),
               ("1.2G + Wu inward", 20.6, 10.10, -0.80, 2.60),
               ("0.9G + Wu inward", 15.5, 10.25, -0.75, 2.60),
               ("0.9G + Wu outward", 15.4, -8.90, 0.75, -2.30),
               ("1.2G + Wu outward", 20.5, -9.10, 0.80, -2.30),
               ("G + 0.3Q + Eu", 17.5, 4.40, -0.40, 1.10)],
          Ns=17.0, Msp=6.4, Msn=5.8)


# isolated leg beside an opening: 600 wide, strip totals (illustrative)
p2 = dict(loc="Example – 600 leg beside a roller door, 180 panel", tw=180, fc=32, lay="Central", strip="Isolated leg / head", bw=600,
          vbar="N12", vs=200, hbar="N12", hs=250, xbar="N16", xs=2, xpos="Inside face",
          lcs=[("1.2G + 1.5Q", 14.0, -0.60, -0.70, 0.10),
               ("1.2G + Wu inward", 31.0, 16.50, -1.20, 1.50),
               ("0.9G + Wu inward", 23.5, 16.70, -1.10, 1.50),
               ("0.9G + Wu outward", 23.4, -14.20, 1.10, -1.30),
               ("1.2G + Wu outward", 31.0, -14.50, 1.20, -1.30)],
          Ns=25.0, Msp=10.5, Msn=9.2)


if __name__ == "__main__":
    build_tab("P1", p1)
    build_tab("P2", p2)
    wb._sheets = [wb["P1"], wb["P2"], wb["Bars"]]
    wb.active = 0
    from openpyxl.workbook.properties import CalcProperties
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    wb.save(OUT)
    print("saved", OUT)
