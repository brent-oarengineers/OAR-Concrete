"""Builds template/OAR - RC Column Design.xlsx (OAR house calc style).

Run from the repo root:  python template/build_column_template.py
One calc tab = one column (rectangular or circular). Typical bar layout (perimeter / circle) plus up to
10 custom bars; up to 6 ULS load cases; braced slenderness (Cl 10.3, 10.4); uniaxial and biaxial capacity
by strain compatibility (OAR Concrete add-in) and Cl 10.6.4; M–N diagrams, moment contour and section plot.
Every code check is a visible worksheet formula; the add-in supplies the section analysis.
"""
import sys, os, re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from oar_style import *  # house-style helpers
from oar_style import BARS, BAR_NOTES

LOGO = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "assets", "oar-logo.png")
OUT = sys.argv[2] if len(sys.argv) > 2 else os.path.join(HERE, "OAR - RC Column Design.xlsx")
_fsproj = open(os.path.join(HERE, "..", "src", "OAR.Concrete.Excel", "OAR.Concrete.Excel.fsproj"), encoding="utf-8").read()
TEMPLATE_VERSION = re.search(r"<Version>([0-9]+\.[0-9]+\.[0-9]+)</Version>", _fsproj).group(1)
_maj, _min, _pat = (int(v) for v in TEMPLATE_VERSION.split("."))

wb = Workbook()
wb.remove(wb.active)

# =========================== DATA TAB: Bars ===========================
# Column template: Bars table carries Es, plus high-strength threadbar for shoring / piling custom bars
TB_ES = 205000
COL_BARS = [r + [200000] for r in BARS] + [[f"TB{d_}", d_n, a_, 950, "N", TB_ES] for d_, d_n, a_ in
                                            (("26.5", 26.5, 552), ("32", 32, 804), ("36", 36, 1018), ("40", 40, 1257), ("47", 47, 1735))]
COL_NOTES = dict(BAR_NOTES)
COL_NOTES["TB26.5"] = "threadbar 950/1050 (e.g. DYWIDAG) – confirm fsy, Es, area with the supplier"
D_ = wb.create_sheet("Bars")
setup_data_tab(D_, "Reinforcement – bar properties",
               "AS/NZS 4671 nominal values. S = 250 MPa deformed (D250N), R = plain round, L = mesh wire (Class L), TB = high-strength threadbar. Edit or add rows (any fsy, Es); the calc tabs read this list.",
               {"B": 10, "C": 10, "D": 10, "E": 10, "F": 10, "G": 10, "H": 44})
for j, t in enumerate(["Bar", "db (mm)", "Ab (mm²)", "fsy (MPa)", "Class", "Es (MPa)", "Note"]):
    c = D_.cell(4, 2 + j); c.value = t; c.font = font(8, True); c.fill = fill(HDR_FILL); c.border = BOX; c.alignment = CEN
BAR0 = 5; BAR_LAST = BAR0 + len(COL_BARS) + 5
for i in range(BAR_LAST - BAR0 + 1):
    row = COL_BARS[i] + [COL_NOTES.get(COL_BARS[i][0], "")] if i < len(COL_BARS) else [None] * 7
    for j, v in enumerate(row):
        c = D_.cell(BAR0 + i, 2 + j); c.value = v; c.font = font(8, color="000000"); c.fill = fill(IN_FILL); c.border = BOX
        c.alignment = LEFT if j in (0, 6) else CEN
EC_C = 11  # f'c / Ec table in K:L
for j, t in enumerate(["f'c (MPa)", "Ec (MPa)"]):
    c = D_.cell(4, EC_C + j); c.value = t; c.font = font(8, True); c.fill = fill(HDR_FILL); c.border = BOX; c.alignment = CEN
for i, (fc_, ec_) in enumerate(zip([20, 25, 32, 40, 50, 65, 80, 100], [24000, 26700, 30100, 32800, 34800, 37400, 39600, 42200])):
    for j, v in enumerate((fc_, ec_)):
        c = D_.cell(5 + i, EC_C + j); c.value = v; c.font = font(8, color="000000"); c.fill = fill(IN_FILL); c.border = BOX; c.alignment = CEN
c = D_.cell(13, EC_C); c.value = "AS 3600 Table 3.1.2 (ρ = 2400 kg/m³); calc tabs interpolate linearly."; c.font = font(7, i=True, color=GREY)
D_.column_dimensions["I"].width = 3; D_.column_dimensions["J"].width = 3; D_.column_dimensions["K"].width = 10; D_.column_dimensions["L"].width = 10
wb.defined_names["BarList"] = DefinedName("BarList", attr_text=f"Bars!$B${BAR0}:$B${BAR_LAST}")
wb.defined_names["EcFc"] = DefinedName("EcFc", attr_text="Bars!$K$5:$K$12")
wb.defined_names["EcVal"] = DefinedName("EcVal", attr_text="Bars!$L$5:$L$12")
BARTAB = f"Bars!$B${BAR0}:$G${BAR_LAST}"
def barv(ref, col): return f"VLOOKUP({ref},{BARTAB},{col},FALSE)"

NLC = 6          # load cases
NTYP = 40        # typical bar rows
NCUS = 10        # custom bar rows


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
        s = s.replace("«CIRC»", '«shape»="Circular"').replace("«POLY»", '«shape»="Custom polygon"').replace("«SHAPE»", 'IF(«shape»="Circular","Circle",IF(«shape»="Custom polygon","Polygon","Rect"))')
        for k in sorted(N, key=len, reverse=True): s = s.replace(f"«{k}»", N[k])
        assert "«" not in s, s
        return s
    def lab(ref, t, **k): return label(ws, ref, t, **k)
    def F(ref, f):  # formula text into a cell, resolved at the end
        pend.append((ref, f))
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

    setup_calc_tab(ws, LOGO, '="RC COLUMN (AS 3600-2018)  –  "&IF(G3="","(ref)",G3)&IF(I3="","","  ·  "&I3)', "COLUMN", name, "")
    ws["I3"].value = d["loc"]
    location_frame(ws, "Column")

    # ---- add-in version check (as the flexural sheet; V3:V7, not printed)
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
                      '"Copy this tab for additional columns.   ·   Calculated with OAR Concrete "&$V$4))')
    ws.conditional_formatting.add("B7", FormulaRule(formula=['OR($V$7,$V$4="")'], font=Font(color=RED, bold=True, italic=False)))

    # ================================================================== KEY POINTS OF THE INTERACTION CURVE (summary, page 1)
    heading(ws, 29, "Interaction Curve – Key Points (design capacities, uniaxial)")
    table_header(ws, 30, [("B30:C30", "Point"), ("D30", "Definition"), ("E30", "φN (kN)"), ("F30", "φMx (kNm)"),
                          ("G30", "φN (kN)"), ("H30", "φMy (kNm)"), ("I30", "φ  x"), ("J30", "φ  y")])
    def phin(a, n): return f"IF({n}>=«Nub{a}»,«phi0»,«phi0»+(«phib{a}»-«phi0»)*(1-{n}/«Nub{a}»))"
    KP = [("Squash", "pure compression, M = 0 (Cl 10.6.2.2)",
           lambda a: ("«phi0»*«Nuo{}»".format(a), "0", "«phi0»")),
          ("Decompression", "zero strain at outermost bar (kuo = 1)",
           lambda a: (f"{phin(a, f'«Nud{a}»')}*«Nud{a}»", f"{phin(a, f'«Nud{a}»')}*«Mud{a}»", phin(a, f"«Nud{a}»"))),
          ("Balanced", "outermost tension bar at yield",
           lambda a: (f"«phi0»*«Nub{a}»", f"«phi0»*«Mub{a}»", "«phi0»")),
          ("Pure bending", "N = 0",
           lambda a: ("0", f"«phib{a}»*«Muo{a}»", f"«phib{a}»")),
          ("Pure tension", "all bars yielded, M = 0",
           lambda a: (f"«phit{a}»*«Nuot{a}»", "0", f"«phit{a}»"))]
    for k, (pt, ds, fn) in enumerate(KP):
        rr = 31 + k
        ws.merge_cells(f"B{rr}:C{rr}"); lab(f"B{rr}", pt); desc(rr, ds)
        for a, cN, cM, cP in (("x", "E", "F", "I"), ("y", "G", "H", "J")):
            fN, fM, fP = fn(a)
            CA(f"{cN}{rr}", f'=IFERROR({fN},"")', "0"); CA(f"{cM}{rr}", f'=IFERROR({fM},"")', "0.0"); CA(f"{cP}{rr}", f'=IFERROR({fP},"")', "0.00")
    note(ws, 36, "Uniaxial points from OAR.CONCRETE.COLUMNAXIS (calculation page). N compression +. φ from Table 2.2.2: φ0 at and above Nub, φb at N = 0, Class N in tension → 0.85.")

    # ================================================================== INPUTS (page 1)
    heading(ws, 38, "Section and Concrete")
    lab("B39", "Shape"); IN("E39:F39", d["shape"], name="shape")
    F("B40", '=IF(«CIRC»,"Diameter D","Width b (x direction)")'); ws["B40"].font = font(9)
    IN("E40", d["b"], "0", name="b"); U("F40", "mm")
    lab("B41", "Depth D (y direction)"); IN("E41", d.get("D"), "0", name="D"); U("F41", "mm"); small("H41", "rectangular only")
    lab("B42", "Concrete strength f'c"); IN("E42", d["fc"], "0", name="fc"); U("F42", "MPa")
    lab("B43", "Stress block (ULS)"); IN("E43:F43", d.get("model", "Rectangular"), name="model")
    lab("B44", "Ec override"); IN("E44", None, "0", name="ecin"); U("F44", "MPa"); small("H44", "blank = Table 3.1.2")
    lab("B45", "φ0 for compression, Table 2.2.2(d)"); IN("E45", d.get("phi0", 0.6), "0.00", name="phi0"); small("H45", "Amdt 2: φ0 = 0.65 kφ – enter the value")
    lab("B46", "Biaxial check governing"); IN("E46:F46", d.get("method", "Cl 10.6.4"), name="method"); small("H46", "circular / custom: strain compatibility")

    heading(ws, 48, "Reinforcement – typical layout")
    lab("B49", "Cover to ligature"); IN("E49", d["cover"], "0", name="cover"); U("F49", "mm")
    lab("B50", "Ligature bar"); IN("E50", d["lig"], name="lig")
    lab("B51", "Main bar (typical)"); IN("E51", d["bar"], name="bar"); small("H51", '"None" = custom bars only')
    F("B52", '=IF(«CIRC»,"Number of bars","Bars per b face (incl. corners)")'); ws["B52"].font = font(9)
    IN("E52", d["n1"], "0", name="n1")
    lab("B53", "Bars per D face (incl. corners)"); IN("E53", d.get("n2"), "0", name="n2"); small("H53", "rectangular only")
    lab("B54", "fsy override, typical bars (development)"); IN("E54", d.get("fsyo"), "0", name="fsyo"); U("F54", "MPa")
    small("H54", "blank = bar fsy; short laps / starters: fsy × Lprov / Lsy.t")
    note(ws, 55, "Custom bars (x, y, bar) go in the Bar Schedule (last page) and use their own bar's fsy; the override applies to the typical bars only.")

    heading(ws, 56, "Slenderness – braced column (Cl 10.3, 10.4)")
    table_header(ws, 57, [("E57", "x axis (Mx)"), ("F57", "y axis (My)")])
    lab("B58", "Unsupported length Lu"); IN("E58", d["Lu"], "0", name="Lu"); U("G58", "mm")
    lab("B59", "Effective length factor k"); IN("E59", d.get("kx", 1.0), "0.00", name="kx"); IN("F59", d.get("ky", 1.0), "0.00", name="ky")
    lab("B60", "βd = G / (G + Q), sustained axial load"); IN("E60", d.get("betad", 0.6), "0.00", name="betad")
    note(ws, 61, "Unbraced (sway) columns are not covered. βd is taken as 0 where Le/r ≤ 40 and N* ≤ M2*/(2D) (Cl 10.4.3).")

    heading(ws, 63, "Load Cases (ULS)")
    table_header(ws, 64, [("B64:C64", "Case"), ("D64", "Description"), ("E64", "N* (kN)"), ("F64", "M*x top"), ("G64", "M*x bot"),
                          ("H64", "M*y top"), ("I64", "M*y bot")])
    LC0 = 65
    for i in range(NLC):
        r = LC0 + i; lc = d["lcs"][i] if i < len(d["lcs"]) else None
        IN(f"B{r}:C{r}", f"LC{i + 1}")
        IN(f"D{r}", lc[0] if lc else None, al=LEFT)
        for j, col in enumerate("EFGHI"): IN(f"{col}{r}", lc[1 + j] if lc else None, "0.0")
    small("J64", "kNm")
    note(ws, 71, "N* compression +. End moments: same sign = single curvature. Blank N* = case not used. Minimum eccentricity 0.05D applies about each axis (Cl 10.1.2).")
    note(ws, 72, "Diagrams (section, M–N, moment contour for each case): OAR ribbon > Column graphics.")

    # ================================================================== SECTION + BAR SCHEDULE (helpers first; formulas reference them)
    # bar schedule rows
    BS0 = 208
    # custom (input) rows first, then the typical layout (add-in array)
    CUS0, CUS1 = BS0 + 1, BS0 + NCUS
    TYP0, TYP1 = CUS1 + 1, CUS1 + NTYP
    RS0, RS1 = CUS0, TYP1
    BARS_ARG = f"$L${RS0}:$P${RS1}"
    CLS_ARG = f"$J${RS0}:$J${RS1}"
    SEC = '«SHAPE»,IF(«POLY»,«corners»,«b»:«D»)'
    MOD = '«fc»,«model»'

    # ================================================================== CALCULATIONS (page 2)
    # custom outline (Shape = Custom polygon)
    heading(ws, 74, "Custom Section Outline (Shape = Custom polygon)")
    note(ws, 75, "Corners in order round the section, x right / y up (mm, any origin). Rise: the edge to the next corner is an arc, its mid-point Rise from the chord (+ out, − in).")
    table_header(ws, 76, [("B76:C76", "Corner"), ("D76", "edge to next corner"), ("E76", "x (mm)"), ("F76", "y (mm)"), ("G76", "rise (mm)"), ("H76", "bars on edge")])
    NCOR = 12
    cors = d.get("corners", [])
    for k_ in range(NCOR):
        rr = 77 + k_
        ws.merge_cells(f"B{rr}:C{rr}"); ws[f"B{rr}"].value = k_ + 1; ws[f"B{rr}"].font = font(8, color=GREY); ws[f"B{rr}"].alignment = CEN
        c_ = cors[k_] if k_ < len(cors) else (None, None, None, None)
        F(f"D{rr}", f'=IF(E{rr}="","",IF(G{rr}<>0,"arc","straight")&" to corner "&IF(OR(E{rr + 1}="",{k_ + 1}={NCOR}),1,{k_ + 2}))' if k_ < NCOR - 1 else f'=IF(E{rr}="","",IF(G{rr}<>0,"arc","straight")&" to corner 1")')
        ws[f"D{rr}"].font = font(8, color=GREY)
        IN(f"E{rr}", c_[0], "0.0"); IN(f"F{rr}", c_[1], "0.0"); IN(f"G{rr}", c_[2], "0.0"); IN(f"H{rr}", c_[3], "0")
    N["corners"] = f"$E$77:$H${76 + NCOR}"
    note(ws, 89, "Typical bars: one at every corner plus 'bars on edge' along each edge, at cover + ligature + db/2 from the faces (following arcs). Bars per face inputs not used.")

    heading(ws, 91, "Section Properties")
    table_header(ws, 92, [("E92", "x axis (Mx)"), ("F92", "y axis (My)")])
    pend.append(("M93", f'=OAR.CONCRETE.SECTIONPROPS({SEC})')); arrays.append("M93:M97")
    for k_, t_ in enumerate(["b (bounding)", "D (bounding)", "Ag", "Ix", "Iy"]):
        lab(f"N{93 + k_}", "SECTIONPROPS: " + t_, sz=7, color=GREY); ws[f"M{93 + k_}"].font = font(7, color=GREY)
    nm("pb", "M93"); nm("pD", "M94"); nm("pA", "M95"); nm("pIx", "M96"); nm("pIy", "M97")
    r = 93; sym(r, "Ag"); desc(r, '=IF(«CIRC»,"π D² / 4",IF(«POLY»,"custom outline (arcs as 48 chords)","b D"))'); CA(f"E{r}", '=IF(«POLY»,«pA»,IF(«CIRC»,PI()*«b»^2/4,«b»*«D»))', "0", name="Ag"); U(f"G{r}", "mm²")
    r = 94; sym(r, "As"); desc(r, "all bars (Bar Schedule)"); CA(f"E{r}", f"=SUM($L${RS0}:$L${RS1})", "0", name="As"); U(f"G{r}", "mm²")
    r = 95; sym(r, "ρ"); desc(r, "As / Ag"); CA(f"E{r}", "=«As»/«Ag»", "0.00%", name="rho")
    r = 96; sym(r, "n"); desc(r, "typical | custom bars")
    CA(f"E{r}", f'=COUNTIFS($G${TYP0}:$G${TYP1},"<>",$H${TYP0}:$H${TYP1},">0")', "0", name="ntyp")
    CA(f"F{r}", f'=COUNTIF($H${CUS0}:$H${CUS1},">0")', "0", name="ncus")
    r = 97; sym(r, "edge"); desc(r, "cover + ligature + db / 2")
    CA(f"E{r}", f'=«cover»+IFERROR({barv("«lig»", 2)},0)+IFERROR({barv("«bar»", 2)},0)/2', "0.0", name="edge"); U(f"G{r}", "mm")
    r = 98; sym(r, "Dd"); desc(r, "section depth in the bending direction")
    CA(f"E{r}", '=IF(«POLY»,«pD»,IF(«CIRC»,«b»,«D»))', "0", name="Dx"); CA(f"F{r}", '=IF(«POLY»,«pb»,«b»)', "0", name="Dy"); U(f"G{r}", "mm")
    r = 99; sym(r, "r"); desc(r, '=IF(«POLY»,"√(I / Ag), custom outline (Cl 10.5.2)",IF(«CIRC»,"0.25 D (Cl 10.5.2)","0.3 × depth (Cl 10.5.2)"))')
    CA(f"E{r}", '=IF(«POLY»,SQRT(«pIx»/«pA»),IF(«CIRC»,0.25,0.3)*«Dx»)', "0.0", name="rx"); CA(f"F{r}", '=IF(«POLY»,SQRT(«pIy»/«pA»),IF(«CIRC»,0.25,0.3)*«Dy»)', "0.0", name="ry"); U(f"G{r}", "mm")
    r = 100; sym(r, "Le"); desc(r, "k Lu"); CA(f"E{r}", "=«kx»*«Lu»", "0", name="Lex"); CA(f"F{r}", "=«ky»*«Lu»", "0", name="Ley"); U(f"G{r}", "mm")
    r = 101; sym(r, "Le / r"); desc(r, "slenderness ratio"); CA(f"E{r}", "=«Lex»/«rx»", "0.0", name="slx"); CA(f"F{r}", "=«Ley»/«ry»", "0.0", name="sly")
    r = 102; sym(r, "Ec"); desc(r, '=IF(N(«ecin»)>0,"entered","Table 3.1.2, interpolated")')
    I_ = "MIN(MATCH(«fc»,EcFc,1),7)"
    CA(f"E{r}", f"=IF(N(«ecin»)>0,«ecin»,INDEX(EcVal,{I_})+(«fc»-INDEX(EcFc,{I_}))*(INDEX(EcVal,{I_}+1)-INDEX(EcVal,{I_}))/(INDEX(EcFc,{I_}+1)-INDEX(EcFc,{I_})))", "0", name="Ec")
    U(f"G{r}", "MPa")

    heading(ws, 104, "Axis Properties (OAR Concrete add-in)")
    note(ws, 105, "OAR.CONCRETE.COLUMNAXIS: strain compatibility, εcu 0.003. Nuo = α1 f'c Ac + fsy As (Ac net of bars, Cl 10.6.2.2); balanced point with the outermost tension bar at yield.")
    table_header(ws, 106, [("E106", "x axis (Mx)"), ("F106", "y axis (My)")])
    AX = [("Nuo", "squash load", "0", "kN"), ("Nuot", "pure tension", "0", "kN"), ("Nub", "balanced point", "0", "kN"),
          ("Mub", "moment at Nub (Mc, Cl 10.4.4)", "0.0", "kNm"), ("Muo", "pure bending", "0.0", "kNm"), ("kuo", "at pure bending", "0.000", ""),
          ("φb", "Tbl 2.2.2, pure bending", "0.000", ""), ("do", "to outermost tension bar", "0.0", "mm"),
          ("Nud", "decompression (dn = do)", "0", "kN"), ("Mud", "moment at Nud", "0.0", "kNm"), ("φt", "Tbl 2.2.2, pure tension", "0.000", "")]
    AX0 = 107
    for i, (s_, ds, nf, u) in enumerate(AX):
        rr = AX0 + i; sym(rr, s_); desc(rr, ds); calc(ws, f"E{rr}", None, nf); calc(ws, f"F{rr}", None, nf); U(f"G{rr}", u)
    AX1 = AX0 + len(AX) - 1
    for col, ax in (("E", "x"), ("F", "y")):
        F(f"{col}{AX0}", f'=OAR.CONCRETE.COLUMNAXIS({SEC},{BARS_ARG},{MOD},"{ax}",{CLS_ARG},«Ec»)'); arrays.append(f"{col}{AX0}:{col}{AX1}")
        for k, n_ in enumerate(["Nuo", "Nuot", "Nub", "Mub", "Muo", "kuo", "phib", "do", "Nud", "Mud", "phit"]): nm(f"{n_}{ax}", f"{col}{AX0 + k}")
    r = AX1 + 1; sym(r, "φNuo"); desc(r, "φ0 Nuo"); CA(f"E{r}", "=«phi0»*«Nuox»", "0", name="phiNuo"); U(f"G{r}", "kN")

    # ---- design moments, per axis
    def dm_table(top_row, ax, ctop, cbot, Dd, sl, Le, Nuo, Mub, do_):
        heading(ws, top_row, f"Design Moments – {'M*x (x axis)' if ax == 'x' else 'M*y (y axis)'}", sz=8)
        h = top_row + 1
        table_header(ws, h, [(f"B{h}:C{h}", "Case"), (f"D{h}", "N* (kN)"), (f"E{h}", "M2* (kNm)"), (f"F{h}", "M1*/M2*"), (f"G{h}", "αc"),
                             (f"H{h}", "Le/r limit"), (f"I{h}", "Nc (kN)"), (f"J{h}", "δb"), (f"K{h}", "M*d")])
        for i in range(NLC):
            rr = h + 1 + i; lr = LC0 + i
            Nn = f"$E${lr}"; T = f"${ctop}${lr}"; Bm = f"${cbot}${lr}"
            blank = f'{Nn}=""'
            ws.merge_cells(f"B{rr}:C{rr}"); F(f"B{rr}", f"=$B${lr}"); ws[f"B{rr}"].font = font(8); ws[f"B{rr}"].alignment = CEN
            CA(f"D{rr}", f'=IF({blank},"",{Nn})', "0")
            CA(f"E{rr}", f'=IF({blank},"",MAX(ABS(N({T})),ABS(N({Bm}))))', "0.0")
            emin = f"0.05*{Dd}/1000*{Nn}"
            CA(f"F{rr}", f'=IF({blank},"",IF(E{rr}<={emin},-1,IF(N({T})*N({Bm})>=0,-1,1)*MIN(ABS(N({T})),ABS(N({Bm})))/E{rr}))', "0.00")
            q = f"{Nn}/(0.6*{Nuo})"
            CA(f"G{rr}", f'=IF({blank},"",IF({Nn}<=0,9.99,IF({q}>=0.15,SQRT(MAX(0,2.25-2.5*{q})),SQRT(1/(3.5*{q})))))', "0.00")
            CA(f"H{rr}", f'=IF({blank},"",MAX(25,G{rr}*(38-«fc»/15)*(1+F{rr})))', "0.0")
            bd = f"IF(AND({sl}<=40,{Nn}<=E{rr}/(2*{Dd}/1000)),0,«betad»)"
            CA(f"I{rr}", f'=IF(OR({blank},{Le}=0),"",PI()^2/{Le}^2*182*{do_}*«phi0»*{Mub}*1E6/(1+{bd})/1000)', "0")
            km = f"MAX(0.4,0.6-0.4*F{rr})"
            CA(f"J{rr}", f'=IF({blank},"",IF({sl}<=H{rr},1,IF({Nn}>=I{rr},99,MAX(1,{km}/(1-{Nn}/I{rr})))))', "0.00")
            sgn = f"IF(ABS(N({T}))>=ABS(N({Bm})),IF(N({T})<0,-1,1),IF(N({Bm})<0,-1,1))"
            CA(f"K{rr}", f'=IF({blank},"",{sgn}*J{rr}*MAX(E{rr},{emin}))', "0.0")
        return h + 1
    heading(ws, 120, "Design Moments – minimum eccentricity and moment magnification (braced)")
    note(ws, 121, "M2* larger end moment; M1*/M2* negative in single curvature, −1 if M2* ≤ 0.05 D N*. Short if Le/r ≤ max(25, αc (38 − f'c/15)(1 + M1*/M2*)), αc from N*/(0.6 Nuo).")
    note(ws, 122, "Nc = π²/Le² × 182 do φ0 Mc / (1 + βd), Mc = Mub.  δb = km / (1 − N*/Nc) ≥ 1, km = 0.6 − 0.4 M1*/M2* ≥ 0.4 (δb = 1 if short).  M*d = δb × max(M2*, 0.05 D N*).")
    DX0 = dm_table(123, "x", "F", "G", "«Dx»", "«slx»", "«Lex»", "«Nuox»", "«Mubx»", "«dox»")
    DY0 = dm_table(131, "y", "H", "I", "«Dy»", "«sly»", "«Ley»", "«Nuoy»", "«Muby»", "«doy»")

    # ---- capacity along the load direction (strain compatibility, biaxial)
    CP = 140
    heading(ws, CP, "Capacity – strain compatibility along the load direction (OAR Concrete add-in)")
    note(ws, CP + 1, "OAR.CONCRETE.COLUMN: finds the neutral-axis direction θ for which the capacity (Mux, Muy) at N* points along (M*x,d, M*y,d); φ from Tbl 2.2.2 (φ0 … φb by N*/Nub).")
    h = CP + 2
    table_header(ws, h, [(f"{c}{h}", "") for c in "EFGHIJ"])
    for i, c in enumerate("EFGHIJ"): F(f"{c}{h}", f"=$B${LC0 + i}")
    COLS = "EFGHIJ"
    CPR = [("φ", "Tbl 2.2.2 at N*", "0.000", ""), ("Mu", "capacity, load direction", "0.0", "kNm"), ("φMu", "", "0.0", "kNm"),
           ("Mux", "component about x", "0.0", "kNm"), ("Muy", "component about y", "0.0", "kNm"), ("θ", "max compression direction", "0.0", "deg"),
           ("dn", "neutral axis depth", "0.0", "mm"), ("Nub", "balanced, this direction", "0", "kN"), ("Nuo", "squash", "0", "kN"),
           ("εs", "outermost tension bar strain at capacity", "0.00%", "")]
    C0 = h + 1
    for k, (s_, ds, nf, u) in enumerate(CPR):
        rr = C0 + k; sym(rr, s_); desc(rr, ds)
        for c in COLS: calc(ws, f"{c}{rr}", None, nf)
        U(f"K{rr}", u)
    C1 = C0 + len(CPR) - 1
    for i, c in enumerate(COLS):
        lr = LC0 + i
        F(f"{c}{C0}", f'=IF($E${lr}="","",OAR.CONCRETE.COLUMN({SEC},{BARS_ARG},{MOD},$E${lr},$K${DX0 + i},$K${DY0 + i},«phi0»,{CLS_ARG},«Ec»))')
        arrays.append(f"{c}{C0}:{c}{C1}")
    r = C1 + 1; sym(r, "|M*d|"); desc(r, "√(M*x,d² + M*y,d²)")
    for i, c in enumerate(COLS): CA(f"{c}{r}", f'=IF($E${LC0 + i}="","",SQRT($K${DX0 + i}^2+$K${DY0 + i}^2))', "0.0")
    U(f"K{r}", "kNm"); RM = r
    r += 1; sym(r, "ratio"); desc(r, "|M*d| / φMu")
    for i, c in enumerate(COLS): CA(f"{c}{r}", f'=IF($E${LC0 + i}="","",{c}{RM}/{c}{C0 + 2})', "0.00")
    RSC = r

    # ---- Cl 10.6.4
    BX = RSC + 2
    heading(ws, BX, "Biaxial Bending – Cl 10.6.4")
    note(ws, BX + 1, "(M*x,d/φMux)^αn + (M*y,d/φMuy)^αn ≤ 1, αn = 0.7 + 1.7 N*/(0.6 Nuo), 1 ≤ αn ≤ 2. φMux, φMuy: uniaxial capacity at N* (OAR.CONCRETE.COLUMN).")
    h = BX + 2
    table_header(ws, h, [(f"{c}{h}", "") for c in COLS])
    for i, c in enumerate(COLS): F(f"{c}{h}", f"=$B${LC0 + i}")
    rows_b = [("N*", "kN", "0"), ("M*x,d", "kNm", "0.0"), ("M*y,d", "kNm", "0.0"), ("φMux", "kNm", "0.0"), ("φMuy", "kNm", "0.0"), ("αn", "", "0.00"), ("Σ", "", "0.00")]
    B0 = h + 1
    for k, (s_, u, nf) in enumerate(rows_b):
        rr = B0 + k; sym(rr, s_); U(f"K{rr}", u)
        for i, c in enumerate(COLS):
            lr = LC0 + i; blank = f'$E${lr}=""'
            mx, my = f"$K${DX0 + i}", f"$K${DY0 + i}"
            f = {0: f"=IF({blank},\"\",$E${lr})", 1: f"=IF({blank},\"\",{mx})", 2: f"=IF({blank},\"\",{my})",
                 3: f'=IF({blank},"",INDEX(OAR.CONCRETE.COLUMN({SEC},{BARS_ARG},{MOD},$E${lr},IF({mx}<0,-1,1),0,«phi0»,{CLS_ARG},«Ec»),3,1))',
                 4: f'=IF({blank},"",INDEX(OAR.CONCRETE.COLUMN({SEC},{BARS_ARG},{MOD},$E${lr},0,IF({my}<0,-1,1),«phi0»,{CLS_ARG},«Ec»),3,1))',
                 5: f'=IF({blank},"",MIN(2,MAX(1,0.7+1.7*$E${lr}/(0.6*«Nuox»))))',
                 6: f'=IF({blank},"",(ABS({c}{B0 + 1})/{c}{B0 + 3})^{c}{B0 + 5}+(ABS({c}{B0 + 2})/{c}{B0 + 4})^{c}{B0 + 5})'}[k]
            CA(f"{c}{rr}", f, nf)
    desc(B0 + 3, "uniaxial at N*"); desc(B0 + 4, "uniaxial at N*")
    B1064 = B0 + 6
    CALC_END = B1064

    # ================================================================== GRAPHICS: drawn by the add-in window (OAR > Column graphics)
    CH = CALC_END + 2
    heading(ws, CH, "Interaction Diagrams and Section")
    note(ws, CH + 1, "OAR ribbon > Column graphics: the section and bars to scale, M–N about x and y with the design actions, and the moment contour at each case's N*.")
    note(ws, CH + 2, "Same engine and inputs as this sheet. Save PDF from the window (one picture, or all pages) to add the diagrams to the calc record.")

    # ================================================================== BAR SCHEDULE (last page)
    heading(ws, BS0 - 3, "Bar Schedule")
    note(ws, BS0 - 2, "Custom bars C1–C10: enter x, y (mm from the centre) and bar – any bar on the Bars tab (e.g. TB threadbar, or add a row with its own fsy and Es). Typical rows: OAR.CONCRETE.BARLAYOUT (do not edit).")
    table_header(ws, BS0 - 1, [(f"B{BS0 - 1}:C{BS0 - 1}", "#"), (f"D{BS0 - 1}", "Source"), (f"E{BS0 - 1}", "x (mm)"), (f"F{BS0 - 1}", "y (mm)"),
                               (f"G{BS0 - 1}", "Bar"), (f"H{BS0 - 1}", "As (mm²)"), (f"I{BS0 - 1}", "fsy (MPa)"), (f"J{BS0 - 1}", "Class"), (f"K{BS0 - 1}", "Es (GPa)")])
    ws.row_dimensions[BS0].height = 6
    lab(f"L{BS0 - 1}", "engine: As | depth | fsy | x from left | Es", sz=7, color=GREY)
    for k in range(NCUS + NTYP):
        rr = CUS0 + k; typ = k >= NCUS
        ws.merge_cells(f"B{rr}:C{rr}"); ws[f"B{rr}"].value = (k - NCUS + 1) if typ else f"C{k + 1}"
        ws[f"B{rr}"].font = font(8, color=GREY); ws[f"B{rr}"].alignment = CEN
        ws[f"D{rr}"].value = "typical (calculated)" if typ else "custom – enter x, y, bar"
        ws[f"D{rr}"].font = font(8, color=GREY if typ else "000000", i=not typ)
        if typ:
            calc(ws, f"E{rr}", None, "0.0"); calc(ws, f"F{rr}", None, "0.0")
            CA(f"G{rr}", f'=IF(E{rr}="","",«bar»)')
        else:
            c = d["custom"][k] if k < len(d.get("custom", [])) else (None, None, None)
            IN(f"E{rr}", c[0], "0.0"); IN(f"F{rr}", c[1], "0.0"); IN(f"G{rr}", c[2])
        CA(f"H{rr}", f'=IF(OR(G{rr}="",G{rr}="None"),"",{barv(f"G{rr}", 3)})', "0")
        CA(f"I{rr}", f'=IF(H{rr}="","",{f'IF(«fsyo»="",{barv(f"G{rr}", 4)},MIN(«fsyo»,{barv(f"G{rr}", 4)}))' if typ else barv(f"G{rr}", 4)})', "0")
        CA(f"J{rr}", f'=IF(H{rr}="","",{barv(f"G{rr}", 5)})')
        CA(f"K{rr}", f'=IF(H{rr}="","",IFERROR({barv(f"G{rr}", 6)}/1000,200))', "0")
        # engine columns
        for col, f in (("L", f"=N(H{rr})"), ("M", f'=IF(L{rr}=0,"",«Dx»/2-F{rr})'),
                       ("N", f'=IF(L{rr}=0,"",I{rr})'), ("O", f'=IF(L{rr}=0,"",«Dy»/2+E{rr})'), ("P", f'=IF(L{rr}=0,"",K{rr}*1000)')):
            pend.append((f"{col}{rr}", f)); ws[f"{col}{rr}"].font = font(7, color=GREY)
    pend.append((f"E{TYP0}", f'=IF(OR(«bar»="",«bar»="None"),"",OAR.CONCRETE.BARLAYOUT({SEC},«edge»,«n1»,«n2»,{NTYP}))'))
    arrays.append(f"E{TYP0}:F{TYP1}")
    LAST = RS1 + 2
    sym(RS1 + 1, "Σ"); desc(RS1 + 1, "total"); CA(f"H{RS1 + 1}", f"=«As»", "0"); U(f"I{RS1 + 1}", "mm²")
    note(ws, LAST, "Ligature spacing and size, laps and anchorage are not covered by this calc.")

    # ================================================================== SUMMARY + CHECKS (page 1)
    heading(ws, 8, "Design Summary")
    lab("B9", "Adopted section", sz=7, color=GREY); lab("F9", "Max utilisation", sz=7, color=GREY); lab("I9", "Status", sz=7, color=GREY)
    ws.merge_cells("B10:D10"); ws.merge_cells("F10:G10"); ws.merge_cells("I10:K10")
    F("B10", '=IF(«CIRC»,"Ø"&«b»,IF(«POLY»,"custom "&ROUND(«Dy»,0)&" × "&ROUND(«Dx»,0),«b»&" × "&«D»))&IF(«ntyp»>0,"  ·  "&«ntyp»&«bar»,"")&IF(«ncus»>0," + "&«ncus»&" custom","")&"  ·  f\'c "&«fc»')
    ws["B10"].font = font(12, True, color=GREEN); ws["B10"].alignment = LEFT
    CHK0, CHK1 = 15, 24
    c = ws["F10"]; c.value = f'=IF(COUNT(I{CHK0}:I{CHK1})=0,"–",MAX(I{CHK0}:I{CHK1}))'; c.font = font(12, True, color=GREEN); c.number_format = "0%"; c.alignment = LEFT
    errs = f'COUNTIF(E{AX0}:F{AX1},"#OAR*")+COUNTIF(E{C0}:J{C1},"#OAR*")+COUNTIF(E{TYP0}:E{TYP1},"#OAR*")'
    c = ws["I10"]; c.value = (f'=IF($V$7,"UPDATE ADD-IN",IF({errs}>0,"INPUT ERROR",IF(COUNT(I{CHK0}:I{CHK1})=0,"PENDING",'
                              f'IF(OR(MAX(I{CHK0}:I{CHK1})>1,COUNTIF(J{CHK0}:J{CHK1},"NG")>0),"FAIL","PASS"))))')
    c.font = font(10, True, color=GREY); c.alignment = CEN
    for x in ws["I10:K10"][0]: x.border = BOX
    passfail_cf(ws, "I10:K10", "$I$10")
    ws.conditional_formatting.add("I10:K10", FormulaRule(formula=['$I$10="UPDATE ADD-IN"'], fill=PatternFill(bgColor="C55A11"), font=Font(color="FFFFFF", bold=True)))
    F("B11", '="Lu "&«Lu»&" mm  ·  φ0 "&TEXT(«phi0»,"0.00")&IF(«fsyo»="","","  ·  fsy "&«fsyo»&" (typ.)")&"  ·  "&«model»&" block  ·  biaxial: "&IF(«shape»<>"Rectangular","strain compatibility",«method»)')
    ws["B11"].font = font(7, color=GREY)
    firsterr = f'IFERROR(INDEX(E{AX0}:F{AX1},MATCH("#OAR*",E{AX0}:E{AX1},0),1),IFERROR(INDEX(E{C0}:J{C0},1,MATCH("#OAR*",E{C0}:J{C0},0)),IFERROR(INDEX(E{TYP0}:E{TYP1},MATCH("#OAR*",E{TYP0}:E{TYP1},0)),"")))'
    ws["F11"].value = (f'=IFERROR(IF({errs}>0,{firsterr},IF(COUNT(I{CHK0}:I{CHK1})=0,"","Governs: "&INDEX(B{CHK0}:B{CHK1},MATCH(MAX(I{CHK0}:I{CHK1}),I{CHK0}:I{CHK1},0))'
                       f'&" – "&INDEX(D{CHK0}:D{CHK1},MATCH(MAX(I{CHK0}:I{CHK1}),I{CHK0}:I{CHK1},0)))),"")')
    ws["F11"].font = font(7, color=GREY)
    table_header(ws, 13, [("B13:C13", "Check"), ("D13", "Basis"), ("E13", "Actual"), ("F13", "Limit"), ("G13", "Unit"),
                          ("H13", "Other method"), ("I13", "Utilisation"), ("J13:K13", "Status")])
    heading(ws, 14, "Strength", sz=8)
    def chk(rr, s_, basis, actual, limit, unit_, util, status=None, nf="0.0", ratio=None, rnf="0.00"):
        ws.merge_cells(f"B{rr}:C{rr}")
        if s_.startswith("="): F(f"B{rr}", s_)
        else: ws[f"B{rr}"].value = s_
        ws[f"B{rr}"].font = font(9); ws[f"B{rr}"].alignment = CEN
        desc(rr, basis)
        CA(f"E{rr}", actual, nf); CA(f"F{rr}", limit, nf)
        if unit_.startswith("="): F(f"G{rr}", unit_); ws[f"G{rr}"].font = font(9, color=GREY)
        else: U(f"G{rr}", unit_)
        ws[f"G{rr}"].alignment = CEN
        CA(f"H{rr}", ratio, rnf); CA(f"I{rr}", util, "0%")
        ws.merge_cells(f"J{rr}:K{rr}"); F(f"J{rr}", status if status else f'=IF(I{rr}="","–",IF(I{rr}<=1,"OK","NG"))')
        ws[f"J{rr}"].font = font(9, True, color=GREEN); ws[f"J{rr}"].alignment = CEN
    chk(15, "N* ≤ φNuo", "largest N*, squash (Cl 10.6.2.2)", f'=IF(COUNT($E${LC0}:$E${LC0 + NLC - 1})=0,"",MAX($E${LC0}:$E${LC0 + NLC - 1}))',
        "=«phiNuo»", "kN", '=IF(E15="","",E15/F15)', nf="0")
    for i in range(NLC):
        rr = 16 + i; c = COLS[i]; lr = LC0 + i; blank = f'$E${lr}=""'
        use1064 = 'AND(«shape»="Rectangular",«method»="Cl 10.6.4")'
        chk(rr, f'=$B${lr}&" N–M"', f'=IF({blank},"(not used)",IF({use1064},"Cl 10.6.4, αn = "&TEXT({c}{B0 + 5},"0.00"),"strain compat., load direction")&IF(MAX($J${DX0 + i},$J${DY0 + i})>1,"  ·  δb "&TEXT(MAX($J${DX0 + i},$J${DY0 + i}),"0.00"),""))',
            f'=IF({blank},"",IF({use1064},{c}{B1064},{c}{RM}))', f'=IF({blank},"",IF({use1064},1,{c}{C0 + 2}))',
            f'=IF({use1064},"–","kNm")', f'=IF({blank},"",IF({use1064},{c}{B1064},{c}{RSC}))',
            ratio=f'=IF({blank},"",IF({use1064},{c}{RSC},IF(«shape»<>"Rectangular","",{c}{B1064})))', nf="0.00")
    chk(22, "ρ ≥ 0.01", '=IF(«rho»>=0.01,"Cl 10.7.1(a)","below 1 %: Cl 10.7.1(b) lightly loaded / AS 2159 piles – engineer to confirm")',
        "=0.01", "=«rho»", "", '=IF(OR(«rho»=0,«rho»<0.01),"",0.01/«rho»)', nf="0.00%",
        status='=IF(«rho»=0,"–",IF(«rho»>=0.01,"OK","OK*"))')
    chk(23, "ρ ≤ 0.04", "Cl 10.7.1 (laps excluded)", "=«rho»", "=0.04", "", '=«rho»/0.04', nf="0.00%")
    # bar strain at capacity vs the fracture strain (AS/NZS 4671 uniform elongation: Class N 5 %, Class L 1.5 %)
    ES_ROW = C0 + 9
    EPS_U = f'IF(COUNTIF({CLS_ARG},"L")>0,0.015,0.05)'
    epsmax = f'MAX({COLS[0]}{ES_ROW}:{COLS[-1]}{ES_ROW})'
    chk(24, "εs ≤ εsu", f'=IF(COUNT({COLS[0]}{ES_ROW}:{COLS[-1]}{ES_ROW})=0,"",IF(COUNTIF({CLS_ARG},"L")>0,"Class L 1.5 %","Class N 5 %")&" fracture · max at "&INDEX($B${LC0}:$B${LC0 + NLC - 1},MATCH({epsmax},{COLS[0]}{ES_ROW}:{COLS[-1]}{ES_ROW},0)))',
        f'=IF(COUNT({COLS[0]}{ES_ROW}:{COLS[-1]}{ES_ROW})=0,"",{epsmax})', f"={EPS_U}", "",
        f'=IF(E24="","",E24/F24)', nf="0.00%")
    heading(ws, 25, "Slenderness (information)", sz=8)
    for rr, ax, sl in ((26, "x", "«slx»"), (27, "y", "«sly»")):
        D0 = DX0 if ax == "x" else DY0
        chk(rr, f"Le/r ({ax})", f'=IF(COUNT($J${D0}:$J${D0 + NLC - 1})=0,"",IF(MAX($J${D0}:$J${D0 + NLC - 1})>1,"slender: δb up to "&TEXT(MAX($J${D0}:$J${D0 + NLC - 1}),"0.00")&" applied","short for all cases"))',
            f"={sl}", f'=IF(COUNT($H${D0}:$H${D0 + NLC - 1})=0,"",MIN($H${D0}:$H${D0 + NLC - 1}))', "", None, status="info")
        ws[f"J{rr}"].font = font(8, i=True, color=GREY)
    status_cf(ws, f"J{CHK0}:K{CHK1}"); util_cf(ws, f"I{CHK0}:I{CHK1}", f"I{CHK0}")

    # ================================================================== resolve, arrays, validation, grey-outs, print
    for ref, f in pend: ws[ref].value = Fx(f)
    for ref in arrays:
        top = ref.split(":")[0]; ws[top].value = ArrayFormula(ref, ws[top].value)
    def dvl(formula, ref):
        v = DataValidation(type="list", formula1=formula, allow_blank=True); ws.add_data_validation(v); v.add(ref)
    dvl('"Rectangular,Circular,Custom polygon"', "E39"); dvl('"Rectangular,Curvilinear,Parabolic"', "E43"); dvl('"Cl 10.6.4,Strain compatibility"', "E46")
    dvl("=BarList", "E50:E51"); dvl("=BarList", f"G{CUS0}:G{CUS1}")
    for ref, cond, labref in (("E40", "«POLY»", "B40"), ("E41", "OR(«CIRC»,«POLY»)", "B41"), ("E52", "«POLY»", "B52"),
                              ("E53", "OR(«CIRC»,«POLY»)", "B53"), ("E46:F46", '«shape»<>"Rectangular"', "B46"),
                              ("E77:H88", "NOT(«POLY»)", "B77:D88")):
        f = Fx(cond)
        ws.conditional_formatting.add(ref, FormulaRule(formula=[f], fill=PatternFill(bgColor=NA_FILL), font=Font(color=NA_FONT)))
        ws.conditional_formatting.add(labref, FormulaRule(formula=[f], font=Font(color=NA_FONT)))
    for rr in range(15, LAST + 1):
        if ws.row_dimensions[rr].height is None: ws.row_dimensions[rr].height = 14
    ws.print_area = f"A1:K{LAST}"
    for b_ in (62, 119, BS0 - 4): ws.row_breaks.append(Break(id=b_))
    return ws


# ---------------- examples ----------------
# C1: Brent's previous column sheet example (Port Macquarie capping beam, L1): 300 x 600, f'c 40, 2N24 top/bottom + 1N24 sides,
#     cover 50, R10 ligs, H 3900, k 1.0, G/(G+Q) 0.8, N* 1993.6 kN.
c1 = dict(loc="Port Macquarie – capping beam col L1 (previous sheet)", shape="Rectangular", b=300, D=600, fc=40, cover=50, lig="R10",
          bar="N24", n1=2, n2=3, Lu=3900, kx=1.0, ky=1.0, betad=0.8, phi0=0.6, method="Cl 10.6.4",
          lcs=[("previous sheet", 1993.6, 1, 0, 0, 0), ("1.2G + 1.5Q, biaxial", 1500, 180, -90, 40, 40), ("wind", 900, 250, 120, 15, -15)],
          custom=[])
# C2: Ø600 pile/column like Brent's pile calculator: 6N20, N12 ligs, cover 50, f'c 32 (+ 2 custom N16 to show the feature)
c2 = dict(loc="Ø600 bored pier (pile calculator check)", shape="Circular", b=600, D=None, fc=32, cover=50, lig="N12",
          bar="N20", n1=6, n2=None, Lu=3000, kx=1.0, ky=1.0, betad=0.6, phi0=0.6, method="Strain compatibility",
          lcs=[("bending only", 0, 150, 150, 0, 0), ("1.2G + 1.5Q", 1500, 120, -60, 60, 30)],
          custom=[])
# C3: custom outline – 1000 wide x 200 deep, soffit curving up to 150 deep at mid-width; bars follow the curve (illustrative actions)
c3 = dict(loc="Custom outline – 1000 × 200 blade, curved soffit to 150 at mid-width", shape="Custom polygon", b=None, D=None, fc=40, cover=40,
          lig="N10", bar="N16", n1=None, n2=None, Lu=3000, kx=1.0, ky=1.0, betad=0.6, phi0=0.6, method="Strain compatibility",
          corners=[(0, 0, -50, 3), (1000, 0, 0, 0), (1000, 200, 0, 3), (0, 200, 0, 0)],
          lcs=[("1.2G + 1.5Q", 600, 20, -10, 0, 0), ("1.2G + Wu", 350, 35, 35, 10, -10), ("0.9G + Wu", 250, -35, -35, -10, 10)],
          custom=[])
if __name__ == "__main__":
    for nm_, dd in (("C1", c1), ("C2", c2), ("C3", c3)): build_tab(nm_, dd)
    wb._sheets = [wb[n] for n in ("C1", "C2", "C3", "Bars")]
    wb.active = 0
    from openpyxl.workbook.properties import CalcProperties
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    wb.save(OUT)
    print("saved", OUT)
