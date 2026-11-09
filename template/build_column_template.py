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
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, RichTextProperties

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
D_ = wb.create_sheet("Bars")
setup_data_tab(D_, "Reinforcement – bar properties",
               "AS/NZS 4671 nominal values. S = 250 MPa deformed (D250N), R = plain round, L = mesh wire (Class L). Edit or add rows; the calc tabs read this list.",
               {"B": 10, "C": 10, "D": 10, "E": 10, "F": 10, "G": 40})
for j, t in enumerate(["Bar", "db (mm)", "Ab (mm²)", "fsy (MPa)", "Class", "Note"]):
    c = D_.cell(4, 2 + j); c.value = t; c.font = font(8, True); c.fill = fill(HDR_FILL); c.border = BOX; c.alignment = CEN
BAR0 = 5; BAR_LAST = BAR0 + len(BARS) + 5
for i in range(BAR_LAST - BAR0 + 1):
    row = BARS[i] + [BAR_NOTES.get(BARS[i][0], "")] if i < len(BARS) else [None] * 6
    for j, v in enumerate(row):
        c = D_.cell(BAR0 + i, 2 + j); c.value = v; c.font = font(8, color="000000"); c.fill = fill(IN_FILL); c.border = BOX
        c.alignment = LEFT if j in (0, 5) else CEN
for j, t in enumerate(["f'c (MPa)", "Ec (MPa)"]):
    c = D_.cell(4, 9 + j); c.value = t; c.font = font(8, True); c.fill = fill(HDR_FILL); c.border = BOX; c.alignment = CEN
for i, (fc_, ec_) in enumerate(zip([20, 25, 32, 40, 50, 65, 80, 100], [24000, 26700, 30100, 32800, 34800, 37400, 39600, 42200])):
    for j, v in enumerate((fc_, ec_)):
        c = D_.cell(5 + i, 9 + j); c.value = v; c.font = font(8, color="000000"); c.fill = fill(IN_FILL); c.border = BOX; c.alignment = CEN
c = D_.cell(13, 9); c.value = "AS 3600 Table 3.1.2 (ρ = 2400 kg/m³); calc tabs interpolate linearly."; c.font = font(7, i=True, color=GREY)
D_.column_dimensions["H"].width = 3; D_.column_dimensions["I"].width = 10; D_.column_dimensions["J"].width = 10
wb.defined_names["BarList"] = DefinedName("BarList", attr_text=f"Bars!$B${BAR0}:$B${BAR_LAST}")
wb.defined_names["EcFc"] = DefinedName("EcFc", attr_text="Bars!$I$5:$I$12")
wb.defined_names["EcVal"] = DefinedName("EcVal", attr_text="Bars!$J$5:$J$12")
BARTAB = f"Bars!$B${BAR0}:$F${BAR_LAST}"
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
        s = s.replace("«CIRC»", '«shape»="Circular"').replace("«SHAPE»", 'IF(«shape»="Circular","Circle","Rect")')
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
    heading(ws, 28, "Interaction Curve – Key Points (design capacities, uniaxial)")
    table_header(ws, 29, [("B29:C29", "Point"), ("D29", "Definition"), ("E29", "φN (kN)"), ("F29", "φMx (kNm)"),
                          ("G29", "φN (kN)"), ("H29", "φMy (kNm)"), ("I29", "φ  x"), ("J29", "φ  y")])
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
        rr = 30 + k
        ws.merge_cells(f"B{rr}:C{rr}"); lab(f"B{rr}", pt); desc(rr, ds)
        for a, cN, cM, cP in (("x", "E", "F", "I"), ("y", "G", "H", "J")):
            fN, fM, fP = fn(a)
            CA(f"{cN}{rr}", f'=IFERROR({fN},"")', "0"); CA(f"{cM}{rr}", f'=IFERROR({fM},"")', "0.0"); CA(f"{cP}{rr}", f'=IFERROR({fP},"")', "0.00")
    note(ws, 35, "Uniaxial points from OAR.CONCRETE.COLUMNAXIS (calculation page). N compression +. φ from Table 2.2.2: φ0 at and above Nub, φb at N = 0, Class N in tension → 0.85.")

    # ================================================================== INPUTS (page 1)
    heading(ws, 37, "Section and Concrete")
    lab("B38", "Shape"); IN("E38:F38", d["shape"], name="shape")
    F("B39", '=IF(«CIRC»,"Diameter D","Width b (x direction)")'); ws["B39"].font = font(9)
    IN("E39", d["b"], "0", name="b"); U("F39", "mm")
    lab("B40", "Depth D (y direction)"); IN("E40", d.get("D"), "0", name="D"); U("F40", "mm"); small("H40", "rectangular only")
    lab("B41", "Concrete strength f'c"); IN("E41", d["fc"], "0", name="fc"); U("F41", "MPa")
    lab("B42", "Stress block (ULS)"); IN("E42:F42", d.get("model", "Rectangular"), name="model")
    lab("B43", "Ec override"); IN("E43", None, "0", name="ecin"); U("F43", "MPa"); small("H43", "blank = Table 3.1.2")
    lab("B44", "φ0 for compression, Table 2.2.2(d)"); IN("E44", d.get("phi0", 0.6), "0.00", name="phi0"); small("H44", "Amdt 2: φ0 = 0.65 kφ – enter the value")
    lab("B45", "Biaxial check governing"); IN("E45:F45", d.get("method", "Cl 10.6.4"), name="method"); small("H45", "circular: strain compatibility")

    heading(ws, 47, "Reinforcement – typical layout")
    lab("B48", "Cover to ligature"); IN("E48", d["cover"], "0", name="cover"); U("F48", "mm")
    lab("B49", "Ligature bar"); IN("E49", d["lig"], name="lig")
    lab("B50", "Main bar (typical)"); IN("E50", d["bar"], name="bar"); small("H50", '"None" = custom bars only')
    F("B51", '=IF(«CIRC»,"Number of bars","Bars per b face (incl. corners)")'); ws["B51"].font = font(9)
    IN("E51", d["n1"], "0", name="n1")
    lab("B52", "Bars per D face (incl. corners)"); IN("E52", d.get("n2"), "0", name="n2"); small("H52", "rectangular only")
    note(ws, 53, "Custom bars: x, y from the section centre and bar size in the Bar Schedule (last page). They are added to the typical layout.")

    heading(ws, 55, "Slenderness – braced column (Cl 10.3, 10.4)")
    table_header(ws, 56, [("E56", "x axis (Mx)"), ("F56", "y axis (My)")])
    lab("B57", "Unsupported length Lu"); IN("E57", d["Lu"], "0", name="Lu"); U("G57", "mm")
    lab("B58", "Effective length factor k"); IN("E58", d.get("kx", 1.0), "0.00", name="kx"); IN("F58", d.get("ky", 1.0), "0.00", name="ky")
    lab("B59", "βd = G / (G + Q), sustained axial load"); IN("E59", d.get("betad", 0.6), "0.00", name="betad")
    note(ws, 60, "Unbraced (sway) columns are not covered. βd is taken as 0 where Le/r ≤ 40 and N* ≤ M2*/(2D) (Cl 10.4.3).")

    heading(ws, 62, "Load Cases (ULS)")
    table_header(ws, 63, [("B63:C63", "Case"), ("D63", "Description"), ("E63", "N* (kN)"), ("F63", "M*x top"), ("G63", "M*x bot"),
                          ("H63", "M*y top"), ("I63", "M*y bot")])
    LC0 = 64
    for i in range(NLC):
        r = LC0 + i; lc = d["lcs"][i] if i < len(d["lcs"]) else None
        IN(f"B{r}:C{r}", f"LC{i + 1}")
        IN(f"D{r}", lc[0] if lc else None, al=LEFT)
        for j, col in enumerate("EFGHI"): IN(f"{col}{r}", lc[1 + j] if lc else None, "0.0")
    small("J63", "kNm")
    note(ws, 70, "N* compression +. End moments: same sign = single curvature. Blank N* = case not used. Minimum eccentricity 0.05D applies about each axis (Cl 10.1.2).")
    lab("B71", "Moment contour plotted for case"); IN("E71", d.get("contour", 1), "0", name="lcsel"); small("F71", "1 – 6")

    # ================================================================== SECTION + BAR SCHEDULE (helpers first; formulas reference them)
    # bar schedule rows
    BS0 = 190
    TYP0, TYP1 = BS0 + 1, BS0 + NTYP
    CUS0, CUS1 = TYP1 + 1, TYP1 + NCUS
    BARS_ARG = f"$L${TYP0}:$O${CUS1}"
    CLS_ARG = f"$J${TYP0}:$J${CUS1}"
    SEC = '«SHAPE»,«b»:«D»'
    MOD = '«fc»,«model»'

    # ================================================================== CALCULATIONS (page 2)
    heading(ws, 73, "Section Properties")
    table_header(ws, 74, [("E74", "x axis (Mx)"), ("F74", "y axis (My)")])
    r = 75; sym(r, "Ag"); desc(r, '=IF(«CIRC»,"π D² / 4","b D")'); CA(f"E{r}", '=IF(«CIRC»,PI()*«b»^2/4,«b»*«D»)', "0", name="Ag"); U(f"G{r}", "mm²")
    r = 76; sym(r, "As"); desc(r, "all bars (Bar Schedule)"); CA(f"E{r}", f"=SUM($L${TYP0}:$L${CUS1})", "0", name="As"); U(f"G{r}", "mm²")
    r = 77; sym(r, "ρ"); desc(r, "As / Ag"); CA(f"E{r}", "=«As»/«Ag»", "0.00%", name="rho")
    r = 78; sym(r, "n"); desc(r, "typical | custom bars")
    CA(f"E{r}", f'=COUNTIFS($G${TYP0}:$G${TYP1},"<>",$H${TYP0}:$H${TYP1},">0")', "0", name="ntyp")
    CA(f"F{r}", f'=COUNTIF($H${CUS0}:$H${CUS1},">0")', "0", name="ncus")
    r = 79; sym(r, "edge"); desc(r, "cover + ligature + db / 2")
    CA(f"E{r}", f'=«cover»+IFERROR({barv("«lig»", 2)},0)+IFERROR({barv("«bar»", 2)},0)/2', "0.0", name="edge"); U(f"G{r}", "mm")
    r = 80; sym(r, "Dd"); desc(r, "section depth in the bending direction")
    CA(f"E{r}", '=IF(«CIRC»,«b»,«D»)', "0", name="Dx"); CA(f"F{r}", '=«b»', "0", name="Dy"); U(f"G{r}", "mm")
    r = 81; sym(r, "r"); desc(r, '=IF(«CIRC»,"0.25 D (Cl 10.5.2)","0.3 × depth (Cl 10.5.2)")')
    CA(f"E{r}", '=IF(«CIRC»,0.25,0.3)*«Dx»', "0.0", name="rx"); CA(f"F{r}", '=IF(«CIRC»,0.25,0.3)*«Dy»', "0.0", name="ry"); U(f"G{r}", "mm")
    r = 82; sym(r, "Le"); desc(r, "k Lu"); CA(f"E{r}", "=«kx»*«Lu»", "0", name="Lex"); CA(f"F{r}", "=«ky»*«Lu»", "0", name="Ley"); U(f"G{r}", "mm")
    r = 83; sym(r, "Le / r"); desc(r, "slenderness ratio"); CA(f"E{r}", "=«Lex»/«rx»", "0.0", name="slx"); CA(f"F{r}", "=«Ley»/«ry»", "0.0", name="sly")
    r = 84; sym(r, "Ec"); desc(r, '=IF(N(«ecin»)>0,"entered","Table 3.1.2, interpolated")')
    I_ = "MIN(MATCH(«fc»,EcFc,1),7)"
    CA(f"E{r}", f"=IF(N(«ecin»)>0,«ecin»,INDEX(EcVal,{I_})+(«fc»-INDEX(EcFc,{I_}))*(INDEX(EcVal,{I_}+1)-INDEX(EcVal,{I_}))/(INDEX(EcFc,{I_}+1)-INDEX(EcFc,{I_})))", "0", name="Ec")
    U(f"G{r}", "MPa")

    heading(ws, 86, "Axis Properties (OAR Concrete add-in)")
    note(ws, 87, "OAR.CONCRETE.COLUMNAXIS: strain compatibility, εcu 0.003. Nuo = α1 f'c Ac + fsy As (Ac net of bars, Cl 10.6.2.2); balanced point with the outermost tension bar at yield.")
    table_header(ws, 88, [("E88", "x axis (Mx)"), ("F88", "y axis (My)")])
    AX = [("Nuo", "squash load", "0", "kN"), ("Nuot", "pure tension", "0", "kN"), ("Nub", "balanced point", "0", "kN"),
          ("Mub", "moment at Nub (Mc, Cl 10.4.4)", "0.0", "kNm"), ("Muo", "pure bending", "0.0", "kNm"), ("kuo", "at pure bending", "0.000", ""),
          ("φb", "Tbl 2.2.2, pure bending", "0.000", ""), ("do", "to outermost tension bar", "0.0", "mm"),
          ("Nud", "decompression (dn = do)", "0", "kN"), ("Mud", "moment at Nud", "0.0", "kNm"), ("φt", "Tbl 2.2.2, pure tension", "0.000", "")]
    AX0 = 89
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
    heading(ws, 102, "Design Moments – minimum eccentricity and moment magnification (braced)")
    note(ws, 103, "M2* larger end moment; M1*/M2* negative in single curvature, −1 if M2* ≤ 0.05 D N*. Short if Le/r ≤ max(25, αc (38 − f'c/15)(1 + M1*/M2*)), αc from N*/(0.6 Nuo).")
    note(ws, 104, "Nc = π²/Le² × 182 do φ0 Mc / (1 + βd), Mc = Mub.  δb = km / (1 − N*/Nc) ≥ 1, km = 0.6 − 0.4 M1*/M2* ≥ 0.4 (δb = 1 if short).  M*d = δb × max(M2*, 0.05 D N*).")
    DX0 = dm_table(105, "x", "F", "G", "«Dx»", "«slx»", "«Lex»", "«Nuox»", "«Mubx»", "«dox»")
    DY0 = dm_table(113, "y", "H", "I", "«Dy»", "«sly»", "«Ley»", "«Nuoy»", "«Muby»", "«doy»")

    # ---- capacity along the load direction (strain compatibility, biaxial)
    CP = 122
    heading(ws, CP, "Capacity – strain compatibility along the load direction (OAR Concrete add-in)")
    note(ws, CP + 1, "OAR.CONCRETE.COLUMN: finds the neutral-axis direction θ for which the capacity (Mux, Muy) at N* points along (M*x,d, M*y,d); φ from Tbl 2.2.2 (φ0 … φb by N*/Nub).")
    h = CP + 2
    table_header(ws, h, [(f"{c}{h}", "") for c in "EFGHIJ"])
    for i, c in enumerate("EFGHIJ"): F(f"{c}{h}", f"=$B${LC0 + i}")
    COLS = "EFGHIJ"
    CPR = [("φ", "Tbl 2.2.2 at N*", "0.000", ""), ("Mu", "capacity, load direction", "0.0", "kNm"), ("φMu", "", "0.0", "kNm"),
           ("Mux", "component about x", "0.0", "kNm"), ("Muy", "component about y", "0.0", "kNm"), ("θ", "max compression direction", "0.0", "deg"),
           ("dn", "neutral axis depth", "0.0", "mm"), ("Nub", "balanced, this direction", "0", "kN"), ("Nuo", "squash", "0", "kN")]
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

    # ================================================================== CHARTS (page 3) and helper data (columns M onward, not printed)
    CH = CALC_END + 2
    heading(ws, CH, "Interaction Diagrams")
    HR = 300  # helper rows start (below everything, outside the print area)
    lab(f"M{HR - 1}", "Chart data (not printed)", sz=7, color=GREY, i=True)
    # M–N curves: 40 rows x 5 (N, M, φ, φN, φM)
    for ax, col in (("x", "M"), ("y", "S")):
        c0 = col; c4 = chr(ord(col) + 4)
        lab(f"{c0}{HR}", f"M–N {ax}: N | M | φ | φN | φM", sz=7, color=GREY)
        ws[f"{c0}{HR + 1}"].value = f'=OAR.CONCRETE.MNCOLUMN({SEC},{BARS_ARG},{MOD},"{ax}",«phi0»,{CLS_ARG},«Ec»)'
        pend.append((f"{c0}{HR + 1}", ws[f"{c0}{HR + 1}"].value)); arrays.append(f"{c0}{HR + 1}:{c4}{HR + 40}")
    # load points: |M*d| vs N* per axis (NA() when unused so the charts skip them)
    lab(f"Y{HR}", "loads x: M | N   y: M | N", sz=7, color=GREY)
    for i in range(NLC):
        rr = HR + 1 + i; lr = LC0 + i
        for col, src in (("Y", f"ABS($K${DX0 + i})"), ("Z", f"$E${lr}"), ("AA", f"ABS($K${DY0 + i})"), ("AB", f"$E${lr}")):
            pend.append((f"{col}{rr}", f'=IF($E${lr}="",NA(),{src})'))
    # contour at the selected case
    lab(f"AD{HR}", "contour: Mx | My | φMx | φMy", sz=7, color=GREY)
    NSEL = f"INDEX($E${LC0}:$E${LC0 + NLC - 1},«lcsel»)"
    pend.append((f"AD{HR + 1}", f'=IF({NSEL}="","",OAR.CONCRETE.MCONTOUR({SEC},{BARS_ARG},{MOD},{NSEL},36,«phi0»,{CLS_ARG},«Ec»))'))
    arrays.append(f"AD{HR + 1}:AG{HR + 37}")
    lab(f"AI{HR}", "contour load: Mx | My", sz=7, color=GREY)
    pend.append((f"AI{HR + 1}", f'=IFERROR(INDEX($K${DX0}:$K${DX0 + NLC - 1},«lcsel»)*1,NA())'))
    pend.append((f"AJ{HR + 1}", f'=IFERROR(INDEX($K${DY0}:$K${DY0 + NLC - 1},«lcsel»)*1,NA())'))
    # section outline: rectangle (5 points) or circle (73 points), mm from centre
    lab(f"AL{HR}", "outline: x | y", sz=7, color=GREY)
    for k in range(73):
        rr = HR + 1 + k
        a = f"(PI()/2-2*PI()*{k}/72)"
        rect_pts = [(-1, 1), (1, 1), (1, -1), (-1, -1), (-1, 1)]
        if k < 5:
            sx, sy = rect_pts[k]
            fx = f'=IF(«CIRC»,«b»/2*COS({a}),{sx}*«b»/2)'; fy = f'=IF(«CIRC»,«b»/2*SIN({a}),{sy}*«D»/2)'
        else:
            fx = f'=IF(«CIRC»,«b»/2*COS({a}),NA())'; fy = f'=IF(«CIRC»,«b»/2*SIN({a}),NA())'
        pend.append((f"AL{rr}", fx)); pend.append((f"AM{rr}", fy))

    # ================================================================== BAR SCHEDULE (last page)
    heading(ws, BS0 - 3, "Bar Schedule")
    note(ws, BS0 - 2, "Typical bars from OAR.CONCRETE.BARLAYOUT (rectangular: perimeter, top face left → right, then clockwise; circular: equally spaced, first at the top). Custom bars: enter x, y and bar.")
    table_header(ws, BS0 - 1, [(f"B{BS0 - 1}:C{BS0 - 1}", "#"), (f"D{BS0 - 1}", "Source"), (f"E{BS0 - 1}", "x (mm)"), (f"F{BS0 - 1}", "y (mm)"),
                               (f"G{BS0 - 1}", "Bar"), (f"H{BS0 - 1}", "As (mm²)"), (f"I{BS0 - 1}", "fsy (MPa)"), (f"J{BS0 - 1}:K{BS0 - 1}", "Class")])
    ws.row_dimensions[BS0].height = 6
    lab(f"L{BS0 - 1}", "engine: As | depth | fsy | x from left", sz=7, color=GREY)
    for k in range(NTYP + NCUS):
        rr = TYP0 + k; typ = k < NTYP
        ws.merge_cells(f"B{rr}:C{rr}"); ws[f"B{rr}"].value = k + 1; ws[f"B{rr}"].font = font(8, color=GREY); ws[f"B{rr}"].alignment = CEN
        ws[f"D{rr}"].value = "typical" if typ else "custom"; ws[f"D{rr}"].font = font(8, color=GREY)
        if typ:
            calc(ws, f"E{rr}", None, "0.0"); calc(ws, f"F{rr}", None, "0.0")
            CA(f"G{rr}", f'=IF(E{rr}="","",«bar»)')
        else:
            c = d["custom"][k - NTYP] if k - NTYP < len(d.get("custom", [])) else (None, None, None)
            IN(f"E{rr}", c[0], "0.0"); IN(f"F{rr}", c[1], "0.0"); IN(f"G{rr}", c[2])
        CA(f"H{rr}", f'=IF(OR(G{rr}="",G{rr}="None"),"",{barv(f"G{rr}", 3)})', "0")
        CA(f"I{rr}", f'=IF(H{rr}="","",{barv(f"G{rr}", 4)})', "0")
        CA(f"J{rr}:K{rr}", f'=IF(H{rr}="","",{barv(f"G{rr}", 5)})')
        # engine columns
        for col, f in (("L", f"=N(H{rr})"), ("M", f'=IF(L{rr}=0,"",IF(«CIRC»,«b»,«D»)/2-F{rr})'),
                       ("N", f'=IF(L{rr}=0,"",I{rr})'), ("O", f'=IF(L{rr}=0,"",«b»/2+E{rr})'),
                       ("Q", f'=IF(L{rr}=0,NA(),E{rr})'), ("R", f'=IF(L{rr}=0,NA(),F{rr})')):
            pend.append((f"{col}{rr}", f)); ws[f"{col}{rr}"].font = font(7, color=GREY)
    pend.append((f"E{TYP0}", f'=IF(OR(«bar»="",«bar»="None"),"",OAR.CONCRETE.BARLAYOUT({SEC},«edge»,«n1»,«n2»,{NTYP}))'))
    arrays.append(f"E{TYP0}:F{TYP1}")
    LAST = CUS1 + 2
    sym(CUS1 + 1, "Σ"); desc(CUS1 + 1, "total"); CA(f"H{CUS1 + 1}", f"=«As»", "0"); U(f"I{CUS1 + 1}", "mm²")
    note(ws, LAST, "Ligature spacing and size, laps and anchorage are not covered by this calc.")

    # ================================================================== SUMMARY + CHECKS (page 1)
    heading(ws, 8, "Design Summary")
    lab("B9", "Adopted section", sz=7, color=GREY); lab("F9", "Max utilisation", sz=7, color=GREY); lab("I9", "Status", sz=7, color=GREY)
    ws.merge_cells("B10:D10"); ws.merge_cells("F10:G10"); ws.merge_cells("I10:K10")
    F("B10", '=IF(«CIRC»,"Ø"&«b»,«b»&" × "&«D»)&IF(«ntyp»>0,"  ·  "&«ntyp»&«bar»,"")&IF(«ncus»>0," + "&«ncus»&" custom","")&"  ·  f\'c "&«fc»')
    ws["B10"].font = font(12, True, color=GREEN); ws["B10"].alignment = LEFT
    CHK0, CHK1 = 15, 23
    c = ws["F10"]; c.value = f'=IF(COUNT(I{CHK0}:I{CHK1})=0,"–",MAX(I{CHK0}:I{CHK1}))'; c.font = font(12, True, color=GREEN); c.number_format = "0%"; c.alignment = LEFT
    errs = f'COUNTIF(E{AX0}:F{AX1},"#OAR*")+COUNTIF(E{C0}:J{C1},"#OAR*")+COUNTIF(E{TYP0}:E{TYP1},"#OAR*")'
    c = ws["I10"]; c.value = (f'=IF($V$7,"UPDATE ADD-IN",IF({errs}>0,"INPUT ERROR",IF(COUNT(I{CHK0}:I{CHK1})=0,"PENDING",'
                              f'IF(OR(MAX(I{CHK0}:I{CHK1})>1,COUNTIF(J{CHK0}:J{CHK1},"NG")>0),"FAIL","PASS"))))')
    c.font = font(10, True, color=GREY); c.alignment = CEN
    for x in ws["I10:K10"][0]: x.border = BOX
    passfail_cf(ws, "I10:K10", "$I$10")
    ws.conditional_formatting.add("I10:K10", FormulaRule(formula=['$I$10="UPDATE ADD-IN"'], fill=PatternFill(bgColor="C55A11"), font=Font(color="FFFFFF", bold=True)))
    F("B11", '="Lu "&«Lu»&" mm  ·  φ0 "&TEXT(«phi0»,"0.00")&"  ·  "&«model»&" block  ·  biaxial: "&IF(«CIRC»,"strain compatibility",«method»)')
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
        use1064 = 'AND(NOT(«CIRC»),«method»="Cl 10.6.4")'
        chk(rr, f'=$B${lr}&" N–M"', f'=IF({blank},"(not used)",IF({use1064},"Cl 10.6.4, αn = "&TEXT({c}{B0 + 5},"0.00"),"strain compat., load direction")&IF(MAX($J${DX0 + i},$J${DY0 + i})>1,"  ·  δb "&TEXT(MAX($J${DX0 + i},$J${DY0 + i}),"0.00"),""))',
            f'=IF({blank},"",IF({use1064},{c}{B1064},{c}{RM}))', f'=IF({blank},"",IF({use1064},1,{c}{C0 + 2}))',
            f'=IF({use1064},"–","kNm")', f'=IF({blank},"",IF({use1064},{c}{B1064},{c}{RSC}))',
            ratio=f'=IF({blank},"",IF({use1064},{c}{RSC},IF(«CIRC»,"",{c}{B1064})))', nf="0.00")
    chk(22, "ρ ≥ 0.01", '=IF(«rho»>=0.01,"Cl 10.7.1(a)","below 1 %: Cl 10.7.1(b) lightly loaded / AS 2159 piles – engineer to confirm")',
        "=0.01", "=«rho»", "", '=IF(OR(«rho»=0,«rho»<0.01),"",0.01/«rho»)', nf="0.00%",
        status='=IF(«rho»=0,"–",IF(«rho»>=0.01,"OK","OK*"))')
    chk(23, "ρ ≤ 0.04", "Cl 10.7.1 (laps excluded)", "=«rho»", "=0.04", "", '=«rho»/0.04', nf="0.00%")
    heading(ws, 24, "Slenderness (information)", sz=8)
    for rr, ax, sl in ((25, "x", "«slx»"), (26, "y", "«sly»")):
        D0 = DX0 if ax == "x" else DY0
        chk(rr, f"Le/r ({ax})", f'=IF(COUNT($J${D0}:$J${D0 + NLC - 1})=0,"",IF(MAX($J${D0}:$J${D0 + NLC - 1})>1,"slender: δb up to "&TEXT(MAX($J${D0}:$J${D0 + NLC - 1}),"0.00")&" applied","short for all cases"))',
            f"={sl}", f'=IF(COUNT($H${D0}:$H${D0 + NLC - 1})=0,"",MIN($H${D0}:$H${D0 + NLC - 1}))', "", None, status="info")
        ws[f"J{rr}"].font = font(8, i=True, color=GREY)
    status_cf(ws, f"J{CHK0}:K{CHK1}"); util_cf(ws, f"I{CHK0}:I{CHK1}", f"I{CHK0}")

    # ================================================================== charts
    # Plain engineering plots: no chart title or legend (caption in the cell above), thin lines,
    # light major and minor gridlines, axes at the left / bottom, square border.
    BLUE, ORANGE, DOT = "4472C4", "ED7D31", "FF0000"
    def thin(color, w=9525): return GraphicalProperties(ln=LineProperties(solidFill=color, w=w))
    def tick_font(sz=800):
        return RichText(bodyPr=RichTextProperties(), p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=sz, solidFill="404040")), endParaRPr=CharacterProperties())])
    def line_series(xref, yref, color, width=15875, marker=False):
        s = Series(yref, xref); s.smooth = False
        if marker:
            s.marker.symbol = "circle"; s.marker.size = 5
            s.marker.graphicalProperties = GraphicalProperties(solidFill=color, ln=LineProperties(solidFill=color))
            s.graphicalProperties.line = LineProperties(noFill=True)
        else:
            s.marker.symbol = "none"
            s.graphicalProperties.line = LineProperties(solidFill=color, w=width)
        return s
    def chart(xmin=None, w=9.6, h=6.9):
        c = ScatterChart(); c.style = 2; c.height = h; c.width = w
        c.title = None; c.legend = None; c.roundedCorners = False
        c.graphical_properties = GraphicalProperties(ln=LineProperties(solidFill="000000", w=9525))
        for a in (c.x_axis, c.y_axis):
            a.delete = False; a.title = None
            a.majorGridlines = ChartLines(spPr=thin("D9D9D9"))
            a.minorGridlines = ChartLines(spPr=thin("F2F2F2"))
            a.spPr = thin("808080"); a.txPr = tick_font()
            a.crosses = "min"; a.tickLblPos = "low"
            a.majorTickMark = "none"; a.minorTickMark = "none"
            a.number_format = "General"
        if xmin is not None: c.x_axis.scaling.min = xmin
        return c
    def caption(cell, text):
        ws[cell] = text; ws[cell].font = font(9, b=True, color=GREEN)
    note(ws, CH + 1, "Blue = unfactored capacity, orange = φ capacity, red = design actions. Contour at N* of the case selected in the inputs.")
    for ax, col, cap, anchor in (("x", "M", "B", "B"), ("y", "S", "G", "G")):
        cN = col; cM = chr(ord(col) + 1); cPN = chr(ord(col) + 3); cPM = chr(ord(col) + 4)
        caption(f"{cap}{CH + 2}", f"M–N about {ax}:  N (kN) vs M{ax} (kNm)")
        ch = chart(xmin=0)
        ch.series.append(line_series(Reference(ws, range_string=f"'{name}'!${cM}${HR + 1}:${cM}${HR + 40}"),
                                     Reference(ws, range_string=f"'{name}'!${cN}${HR + 1}:${cN}${HR + 40}"), BLUE))
        ch.series.append(line_series(Reference(ws, range_string=f"'{name}'!${cPM}${HR + 1}:${cPM}${HR + 40}"),
                                     Reference(ws, range_string=f"'{name}'!${cPN}${HR + 1}:${cPN}${HR + 40}"), ORANGE))
        lx, ly = ("Y", "Z") if ax == "x" else ("AA", "AB")
        ch.series.append(line_series(Reference(ws, range_string=f"'{name}'!${lx}${HR + 1}:${lx}${HR + NLC}"),
                                     Reference(ws, range_string=f"'{name}'!${ly}${HR + 1}:${ly}${HR + NLC}"), DOT, marker=True))
        ws.add_chart(ch, f"{anchor}{CH + 3}")
    caption(f"B{CH + 18}", "Moment contour at N*:  My vs Mx (kNm)")
    ch = chart()
    ch.series.append(line_series(Reference(ws, range_string=f"'{name}'!$AD${HR + 1}:$AD${HR + 37}"),
                                 Reference(ws, range_string=f"'{name}'!$AE${HR + 1}:$AE${HR + 37}"), BLUE))
    ch.series.append(line_series(Reference(ws, range_string=f"'{name}'!$AF${HR + 1}:$AF${HR + 37}"),
                                 Reference(ws, range_string=f"'{name}'!$AG${HR + 1}:$AG${HR + 37}"), ORANGE))
    ch.series.append(line_series(Reference(ws, range_string=f"'{name}'!$AI${HR + 1}:$AI${HR + 1}"),
                                 Reference(ws, range_string=f"'{name}'!$AJ${HR + 1}:$AJ${HR + 1}"), DOT, marker=True))
    ws.add_chart(ch, f"B{CH + 19}")
    caption(f"G{CH + 18}", "Section:  y vs x (mm from centre)")
    ch = chart()
    ch.series.append(line_series(Reference(ws, range_string=f"'{name}'!$AL${HR + 1}:$AL${HR + 73}"),
                                 Reference(ws, range_string=f"'{name}'!$AM${HR + 1}:$AM${HR + 73}"), "404040"))
    ch.series.append(line_series(Reference(ws, range_string=f"'{name}'!$Q${TYP0}:$Q${CUS1}"),
                                 Reference(ws, range_string=f"'{name}'!$R${TYP0}:$R${CUS1}"), GREEN, marker=True))
    ch.series[-1].marker.size = 6
    ws.add_chart(ch, f"G{CH + 19}")

    # ================================================================== resolve, arrays, validation, grey-outs, print
    for ref, f in pend: ws[ref].value = Fx(f)
    for ref in arrays:
        top = ref.split(":")[0]; ws[top].value = ArrayFormula(ref, ws[top].value)
    def dvl(formula, ref):
        v = DataValidation(type="list", formula1=formula, allow_blank=True); ws.add_data_validation(v); v.add(ref)
    dvl('"Rectangular,Circular"', "E38"); dvl('"Rectangular,Curvilinear,Parabolic"', "E42"); dvl('"Cl 10.6.4,Strain compatibility"', "E45")
    dvl("=BarList", "E49:E50"); dvl("=BarList", f"G{CUS0}:G{CUS1}")
    for ref, cond, labref in (("E40", "«CIRC»", "B40"), ("E52", "«CIRC»", "B52"), ("E45:F45", "«CIRC»", "B45")):
        f = Fx(cond)
        ws.conditional_formatting.add(ref, FormulaRule(formula=[f], fill=PatternFill(bgColor=NA_FILL), font=Font(color=NA_FONT)))
        ws.conditional_formatting.add(labref, FormulaRule(formula=[f], font=Font(color=NA_FONT)))
    for rr in range(15, LAST + 1):
        if ws.row_dimensions[rr].height is None: ws.row_dimensions[rr].height = 14
    ws.print_area = f"A1:K{LAST}"
    for b_ in (61, 101, CALC_END + 1, BS0 - 4): ws.row_breaks.append(Break(id=b_))
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
if __name__ == "__main__":
    for nm_, dd in (("C1", c1), ("C2", c2)): build_tab(nm_, dd)
    wb._sheets = [wb[n] for n in ("C1", "C2", "Bars")]
    wb.active = 0
    from openpyxl.workbook.properties import CalcProperties
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    wb.save(OUT)
    print("saved", OUT)
