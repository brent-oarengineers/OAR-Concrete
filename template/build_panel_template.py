"""Builds template/OAR - Tilt Panel Design.xlsx (OAR house calc style).

Run from the repo root:  python template/build_panel_template.py
One calc tab = one panel design strip: a solid panel (b = 1000) or a jamb strip beside an opening (own design
width, wind and roof tributary widths, extra G at the head). In-service design only (wind, roof axial, earthquake
face load); lifting and temporary bracing are by the contractor's lifting designer.

Supports: base pinned (doweled, or tied to the slab but designed as pinned); top either continuous (wind beam /
raking angle to the roof diaphragm) or at the panel joints only (rafters) - then the top band is checked as a
horizontal beam between the joints and its deflection added (approximate; RFEM for anything unusual).
Optional parapet above the top support. Actions: calculated on the sheet, or entered from RFEM (Manual).

Slenderness: AS 3600-2018 braced moment magnifier (Cl 10.4) on the vertical strip, km = 1.0 for lateral load,
βd per load case from the permanent share of N*, Nc from Mc = Mub. Capacity at N* by strain compatibility
(OAR.CONCRETE.COLUMN on the strip section). Every code check is a visible worksheet formula; the add-in supplies
the section analysis only (COLUMNAXIS, COLUMN, ULT, CRACKED - no new engine functions).
"""
import sys, os, re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
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

NLC = 6                      # ULS load cases (+ 1 SLS row)
TOP_OPTS = ("Continuous (wind beam)", "At panel joints (rafters)")
MANUAL = "Manual (RFEM)"


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
        s = (s.replace("«MAN»", f'«mode»="{MANUAL}"').replace("«TWO»", '«lay»="Two layers"')
              .replace("«BAND»", f'AND(«top»="{TOP_OPTS[1]}",NOT(«MAN»))').replace("«MAN»", f'«mode»="{MANUAL}"'))
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
    def grey_hdr(ref, t): lab(ref, t, sz=7, color=GREY, i=True)

    setup_calc_tab(ws, LOGO, '="TILT PANEL (AS 3600-2018)  –  "&IF(G3="","(ref)",G3)&IF(I3="","","  ·  "&I3)', "PANEL", name, "")
    ws["I3"].value = d["loc"]
    location_frame(ws, "Panel")

    # ---- add-in version check (as the other OAR Concrete sheets; V3:V7, not printed)
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
                      '"Copy this tab for additional panels / strips.   ·   Calculated with OAR Concrete "&$V$4))')
    ws.conditional_formatting.add("B7", FormulaRule(formula=['OR($V$7,$V$4="")'], font=Font(color=RED, bold=True, italic=False)))

    # ================================================================== INPUTS (page 1)
    heading(ws, 35, "Panel, Concrete and Analysis")
    lab("B36", "Panel thickness tw"); IN("E36", d["tw"], "0", name="tw"); U("F36", "mm")
    lab("B37", "Concrete strength f'c"); IN("E37", d["fc"], "0", name="fc"); U("F37", "MPa")
    lab("B38", "Stress block (ULS)"); IN("E38:F38", d.get("model", "Rectangular"), name="model")
    lab("B39", "Ec override"); IN("E39", None, "0", name="ecin"); U("F39", "MPa"); small("H39", "blank = Table 3.1.2")
    lab("B40", "φ0 for compression, Table 2.2.2"); IN("E40", d.get("phi0", 0.6), "0.00", name="phi0"); small("H40", "Amdt 2: φ0 = 0.65 kφ – enter the value")
    lab("B41", "Concrete density (incl. reinforcement)"); IN("E41", d.get("rhoc", 25), "0.0", name="rhoc"); U("F41", "kN/m³")
    lab("B42", "Design actions"); IN("E42:F42", d.get("mode", "Calculated"), name="mode"); small("H42", f'"{MANUAL}": openings / point supports')
    lab("B43", "RFEM moments include 2nd-order effects"); IN("E43", d.get("rfem2", "No"), name="rfem2"); small("H43", "Yes = δb taken as 1 (manual only)")

    heading(ws, 44, "Geometry and Supports")
    lab("B45", "Span Hs, base support to top support"); IN("E45", d["Hs"], "0", name="Hs"); U("F45", "mm")
    lab("B46", "Parapet above top support hp"); IN("E46", d.get("hp", 0), "0", name="hp"); U("F46", "mm"); small("H46", "0 = none")
    lab("B47", "Panel width W (top band span)"); IN("E47", d["W"], "0", name="W"); U("F47", "mm"); small("H47", "between joint supports")
    lab("B48", "Top support"); IN("E48:G48", d.get("top", TOP_OPTS[0]), al=LEFT, name="top")
    lab("B49", "Top band depth hb (joint support)"); IN("E49", d.get("hb", 1200), "0", name="hb"); U("F49", "mm"); small("H49", "band acting as a beam")
    lab("B50", "Effective length factor k (vertical)"); IN("E50", d.get("k", 1.0), "0.00", name="k"); small("H50", "pinned base and top: 1.0")
    note(ws, 51, "Base pinned (doweled; panels tied into the slab are also designed as pinned, the tie checked for the base reaction). Lifting and bracing by others.")

    heading(ws, 52, "Design Strip")
    lab("B53", "Strip description"); IN("E53:H53", d.get("strip", "Solid panel"), al=LEFT, name="stripd")
    lab("B54", "Strip design width b"); IN("E54", d.get("b", 1000), "0", name="b"); U("F54", "mm"); small("H54", "solid 1000; jamb: jamb width")
    lab("B55", "Tributary width – wind bt"); IN("E55", d.get("bt", 1000), "0", name="bt"); U("F55", "mm"); small("H55", "jamb: b + ½ each opening")
    lab("B56", "Tributary width – roof load br"); IN("E56", d.get("br", 1000), "0", name="br"); U("F56", "mm")
    lab("B57", "Additional G at head of strip"); IN("E57", d.get("gadd", 0), "0.0", name="gadd"); U("F57", "kN"); small("H57", "lintel / panel over opening")
    note(ws, 58, "Wind and parapet pressure act over bt; roof reactions over br; G at head applied at the top support (conservative for axial). Self-weight over b.")

    heading(ws, 59, "Reinforcement")
    lab("B60", "Layout"); IN("E60:F60", d.get("lay", "Central"), name="lay")
    lab("B61", "Cover to vertical bars (two layers)"); IN("E61", d.get("cover", 40), "0", name="cover"); U("F61", "mm")
    lab("B62", "Vertical bars: bar | spacing (each layer)"); IN("E62", d["vbar"], name="vbar"); IN("F62", d["vs"], "0", name="vs"); U("G62", "mm")
    lab("B63", "Horizontal bars: bar | spacing (each layer)"); IN("E63", d["hbar"], name="hbar"); IN("F63", d["hs"], "0", name="hs"); U("G63", "mm")
    lab("B64", "Extra vertical bars in strip: bar | spacing (each layer)"); IN("E64", d.get("xbar", "None"), name="xbar"); IN("F64", d.get("xs"), "0", name="xs"); U("G64", "mm")
    lab("B65", "Extra horizontal bars in top band: bar | spacing (each layer)"); IN("E65", d.get("hbbar", "None"), name="hbbar"); IN("F65", d.get("hbs"), "0", name="hbs"); U("G65", "mm")
    lab("B66", "Restrained against horizontal shrinkage"); IN("E66", d.get("hrest", "No"), name="hrest"); small("H66", "No = free to move (panel joints)")
    note(ws, 67, "SL92 = L9-200. Extras (e.g. N12-300) spread over b (band: over hb). Two layers: same bars each face, horizontals inside. Class L → φ 0.65.")

    heading(ws, 68, "Roof Loads at the Top Support (unfactored, per metre of wall)")
    lab("B69", "Roof G reaction"); IN("E69", d.get("Gr", 0), "0.00", name="Gr"); U("F69", "kN/m")
    lab("B70", "Roof Q reaction"); IN("E70", d.get("Qr", 0), "0.00", name="Qr"); U("F70", "kN/m")
    lab("B71", "Roof wind reaction Wu (+ down, − uplift)"); IN("E71", d.get("Wr", 0), "0.00", name="Wr"); U("F71", "kN/m")
    lab("B72", "Eccentricity e from panel centreline"); IN("E72", d.get("ecc", 0), "0", name="ecc"); U("F72", "mm"); small("H72", "+ towards inside face")
    lab("B73", "Earthquake face load coefficient Fc / Wc"); IN("E73", d.get("ceq"), "0.000", name="ceq")
    lab("G73", "pE =", al=RIGHT); CA("H73", "=N(«ceq»)*«rhoc»*«tw»/1000", "0.00", name="pE"); U("I73", "kPa"); small("J73", "AS 1170.4 Sec 8")

    heading(ws, 74, "Wind – AS/NZS 1170.2 net pressures on the panel and parapet")
    lab("B75", "Design wind pressure qz: ULS | SLS (panel)"); IN("E75", d.get("qzu"), "0.00", name="qzu"); IN("F75", d.get("qzs"), "0.00", name="qzs"); U("G75", "kPa")
    lab("B76", "qz at parapet top: ULS | SLS"); IN("E76", d.get("qzpu"), "0.00", name="qzpu"); IN("F76", d.get("qzps"), "0.00", name="qzps"); U("G76", "kPa"); small("H76", "blank = panel qz")
    table_header(ws, 77, [("E77", "Inward"), ("F77", "Outward")])
    WIN = [("78", "Cp,e (+ towards the surface)", "cpe", (0.7, -0.65), "windward | side / leeward"),
           ("79", "Cp,i (+ = internal pressure)", "cpi", (-0.3, 0.2), "internal suction | pressure"),
           ("80", "Ka (area reduction)", "ka", (1.0, 1.0), ""),
           ("81", "Kl (local pressure)", "kl", (1.0, 1.0), "panels near corners"),
           ("82", "Kc,e (combination, external)", "kce", (1.0, 1.0), ""),
           ("83", "Kc,i (combination, internal)", "kci", (1.0, 1.0), ""),
           ("84", "Parapet Cp,n (net, + inward)", "cpn", (1.5, -1.5), "")]
    for rw, t, key, dflt, hint in WIN:
        lab(f"B{rw}", t); v = d.get(key, dflt)
        IN(f"E{rw}", v[0], "0.00", name=key + "i"); IN(f"F{rw}", v[1], "0.00", name=key + "o")
        if hint: small(f"H{rw}", hint)
    qzp = lambda ls: f'IF(«qzp{ls}»="",«qz{ls}»,«qzp{ls}»)'
    for rw, t, ls in (("85", "p ULS = qz (Ka Kc,e Kl Cp,e − Kc,i Cp,i)", "u"), ("86", "p SLS", "s")):
        lab(f"B{rw}", t)
        for col, io, nm_ in (("E", "i", "pin"), ("F", "o", "pout")):
            CA(f"{col}{rw}", f'=N(«qz{ls}»)*(«ka{io}»*«kce{io}»*«kl{io}»*«cpe{io}»-«kci{io}»*«cpi{io}»)', "0.00", name=nm_ + ls)
        U(f"G{rw}", "kPa")
    for rw, t, ls in (("87", "pp parapet ULS = qz,p Cp,n", "u"), ("88", "pp parapet SLS", "s")):
        lab(f"B{rw}", t)
        for col, io, nm_ in (("E", "i", "ppin"), ("F", "o", "ppout")):
            CA(f"{col}{rw}", f'=N({qzp(ls)})*«cpn{io}»', "0.00", name=nm_ + ls)
        U(f"G{rw}", "kPa")
    note(ws, 89, "Net pressure + inward. The load cases below take these values (formulas in the input cells – overwrite to enter a pressure directly). SLS row: larger of inward / outward.")

    heading(ws, 90, "Load Cases")
    table_header(ws, 91, [("B91:C91", "Case"), ("D91", "Description"), ("E91", "γG"), ("F91", "γQ / ψ"), ("G91", "γ roof Wu"),
                          ("H91", "p panel (kPa)"), ("I91", "p parapet (kPa)")])
    LC0 = 92
    for i in range(NLC + 1):
        r = LC0 + i; lc = d["lcs"][i] if i < len(d["lcs"]) else None
        IN(f"B{r}:C{r}", "SLS" if i == NLC else f"LC{i + 1}")
        IN(f"D{r}", lc[0] if lc else None, al=LEFT)
        for j, col in enumerate("EFGHI"):
            v = lc[1 + j] if lc else None
            IN(f"{col}{r}", None if isinstance(v, str) else v, "0.00")
            if isinstance(v, str): F(f"{col}{r}", v)
    SLSR = LC0 + NLC
    note(ws, 99, "p = net pressure on the panel, + inward (towards the building); parapet pressure likewise. Blank γG = case not used. SLS row: deflection (ψ on Q).")
    lab("B100", "SLS deflection limit Hs /"); IN("E100", d.get("dlim", 150), "0", name="dlim")
    lab("G100", "ε*cs (Mcr, Cl 8.5.3.1)", al=RIGHT); IN("H100", d.get("ecs", 0.0006), "0.00000", name="ecs")
    lab("B101", "βd override (blank = permanent share of N*)"); IN("E101", d.get("bdin"), "0.00", name="bdin"); small("H101", "Cl 10.4.3: G share of N*")

    heading(ws, 102, f"Manual Design Actions per Strip (RFEM) – used when Design actions = {MANUAL}")
    note(ws, 103, "N* compression + at the critical section; M* + = inside face in tension. M* top = end moment at the top support. Δ1 = first-order SLS deflection, mm.")
    table_header(ws, 104, [("B104:C104", "Case"), ("D104", "Description"), ("E104", "N* (kN)"), ("F104", "M* span (kNm)"), ("G104", "M* top (kNm)"),
                          ("H104", "V* (kN)"), ("I104", "Δ1 SLS (mm)")])
    MAN0 = 105
    for i in range(NLC + 1):
        r = MAN0 + i; mv = d.get("manual", {}).get(i)
        ws.merge_cells(f"B{r}:C{r}"); F(f"B{r}", f"=$B${LC0 + i}"); ws[f"B{r}"].font = font(8); ws[f"B{r}"].alignment = CEN
        F(f"D{r}", f'=IF($D${LC0 + i}="","",$D${LC0 + i})'); ws[f"D{r}"].font = font(8, color=GREY)
        for j, col in enumerate("EFGHI"):
            if i < NLC and col == "I": continue
            if i == NLC and col in "GH": continue
            IN(f"{col}{r}", mv[j] if mv and j < len(mv) else None, "0.0")
    note(ws, 112, "Manual mode: the magnifier is applied to M* span unless RFEM already includes 2nd-order effects; top band and reactions come from RFEM.")

    # ================================================================== CALCULATIONS (page 2 on)
    r = 114; heading(ws, r, "Strip Properties")
    def prop(rr, s_, ds, f, nf, u, name=None, col="E"):
        sym(rr, s_); desc(rr, ds); CA(f"{col}{rr}", f, nf, name=name); U(f"G{rr}", u)
    prop(115, "b × tw", "strip section (add-in dims)", "=«b»", "0", "mm", name="dimb"); CA("F115", "=«tw»", "0", name="dimt")
    prop(116, "Ag", "b tw", "=«b»*«tw»", "0", "mm²", name="Ag")
    prop(117, "Ig", "b tw³ / 12", "=«b»*«tw»^3/12/1E6", "0.0", "×10⁶ mm⁴", name="Ig")
    prop(118, "Z", "b tw² / 6", "=«b»*«tw»^2/6/1E6", "0.000", "×10⁶ mm³", name="Zs")
    prop(119, "r", "0.3 tw (Cl 10.5.2)", "=0.3*«tw»", "0.0", "mm", name="r")
    prop(120, "Le", "k Hs", "=«k»*«Hs»", "0", "mm", name="Le")
    prop(121, "Le / r", "slenderness ratio", "=«Le»/«r»", "0.0", "", name="sl")
    I_ = "MIN(MATCH(«fc»,EcFc,1),7)"
    prop(122, "Ec", '=IF(N(«ecin»)>0,"entered","Table 3.1.2, interpolated")',
         f"=IF(N(«ecin»)>0,«ecin»,INDEX(EcVal,{I_})+(«fc»-INDEX(EcFc,{I_}))*(INDEX(EcVal,{I_}+1)-INDEX(EcVal,{I_}))/(INDEX(EcFc,{I_}+1)-INDEX(EcFc,{I_})))",
         "0", "MPa", name="Ec")
    prop(123, "f'ct.f", "0.6 √f'c (Cl 3.1.1.3)", "=0.6*SQRT(«fc»)", "0.00", "MPa", name="fctf")
    prop(124, "gs", "self-weight of strip, ρc tw b", "=«rhoc»*«tw»/1000*«b»/1000", "0.00", "kN/m", name="gs")
    prop(125, "L", "Hs", "=«Hs»/1000", "0.000", "m", name="Lm")
    prop(126, "d", '=IF(«TWO»,"tw − cover − db/2 (inside layer)","tw / 2 (central layer)")',
         f'=IF(«TWO»,«tw»-«cover»-IFERROR({barv("«vbar»", 2)},0)/2,«tw»/2)', "0.0", "mm", name="dd")
    prop(127, "dv", "max(0.72 tw, 0.9 d)", "=MAX(0.72*«tw»,0.9*«dd»)", "0.0", "mm", name="dv")
    prop(128, "ρv", "vertical, main bars all layers / (1000 tw)",
         f'=IF(«TWO»,2,1)*IFERROR({barv("«vbar»", 3)},0)/(«vs»*«tw»)', "0.000%", "", name="rhov")
    prop(129, "ρh", "horizontal, main bars all layers / (1000 tw)",
         f'=IF(«TWO»,2,1)*IFERROR({barv("«hbar»", 3)},0)/(«hs»*«tw»)', "0.000%", "", name="rhoh")

    # ---- bar tables for the add-in (not printed, below the print area) - filled in further down
    SB0 = 278; BB0 = 284
    BARS_S = f"$L${SB0}:$P${SB0 + 3}"; CLS_S = f"$Q${SB0}:$Q${SB0 + 3}"
    BARS_B = f"$L${BB0}:$P${BB0 + 3}"; CLS_B = f"$Q${BB0}:$Q${BB0 + 3}"
    SEC_S = '"Rect",«dimb»:«dimt»'
    SEC_B = '"Rect",«dimbb»:«dimbt»'
    MOD = '«fc»,«model»'
    def idx(expr, k): return f'IFERROR(INDEX({expr},{k},1),{expr})'

    heading(ws, 131, "Axis Properties – strip (OAR Concrete add-in)")
    note(ws, 132, "OAR.CONCRETE.COLUMNAXIS on the strip b × tw: strain compatibility, εcu 0.003; balanced point with the outermost tension bar at yield. Mc = Mub (Cl 10.4.4).")
    table_header(ws, 133, [("E133", "+M inside face in tension"), ("F133", "−M outside face in tension")])
    ws.row_dimensions[133].height = 24
    AX = [("Nuo", "squash load", "0", "kN"), ("Nuot", "pure tension", "0", "kN"), ("Nub", "balanced point", "0", "kN"),
          ("Mub", "moment at Nub (Mc)", "0.00", "kNm"), ("Muo", "pure bending", "0.00", "kNm"), ("kuo", "at pure bending", "0.000", ""),
          ("φb", "Tbl 2.2.2, pure bending", "0.000", ""), ("do", "to outermost tension bar", "0.0", "mm"),
          ("Nud", "decompression (dn = do)", "0", "kN"), ("Mud", "moment at Nud", "0.00", "kNm"), ("φt", "Tbl 2.2.2, pure tension", "0.000", "")]
    AX0 = 134
    for i, (s_, ds, nf, u) in enumerate(AX):
        rr = AX0 + i; sym(rr, s_); desc(rr, ds); calc(ws, f"E{rr}", None, nf); calc(ws, f"F{rr}", None, nf); U(f"G{rr}", u)
    AX1 = AX0 + len(AX) - 1
    for col, ax, sfx in (("E", "90", "p"), ("F", "-90", "n")):
        F(f"{col}{AX0}", f'=OAR.CONCRETE.COLUMNAXIS({SEC_S},{BARS_S},{MOD},{ax},{CLS_S},«Ec»)'); arrays.append(f"{col}{AX0}:{col}{AX1}")
        for kk, n_ in enumerate(["Nuo", "Nuot", "Nub", "Mub", "Muo", "kuo", "phib", "do", "Nud", "Mud", "phit"]): nm(f"{n_}{sfx}", f"{col}{AX0 + kk}")

    # ---- first-order actions
    heading(ws, 146, "First-order Actions per Strip (pinned base, pinned top support)")
    note(ws, 147, "x up from the base, L = Hs. w = p bt; top end moment M*T = −pp bt hp²/2 − P roof e (inside face compression −). M(x) = w x (L − x)/2 + M*T x/L.")
    note(ws, 148, "P top = P roof + γG (G head + parapet self-weight). N*(x) = P top + γG gs (L − x). Critical x: max |M| (x = L if the end moment governs).")
    table_header(ws, 149, [("B149:C149", "Case"), ("D149", "w (kN/m)"), ("E149", "P roof (kN)"), ("F149", "P top (kN)"), ("G149", "M*T (kNm)"),
                           ("H149", "x crit (m)"), ("I149", "M2* (kNm)"), ("J149", "N* at x (kN)"), ("K149", "V* (kN)")])
    grey_hdr("L149", "x0 (m)"); grey_hdr("M149", "M(x0)"); grey_hdr("N149", "R base"); grey_hdr("O149", "R top")
    ACT0 = 150
    for i in range(NLC + 1):
        ar = ACT0 + i; lr = LC0 + i; blank = f'$E${lr}=""'
        ws.merge_cells(f"B{ar}:C{ar}"); F(f"B{ar}", f"=$B${lr}"); ws[f"B{ar}"].font = font(8); ws[f"B{ar}"].alignment = CEN
        wpp = f"N($I${lr})*«bt»/1000"
        CA(f"D{ar}", f'=IF({blank},"",N($H${lr})*«bt»/1000)', "0.000")
        CA(f"E{ar}", f'=IF({blank},"",(N($E${lr})*«Gr»+N($F${lr})*«Qr»+N($G${lr})*«Wr»)*«br»/1000)', "0.00")
        CA(f"F{ar}", f'=IF({blank},"",E{ar}+N($E${lr})*(N(«gadd»)+«gs»*N(«hp»)/1000))', "0.00")
        CA(f"G{ar}", f'=IF({blank},"",-{wpp}*(N(«hp»)/1000)^2/2-E{ar}*N(«ecc»)/1000)', "0.00")
        for col, f in (("L", f'=IF({blank},"",IF(D{ar}=0,«Lm»,MIN(«Lm»,MAX(0,«Lm»/2+G{ar}/(D{ar}*«Lm»)))))'),
                       ("M", f'=IF({blank},"",D{ar}*L{ar}*(«Lm»-L{ar})/2+G{ar}*L{ar}/«Lm»)'),
                       ("N", f'=IF({blank},"",D{ar}*«Lm»/2+G{ar}/«Lm»)'),
                       ("O", f'=IF({blank},"",D{ar}*«Lm»/2-G{ar}/«Lm»+{wpp}*N(«hp»)/1000)')):
            F(f"{col}{ar}", f); ws[f"{col}{ar}"].font = font(7, color=GREY)
        CA(f"H{ar}", f'=IF({blank},"",IF(ABS(G{ar})>ABS(M{ar}),«Lm»,L{ar}))', "0.00")
        CA(f"I{ar}", f'=IF({blank},"",IF(ABS(G{ar})>ABS(M{ar}),G{ar},M{ar}))', "0.00")
        CA(f"J{ar}", f'=IF({blank},"",F{ar}+N($E${lr})*«gs»*(«Lm»-H{ar}))', "0.0")
        CA(f"K{ar}", f'=IF({blank},"",MAX(ABS(N{ar}),ABS(N{ar}-D{ar}*«Lm»),ABS({wpp}*N(«hp»)/1000)))', "0.00")
    SLSA = ACT0 + NLC

    # ---- magnifier
    heading(ws, 158, "Design Moments – braced moment magnifier (Cl 10.4)")
    note(ws, 159, "βd = permanent share of N* (Cl 10.4.3), or the override (0 if Le/r ≤ 40 and N* ≤ M2*/2D). Nc = π²/Le² × 182 do φ0 Mc / (1 + βd), Mc = Mub. Short only if Le/r ≤ 25 (transverse load, M1*/M2* = −1).")
    note(ws, 160, "δb = km / (1 − N*/Nc) ≥ 1, km = 1.0 with lateral load (0.6 for the end moment alone). M*d = δb × max(|M2*|, 0.05 tw N*), sign of M2*.")
    table_header(ws, 161, [("B161:C161", "Case"), ("D161", "N* (kN)"), ("E161", "M2* (kNm)"), ("F161", "βd"), ("G161", "Nc (kN)"),
                           ("H161", "δb"), ("I161", "0.05 tw N*"), ("J161", "M*d (kNm)"), ("K161", "source")])
    MAG0 = 162
    for i in range(NLC):
        gr = MAG0 + i; lr = LC0 + i; ar = ACT0 + i; mr = MAN0 + i; blank = f'$E${lr}=""'
        ws.merge_cells(f"B{gr}:C{gr}"); F(f"B{gr}", f"=$B${lr}"); ws[f"B{gr}"].font = font(8); ws[f"B{gr}"].alignment = CEN
        CA(f"D{gr}", f'=IF({blank},"",IF(«MAN»,N($E${mr}),J{ar}))', "0.0")
        CA(f"E{gr}", f'=IF({blank},"",IF(«MAN»,N($F${mr}),I{ar}))', "0.00")
        xg = f"IF(«MAN»,«Lm»/2,H{ar})"
        perm = f"N($E${lr})*(«Gr»*«br»/1000+N(«gadd»)+«gs»*(N(«hp»)/1000+«Lm»-{xg}))"
        CA(f"F{gr}", f'=IF({blank},"",IF(OR(D{gr}<=0,AND(«sl»<=40,D{gr}<=ABS(E{gr})/(2*«tw»/1000))),0,IF(«bdin»<>"",«bdin»,MIN(1,MAX(0,{perm}/D{gr})))))', "0.00")
        CA(f"G{gr}", f'=IF({blank},"",PI()^2/«Le»^2*182*IF(E{gr}>=0,«dop»,«don»)*«phi0»*IF(E{gr}>=0,«Mubp»,«Mubn»)*1E6/(1+F{gr})/1000)', "0")
        km = f"IF(AND(NOT(«MAN»),N($H${lr})=0),0.6,1)"
        CA(f"H{gr}", f'=IF({blank},"",IF(OR(D{gr}<=0,AND(«MAN»,«rfem2»="Yes"),«sl»<=25),1,IF(D{gr}>=G{gr},99,MAX(1,{km}/(1-D{gr}/G{gr})))))', "0.00")
        CA(f"I{gr}", f'=IF({blank},"",MAX(0,0.05*«tw»/1000*D{gr}))', "0.00")
        CA(f"J{gr}", f'=IF({blank},"",IF(E{gr}<0,-1,1)*H{gr}*MAX(ABS(E{gr}),I{gr}))', "0.00")
        F(f"K{gr}", f'=IF({blank},"",IF(«MAN»,"RFEM","calc"))'); ws[f"K{gr}"].font = font(8, color=GREY); ws[f"K{gr}"].alignment = CEN

    # ---- capacity at the critical section
    COLS = "EFGHIJ"
    CP = 169
    heading(ws, CP, "Strength at the Critical Section – strain compatibility (OAR Concrete add-in)")
    note(ws, CP + 1, "OAR.CONCRETE.COLUMN on the strip at N*, M*d: φ from Tbl 2.2.2 (φ0 … φb by N*/Nub; any Class L bar → φb 0.65). Capacity table for each load case.")
    h = CP + 2
    table_header(ws, h, [(f"{c}{h}", "") for c in COLS])
    for i, c in enumerate(COLS): F(f"{c}{h}", f"=$B${LC0 + i}")
    CPR = [("φ", "Tbl 2.2.2 at N*", "0.000", ""), ("Mu", "capacity", "0.00", "kNm"), ("φMu", "", "0.00", "kNm"),
           ("Mux", "", "0.00", "kNm"), ("Muy", "", "0.00", "kNm"), ("θ", "compression direction", "0", "deg"),
           ("dn", "neutral axis depth", "0.0", "mm"), ("Nub", "balanced", "0", "kN"), ("Nuo", "squash", "0", "kN"),
           ("εs", "outermost tension bar strain", "0.00%", "")]
    C0 = h + 1
    for kk, (s_, ds, nf, u) in enumerate(CPR):
        rr = C0 + kk; sym(rr, s_); desc(rr, ds)
        for c in COLS: calc(ws, f"{c}{rr}", None, nf)
        U(f"K{rr}", u)
    C1 = C0 + len(CPR) - 1
    for i, c in enumerate(COLS):
        lr = LC0 + i; gr = MAG0 + i
        F(f"{c}{C0}", f'=IF($E${lr}="","",OAR.CONCRETE.COLUMN({SEC_S},{BARS_S},{MOD},$D${gr},$J${gr},0,«phi0»,{CLS_S},«Ec»))')
        arrays.append(f"{c}{C0}:{c}{C1}")
    RM = C1 + 1; sym(RM, "|M*d|"); desc(RM, "design moment, magnified")
    for i, c in enumerate(COLS): CA(f"{c}{RM}", f'=IF($E${LC0 + i}="","",ABS($J${MAG0 + i}))', "0.00")
    U(f"K{RM}", "kNm")
    RSC = RM + 1; sym(RSC, "ratio"); desc(RSC, "|M*d| / φMu")
    for i, c in enumerate(COLS): CA(f"{c}{RSC}", f'=IF($E${LC0 + i}="","",{c}{RM}/{c}{C0 + 2})', "0.00")

    # ---- top support (end moment)
    TS = RSC + 2   # 169
    heading(ws, TS, "Top Support – end moment from parapet and roof eccentricity (not magnified)")
    h = TS + 1
    table_header(ws, h, [(f"{c}{h}", "") for c in COLS])
    for i, c in enumerate(COLS): F(f"{c}{h}", f"=$B${LC0 + i}")
    T0 = h + 1
    for kk, (s_, ds, nf, u) in enumerate([("N*T", "P top", "0.0", "kN"), ("M*T", "end moment", "0.00", "kNm"),
                                          ("φMu", "capacity at N*T", "0.00", "kNm"), ("ratio", "|M*T| / φMu", "0.00", "")]):
        rr = T0 + kk; sym(rr, s_); desc(rr, ds); U(f"K{rr}", u)
        for i, c in enumerate(COLS):
            lr = LC0 + i; ar = ACT0 + i; mr = MAN0 + i; blank = f'$E${lr}=""'
            f = {0: f'=IF({blank},"",$F${ar})',
                 1: f'=IF({blank},"",IF(«MAN»,N($G${mr}),$G${ar}))',
                 2: f'=IF(OR({blank},N({c}{T0 + 1})=0),"",{idx(f'OAR.CONCRETE.COLUMN({SEC_S},{BARS_S},{MOD},{c}{T0},{c}{T0 + 1},0,«phi0»,{CLS_S},«Ec»)', 3)})',
                 3: f'=IF({c}{T0 + 2}="","",ABS({c}{T0 + 1})/{c}{T0 + 2})'}[kk]
            CA(f"{c}{rr}", f, nf)
    TR = T0 + 3

    # ---- shear and reactions
    SH = TR + 2   # 176
    heading(ws, SH, "Shear and Support Reactions per Strip")
    note(ws, SH + 1, "φVuc = 0.75 kv b dv √f'c, kv = 200 / (1000 + 1.3 dv) ≤ 0.10 (Cl 8.2.4.3, no shear reinforcement). Reactions + = load inward; per metre ÷ bt.")
    h = SH + 2
    table_header(ws, h, [(f"{c}{h}", "") for c in COLS])
    for i, c in enumerate(COLS): F(f"{c}{h}", f"=$B${LC0 + i}")
    R0 = h + 1
    for kk, (s_, ds, nf, u) in enumerate([("V*", "max shear in the strip", "0.00", "kN"), ("R base", "per strip", "0.00", "kN"),
                                          ("R top", "per strip (incl. parapet)", "0.00", "kN"), ("R base", "per metre of wall", "0.00", "kN/m"),
                                          ("R top", "per metre of wall", "0.00", "kN/m")]):
        rr = R0 + kk; sym(rr, s_); desc(rr, ds); U(f"K{rr}", u)
        for i, c in enumerate(COLS):
            lr = LC0 + i; ar = ACT0 + i; mr = MAN0 + i; blank = f'$E${lr}=""'
            f = {0: f'=IF({blank},"",IF(«MAN»,ABS(N($H${mr})),$K${ar}))',
                 1: f'=IF(OR({blank},«MAN»),"",$N${ar})', 2: f'=IF(OR({blank},«MAN»),"",$O${ar})',
                 3: f'=IF({c}{R0 + 1}="","",{c}{R0 + 1}/(«bt»/1000))', 4: f'=IF({c}{R0 + 2}="","",{c}{R0 + 2}/(«bt»/1000))'}[kk]
            CA(f"{c}{rr}", f, nf)
    r = R0 + 5
    sym(r, "kv"); desc(r, "200 / (1000 + 1.3 dv) ≤ 0.10"); CA(f"E{r}", "=MIN(0.1,200/(1000+1.3*«dv»))", "0.000", name="kv")
    r += 1; sym(r, "φVuc"); desc(r, "0.75 kv b dv √f'c (√f'c ≤ 8)"); CA(f"E{r}", "=0.75*«kv»*«b»*«dv»*MIN(SQRT(«fc»),8)/1000", "0.0", name="phiVuc"); U(f"G{r}", "kN")
    VROW = R0

    # ---- top band (joint support)
    TB = r + 2   # 189
    heading(ws, TB, "Top Band – panel supported at the joints only (approximate; RFEM for openings in the band)")
    note(ws, TB + 1, "Band hb × tw spans W between the joint supports as a simply supported horizontal beam, loaded by R top per metre: M* = q W² / 8. Horizontal bars in hb.")
    r = TB + 2
    sym(r, "hb × tw"); desc(r, "band section (add-in dims)"); CA(f"E{r}", "=«hb»", "0", name="dimbb"); CA(f"F{r}", "=«tw»", "0", name="dimbt"); U(f"G{r}", "mm")
    r += 1; sym(r, "As"); desc(r, "horizontal bars in the band, all layers"); CA(f"E{r}", f"=SUM($L${BB0}:$L${BB0 + 3})", "0", name="Asb"); U(f"G{r}", "mm²")
    r += 1; table_header(ws, r, [(f"E{r}", "+ (inward)"), (f"F{r}", "− (outward)")])
    r += 1; sym(r, "Muo"); desc(r, "OAR.CONCRETE.ULT, N = 0")
    r_muo = r
    for col, sense in (("E", "Sagging"), ("F", "Hogging")):
        CA(f"{col}{r}", f'=IF(NOT(«BAND»),"",{idx(f'OAR.CONCRETE.ULT({SEC_B},{BARS_B},{MOD},"{sense}",0,«Ec»,{CLS_B},"Whole section 0.65")', 13)})', "0.00")
    U(f"G{r}", "kNm")
    r += 1; sym(r, "φ"); desc(r, "Tbl 2.2.2 (Class L in tension → 0.65)")
    for col, sense in (("E", "Sagging"), ("F", "Hogging")):
        CA(f"{col}{r}", f'=IF(NOT(«BAND»),"",{idx(f'OAR.CONCRETE.ULT({SEC_B},{BARS_B},{MOD},"{sense}",0,«Ec»,{CLS_B},"Whole section 0.65")', 14)})', "0.000")
    r += 1; sym(r, "φMuo"); desc(r, "")
    for col in "EF": CA(f"{col}{r}", f'=IF({col}{r_muo}="","",{col}{r_muo}*{col}{r - 1})', "0.00")
    U(f"G{r}", "kNm"); nm("phiMbp", f"E{r}"); nm("phiMbn", f"F{r}")
    r += 1
    table_header(ws, r, [(f"{c}{r}", "") for c in COLS])
    for i, c in enumerate(COLS): F(f"{c}{r}", f"=$B${LC0 + i}")
    B0 = r + 1
    for kk, (s_, ds, nf, u) in enumerate([("q*", "R top per metre", "0.00", "kN/m"), ("M*b", "q* W² / 8", "0.00", "kNm"),
                                          ("ratio", "|M*b| / φMuo", "0.00", "")]):
        rr = B0 + kk; sym(rr, s_); desc(rr, ds); U(f"K{rr}", u)
        for i, c in enumerate(COLS):
            blank = f'$E${LC0 + i}=""'
            f = {0: f'=IF(OR({blank},NOT(«BAND»)),"",{c}{R0 + 4})',
                 1: f'=IF({c}{B0}="","",{c}{B0}*(«W»/1000)^2/8)',
                 2: f'=IF({c}{B0 + 1}="","",ABS({c}{B0 + 1})/IF({c}{B0 + 1}>=0,«phiMbp»,«phiMbn»))'}[kk]
            CA(f"{c}{rr}", f, nf)
    BR = B0 + 2
    # band SLS deflection
    r = BR + 1
    sym(r, "q s"); desc(r, "SLS R top per metre"); CA(f"E{r}", f'=IF(OR(NOT(«BAND»),$E${SLSR}=""),"",$O${SLSA}/(«bt»/1000))', "0.00", name="qsb"); U(f"G{r}", "kN/m")
    r += 1; sym(r, "Ms,b"); desc(r, "q s W² / 8"); CA(f"E{r}", '=IF(«qsb»="","",«qsb»*(«W»/1000)^2/8)', "0.00", name="Msb"); U(f"G{r}", "kNm")
    r += 1; sym(r, "Ig,b | Icr,b"); desc(r, "hb tw³/12 | OAR.CONCRETE.CRACKED")
    CA(f"E{r}", "=«hb»*«tw»^3/12/1E6", "0.0", name="Igb")
    CA(f"F{r}", f'=IF(OR(«Msb»="",N(«Msb»)=0),"",{idx('OAR.CONCRETE.CRACKED(' + SEC_B + ',' + BARS_B + ',«Ec»,ABS(«Msb»),IF(«Msb»<0,"Hogging","Sagging"))', 2)})', "0.0", name="Icrb")
    U(f"G{r}", "×10⁶ mm⁴")
    r += 1; sym(r, "Mcr,b"); desc(r, "Zb (f'ct.f − σcs)"); CA(f"E{r}", "=MAX(0,«hb»*«tw»^2/6*(«fctf»-«scs»)/1E6)", "0.00", name="Mcrb"); U(f"G{r}", "kNm")
    r += 1; sym(r, "Ief,b"); desc(r, "Cl 8.5.3.1, Ief.max 0.6 Ig if ρ < 0.005")
    iefmaxb = "IF(«Asb»/(«hb»*«dd»)>=0.005,1,0.6)*«Igb»"
    CA(f"E{r}", f'=IF(«Msb»="","",IF(OR(«Icrb»="",ABS(«Msb»)<=«Mcrb»),{iefmaxb},MIN({iefmaxb},«Icrb»/(1-(1-«Icrb»/«Igb»)*(«Mcrb»/«Msb»)^2))))', "0.0", name="Iefb")
    U(f"G{r}", "×10⁶ mm⁴")
    r += 1; sym(r, "Δb"); desc(r, "5 q s W⁴ / (384 Ec Ief,b), band midspan"); CA(f"E{r}", '=IF(«Msb»="","",5*«qsb»*«W»^4/(384*«Ec»*«Iefb»*1E6))', "0.0", name="Db"); U(f"G{r}", "mm")
    BAND_END = r

    # ---- SLS deflection of the strip
    SD = BAND_END + 2
    heading(ws, SD, "Serviceability – deflection of the strip under SLS (short term)")
    note(ws, SD + 1, "Δ1 = (5 w L⁴/384 + M*T L²/16) / (Ec Ief) at mid-height; Ief Cl 8.5.3.1 (Mcr with σcs and N/Ag); P-Δ: Δ = Δ1 / (1 − Ns/Ncr), Ncr = π² Ec Ief / Le².")
    r = SD + 2
    sl_blank = f'$E${SLSR}=""'
    sym(r, "Ns | Ms"); desc(r, "SLS axial at mid-height | max first-order moment")
    CA(f"E{r}", f'=IF({sl_blank},"",IF(«MAN»,N($E${MAN0 + NLC}),$F${SLSA}+N($E${SLSR})*«gs»*«Lm»/2))', "0.0", name="Ns")
    CA(f"F{r}", f'=IF({sl_blank},"",IF(«MAN»,N($F${MAN0 + NLC}),$I${SLSA}))', "0.00", name="Ms"); U(f"G{r}", "kN | kNm")
    r += 1; sym(r, "ρw | ρcw"); desc(r, "tension | compression steel / (b d), +M")
    CA(f"E{r}", f'=IF(«TWO»,($L${SB0 + 1}+$L${SB0 + 3}),SUM($L${SB0}:$L${SB0 + 3}))/(«b»*«dd»)', "0.0000", name="rhow")
    CA(f"F{r}", f'=IF(«TWO»,($L${SB0}+$L${SB0 + 2})/(«b»*«dd»),0)', "0.0000", name="rhocw")
    r += 1; sym(r, "σcs"); desc(r, "(2.5ρw − 0.8ρcw)/(1 + 50ρw) Es ε*cs ≥ 0"); CA(f"E{r}", "=MAX(0,(2.5*«rhow»-0.8*«rhocw»)/(1+50*«rhow»)*200000*«ecs»)", "0.00", name="scs"); U(f"G{r}", "MPa")
    r += 1; sym(r, "Mcr"); desc(r, "Z (f'ct.f − σcs + Ns/Ag) ≥ 0"); CA(f"E{r}", '=IF(«Ns»="","",MAX(0,«Zs»*(«fctf»-«scs»+«Ns»*1000/«Ag»)))', "0.00", name="Mcr"); U(f"G{r}", "kNm")
    r += 1; sym(r, "Icr"); desc(r, "OAR.CONCRETE.CRACKED at Ms")
    CA(f"E{r}", f'=IF(OR(«Ms»="",N(«Ms»)=0),"",{idx('OAR.CONCRETE.CRACKED(' + SEC_S + ',' + BARS_S + ',«Ec»,ABS(«Ms»),IF(«Ms»<0,"Hogging","Sagging"))', 2)})', "0.00", name="Icr")
    U(f"G{r}", "×10⁶ mm⁴")
    r += 1; sym(r, "Ief"); desc(r, "Icr / (1 − (1 − Icr/Ig)(Mcr/Ms)²) ≤ Ief.max (0.6 Ig if ρw < 0.005)")
    iefmax = "IF(«rhow»>=0.005,1,0.6)*«Ig»"
    CA(f"E{r}", f'=IF(«Ns»="","",IF(OR(«Icr»="",ABS(«Ms»)<=«Mcr»),{iefmax},MIN({iefmax},«Icr»/(1-(1-«Icr»/«Ig»)*(«Mcr»/«Ms»)^2))))', "0.00", name="Ief")
    U(f"G{r}", "×10⁶ mm⁴")
    r += 1; sym(r, "Δ1"); desc(r, '=IF(«MAN»,"entered (RFEM, first order)","first order, mid-height")')
    CA(f"E{r}", f'=IF(«Ns»="","",IF(«MAN»,N($I${MAN0 + NLC}),(5*$D${SLSA}*«Hs»^4/384+$G${SLSA}*1E6*«Hs»^2/16)/(«Ec»*«Ief»*1E6)))', "0.0", name="D1"); U(f"G{r}", "mm")
    r += 1; sym(r, "Ncr"); desc(r, "π² Ec Ief / Le²"); CA(f"E{r}", '=IF(«Ns»="","",PI()^2*«Ec»*«Ief»*1E6/«Le»^2/1000)', "0", name="Ncr"); U(f"G{r}", "kN")
    r += 1; sym(r, "Δ2"); desc(r, "Δ1 / (1 − Ns / Ncr)"); CA(f"E{r}", '=IF(«Ns»="","",IF(«Ns»>=«Ncr»,999,«D1»/(1-MAX(0,«Ns»)/«Ncr»)))', "0.0", name="D2"); U(f"G{r}", "mm")
    r += 1; sym(r, "Δ"); desc(r, "|Δ2| + Δb / 2 (joint support: band deflection at mid-height)")
    CA(f"E{r}", '=IF(«Ns»="","",ABS(«D2»)+IF(AND(«BAND»,«Db»<>""),ABS(N(«Db»))/2,0))', "0.0", name="Dtot"); U(f"G{r}", "mm")
    r += 1; sym(r, "Δ limit"); desc(r, "Hs / limit"); CA(f"E{r}", "=«Hs»/«dlim»", "0.0", name="Dlim"); U(f"G{r}", "mm")
    CALC_END = r
    LAST = CALC_END + 2
    note(ws, LAST, "Not covered: lifting and handling, temporary bracing, lintels over openings (in-plane), dowel / cleat capacity, FRL (AS 3600 Sec 5).")

    # ---- engine bar tables (not printed)
    lab(f"L{SB0 - 2}", "Add-in bar tables (not printed): As | depth from outside face | fsy | x | Es | class", sz=7, color=GREY, i=True)
    lab(f"S{SB0}", "strip: main outside/central | main inside | extra outside/central | extra inside", sz=7, color=GREY)
    lab(f"S{BB0}", "top band: horizontal bars, same order", sz=7, color=GREY)
    def bar_rows(r0, bar, spacing, width, xspacing, xbarref, depth_out, depth_in, xdepth_out, xdepth_in):
        ab = f'IFERROR({barv(bar, 3)},0)'; ax_ = f'IFERROR({barv(xbarref, 3)},0)'
        rows = [(f'=IF(N({spacing})>0,{ab}*{width}/{spacing},0)', depth_out, bar),
                (f'=IF(AND(«TWO»,N({spacing})>0),{ab}*{width}/{spacing},0)', depth_in, bar),
                (f'=IF(N({xspacing})>0,{ax_}*{width}/{xspacing},0)', xdepth_out, xbarref),
                (f'=IF(AND(«TWO»,N({xspacing})>0),{ax_}*{width}/{xspacing},0)', xdepth_in, xbarref)]
        for k_, (fa, fd, bref) in enumerate(rows):
            rr = r0 + k_
            for col, f in (("L", fa), ("M", fd), ("N", f'=IFERROR({barv(bref, 4)},500)'), ("P", f'=IFERROR({barv(bref, 6)},200000)'),
                           ("Q", f'=IFERROR({barv(bref, 5)},"N")')):
                F(f"{col}{rr}", f); ws[f"{col}{rr}"].font = font(7, color=GREY)
    dbv = f'IFERROR({barv("«vbar»", 2)},0)'; dbh = f'IFERROR({barv("«hbar»", 2)},0)'
    dbx = f'IFERROR({barv("«xbar»", 2)},0)'; dbhb = f'IFERROR({barv("«hbbar»", 2)},0)'
    bar_rows(SB0, "«vbar»", "«vs»", "«b»", "«xs»", "«xbar»",
             f"=IF(«TWO»,«cover»+{dbv}/2,«tw»/2)", f"=«tw»-«cover»-{dbv}/2",
             f"=IF(«TWO»,«cover»+{dbx}/2,«tw»/2)", f"=«tw»-«cover»-{dbx}/2")
    bar_rows(BB0, "«hbar»", "«hs»", "«hb»", "«hbs»", "«hbbar»",
             f"=IF(«TWO»,«cover»+{dbv}+{dbh}/2,«tw»/2)", f"=«tw»-«cover»-{dbv}-{dbh}/2",
             f"=IF(«TWO»,«cover»+{dbv}+{dbhb}/2,«tw»/2)", f"=«tw»-«cover»-{dbv}-{dbhb}/2")

    # ================================================================== SUMMARY + CHECKS (page 1)
    heading(ws, 8, "Design Summary")
    lab("B9", "Adopted panel / strip", sz=7, color=GREY); lab("F9", "Max utilisation", sz=7, color=GREY); lab("I9", "Status", sz=7, color=GREY)
    ws.merge_cells("B10:E10"); ws.merge_cells("F10:G10"); ws.merge_cells("I10:K10")
    reo = ('IF(«TWO»,"2 layers ","central ")&«vbar»&"-"&«vs»&IF(AND(N(«xs»)>0,«xbar»<>"None")," + "&«xbar»&"-"&«xs»&" extra","")&IF(«TWO»," EF","")')
    F("B10", f'=«tw»&" thk  ·  "&{reo}&"  ·  f\'c "&«fc»')
    ws["B10"].font = font(12, True, color=GREEN); ws["B10"].alignment = LEFT
    CHK0, CHK1 = 15, 29
    c = ws["F10"]; c.value = f'=IF(COUNT(I{CHK0}:I{CHK1})=0,"–",MAX(I{CHK0}:I{CHK1}))'; c.font = font(12, True, color=GREEN); c.number_format = "0%"; c.alignment = LEFT
    ADD = [f"E{AX0}:F{AX1}", f"E{C0}:J{C0}", f"E{T0 + 2}:J{T0 + 2}", f"E{r_muo}:F{r_muo + 1}", "«Icr»", "«Icrb»"]
    errs = "+".join(f'COUNTIF({a},"#OAR*")' for a in ADD)
    c = ws["I10"]; c.value = (f'=IF($V$7,"UPDATE ADD-IN",IF({errs}>0,"INPUT ERROR",IF(COUNT(I{CHK0}:I{CHK1})=0,"PENDING",'
                              f'IF(OR(MAX(I{CHK0}:I{CHK1})>1,COUNTIF(J{CHK0}:J{CHK1},"NG")>0),"FAIL","PASS"))))')
    pend.append(("I10", c.value))
    c.font = font(10, True, color=GREY); c.alignment = CEN
    for x in ws["I10:K10"][0]: x.border = BOX
    passfail_cf(ws, "I10:K10", "$I$10")
    ws.conditional_formatting.add("I10:K10", FormulaRule(formula=['$I$10="UPDATE ADD-IN"'], fill=PatternFill(bgColor="C55A11"), font=Font(color="FFFFFF", bold=True)))
    F("B11", '=«stripd»&" (b "&«b»&")  ·  span "&TEXT(«Hs»/1000,"0.0")&" m"&IF(N(«hp»)>0," + parapet "&TEXT(«hp»/1000,"0.0")&" m","")'
             '&"  ·  top: "&IF(«top»="' + TOP_OPTS[1] + '","joints only","continuous")&"  ·  "&IF(«MAN»,"actions from RFEM","actions calculated")&"  ·  "&«model»&" block"')
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
        if unit_.startswith("="): F(f"G{rr}", unit_); ws[f"G{rr}"].font = font(9, color=GREY)
        else: U(f"G{rr}", unit_)
        ws[f"G{rr}"].alignment = CEN
        CA(f"H{rr}", other, onf); CA(f"I{rr}", util, "0%")
        ws.merge_cells(f"J{rr}:K{rr}"); F(f"J{rr}", status if status else f'=IF(I{rr}="","–",IF(I{rr}<=1,"OK","NG"))')
        ws[f"J{rr}"].font = font(9, True, color=GREEN); ws[f"J{rr}"].alignment = CEN
    for i in range(NLC):
        rr = 15 + i; c = COLS[i]; lr = LC0 + i; gr = MAG0 + i; ar = ACT0 + i; blank = f'$E${lr}=""'
        basis = (f'=IF({blank},"(not used)",IF(«MAN»,"RFEM","x "&TEXT($H${ar},"0.0")&" m")&"  ·  N* "&TEXT($D${gr},"0")&" kN"'
                 f'&IF($H${gr}>=99,"  ·  N* ≥ Nc: unstable",IF($H${gr}>1,"  ·  δb "&TEXT($H${gr},"0.00"),"")))')
        chk(rr, f'=$B${lr}&" N–M"', basis, f'=IF({blank},"",{c}{RM})', f'=IF({blank},"",{c}{C0 + 2})', "kNm",
            f'=IF({blank},"",{c}{RSC})', other=f'=IF({blank},"",{c}{C0})', onf='"φ "0.00')
    trow = f"E{TR}:J{TR}"
    tcase = f'INDEX($E${T0 - 1}:$J${T0 - 1},MATCH(MAX({trow}),{trow},0))'
    chk(21, "Top support", f'=IF(COUNT({trow})=0,"(no end moment)","end moment  ·  "&{tcase})',
        f'=IF(COUNT({trow})=0,"",ABS(INDEX($E${T0 + 1}:$J${T0 + 1},MATCH(MAX({trow}),{trow},0))))',
        f'=IF(COUNT({trow})=0,"",INDEX($E${T0 + 2}:$J${T0 + 2},MATCH(MAX({trow}),{trow},0)))', "kNm",
        f'=IF(COUNT({trow})=0,"",MAX({trow}))')
    brow = f"E{BR}:J{BR}"
    chk(22, "Top band", f'=IF(NOT(«BAND»),IF(«top»="{TOP_OPTS[1]}","(manual: band from RFEM)","(not used – continuous top support)"),"W "&TEXT(«W»/1000,"0.0")&" m, hb "&«hb»&"  ·  "&IFERROR(INDEX($E${B0 - 1}:$J${B0 - 1},MATCH(MAX({brow}),{brow},0)),""))',
        f'=IF(COUNT({brow})=0,"",ABS(INDEX($E${B0 + 1}:$J${B0 + 1},MATCH(MAX({brow}),{brow},0))))',
        f'=IF(COUNT({brow})=0,"",ABS(INDEX($E${B0 + 1}:$J${B0 + 1},MATCH(MAX({brow}),{brow},0)))/MAX({brow}))', "kNm",
        f'=IF(COUNT({brow})=0,"",MAX({brow}))')
    vrow = f"E{VROW}:J{VROW}"
    chk(23, "V* ≤ φVuc", f'=IF(COUNT({vrow})=0,"","Cl 8.2.4.3, no shear reo  ·  "&INDEX($E${VROW - 1}:$J${VROW - 1},MATCH(MAX({vrow}),{vrow},0)))',
        f'=IF(COUNT({vrow})=0,"",MAX({vrow}))', "=«phiVuc»", "kN", f'=IF(E23="","",E23/F23)')
    heading(ws, 24, "Serviceability", sz=8)
    chk(25, "Δ ≤ Hs / limit", f'=IF(«Ns»="","(no SLS case)","SLS wind, Ief "&TEXT(«Ief»,"0")&"×10⁶, P-Δ "&TEXT(IFERROR(«D2»/«D1»,1),"0.00")&IF(«BAND»," + band",""))',
        '=«Dtot»', '=«Dlim»', "mm", '=IF(OR(«Dtot»="",«Dlim»=0),"",«Dtot»/«Dlim»)', nf="0.0")
    heading(ws, 26, "Detailing", sz=8)
    chk(27, "ρv ≥ 0.0015", "vertical, main bars (Cl 11.7.1)", "=«rhov»", "=0.0015", "", '=IF(«rhov»=0,99,0.0015/«rhov»)', nf="0.000%")
    chk(28, '=IF(«hrest»="Yes","ρh ≥ 0.0025","ρh ≥ 0.0015")',
        '=IF(«hrest»="Yes","horizontal, restrained (Cl 11.7.2)","horizontal, free to move (Cl 11.7.2)")',
        "=«rhoh»", '=IF(«hrest»="Yes",0.0025,0.0015)', "", '=IF(«rhoh»=0,99,F28/«rhoh»)', nf="0.000%")
    chk(29, "spacing", "max(sv, sh) ≤ min(2.5 tw, 350) (Cl 11.7.3)", "=MAX(«vs»,«hs»)", "=MIN(2.5*«tw»,350)", "mm", "=E29/F29", nf="0")
    heading(ws, 30, "Information", sz=8)
    dmax = f"MAX(H{MAG0}:H{MAG0 + NLC - 1})"
    chk(31, "Le / r", f'=IF(«sl»>120,"> 120: confirm magnifier, or RFEM","moment magnifier, δb ≤ "&TEXT({dmax},"0.00"))',
        "=«sl»", "=120", "", None, status='=IF(«sl»>120,"check","info")', nf="0")
    rt = f"$E${R0 + 4}:$J${R0 + 4}"; rb = f"$E${R0 + 3}:$J${R0 + 3}"
    def rinfo(rr, s_, rng, perpanel):
        mx = f"MAX(MAX({rng}),-MIN({rng}))"
        chk(rr, s_, f'=IF(COUNT({rng})=0,IF(«MAN»,"(manual: from RFEM)",""),INDEX($E${R0 - 1}:$J${R0 - 1},IFERROR(MATCH({mx},{rng},0),MATCH(-{mx},{rng},0)))&"  ·  "&TEXT({mx}*«W»/1000,"0.0")&" kN per panel W")',
            f'=IF(COUNT({rng})=0,"",{mx})', None, "kN/m", None, status="info")
    rinfo(32, "R top (ULS)", rt, "")
    rinfo(33, "R base (ULS)", rb, "")
    for rr in (31, 32, 33): ws[f"J{rr}"].font = font(8, i=True, color=GREY)
    ws.conditional_formatting.add("J31:K31", FormulaRule(formula=['$J$31="check"'], font=Font(color=AMBER, bold=True, italic=False)))
    status_cf(ws, f"J{CHK0}:K{CHK1}"); util_cf(ws, f"I{CHK0}:I{CHK1}", f"I{CHK0}")
    # governing line
    F("B12", f'=IFERROR(IF({errs}>0,"Add-in error – check the add-in cells on the calculation pages",IF(COUNT(I{CHK0}:I{CHK1})=0,"","Governs: "&INDEX(B{CHK0}:B{CHK1},MATCH(MAX(I{CHK0}:I{CHK1}),I{CHK0}:I{CHK1},0))'
             f'&" – "&INDEX(D{CHK0}:D{CHK1},MATCH(MAX(I{CHK0}:I{CHK1}),I{CHK0}:I{CHK1},0)))),"")')
    ws["B12"].font = font(7, color=GREY)

    # ================================================================== resolve, arrays, validation, grey-outs, print
    for ref, f in pend: ws[ref].value = Fx(f)
    for ref in arrays:
        top = ref.split(":")[0]; ws[top].value = ArrayFormula(ref, ws[top].value)
    def dvl(formula, ref):
        v = DataValidation(type="list", formula1=formula, allow_blank=True); ws.add_data_validation(v)
        for part in ref.split(): v.add(part)
    dvl('"Rectangular,Curvilinear,Parabolic"', "E38"); dvl(f'"Calculated,{MANUAL}"', "E42"); dvl('"No,Yes"', "E43 E66")
    dvl(f'"{TOP_OPTS[0]},{TOP_OPTS[1]}"', "E48"); dvl('"Central,Two layers"', "E60")
    dvl("=BarList", "E62:E65")
    grey = [("E61", 'NOT(«TWO»)', "B61"), ("E49", f'«top»<>"{TOP_OPTS[1]}"', "B49"), ("E65:F65", f'«top»<>"{TOP_OPTS[1]}"', "B65"),
            ("E43", "NOT(«MAN»)", "B43"), (f"E{MAN0}:I{MAN0 + NLC}", "NOT(«MAN»)", f"B{MAN0}:D{MAN0 + NLC}")]
    for ref, cond, labref in grey:
        f = Fx(cond)
        ws.conditional_formatting.add(ref, FormulaRule(formula=[f], fill=PatternFill(bgColor=NA_FILL), font=Font(color=NA_FONT)))
        ws.conditional_formatting.add(labref, FormulaRule(formula=[f], font=Font(color=NA_FONT)))
    for rr in range(15, LAST + 1):
        if ws.row_dimensions[rr].height is None: ws.row_dimensions[rr].height = 14
    ws.print_area = f"A1:K{LAST}"
    for b_ in (58, 113, CP - 1, SD - 1): ws.row_breaks.append(Break(id=b_))
    return ws


# ---------------- examples ----------------
# Load cases: (description, γG, γQ/ψ, γ roof Wu, p panel kPa, p parapet kPa); last row = SLS.
SLS_P = "=IF(ABS(«pins»)>=ABS(«pouts»),«pins»,«pouts»)"
SLS_PP = "=IF(ABS(«pins»)>=ABS(«pouts»),«ppins»,«ppouts»)"
WIND_LCS = [("1.2G + 1.5Q", 1.2, 1.5, 0, 0, 0),
            ("1.2G + Wu inward", 1.2, 0, 0, "=«pinu»", "=«ppinu»"),
            ("0.9G + Wu inward + roof uplift", 0.9, 0, 1, "=«pinu»", "=«ppinu»"),
            ("0.9G + Wu outward + roof uplift", 0.9, 0, 1, "=«poutu»", "=«ppoutu»"),
            ("1.2G + Wu outward", 1.2, 0, 0, "=«poutu»", "=«ppoutu»"),
            ("G + 0.3Q + Eu (face load)", 1.0, 0.3, 0, "=«pE»", "=«pE»"),
            ("G + ψs Q + Ws", 1.0, 0.7, 0, SLS_P, SLS_PP)]
WIND = dict(qzu=1.2, qzs=0.6)
p1 = dict(**WIND, loc="Typical solid panel – 9.0 m span, 0.9 m parapet", tw=200, fc=32, Hs=9000, hp=900, W=6000, top=TOP_OPTS[0],
          strip="Solid panel", b=1000, bt=1000, br=1000, gadd=0, lay="Central", vbar="N16", vs=200, hbar="N10", hs=200,
          Gr=2.5, Qr=1.2, Wr=-3.0, ecc=150, ceq=0.15, lcs=WIND_LCS)
p2 = dict(**WIND, loc="Jamb strip beside a 4.0 m roller door", tw=200, fc=32, Hs=9000, hp=900, W=6000, top=TOP_OPTS[0],
          strip="Jamb strip – 4 m roller door", b=1200, bt=3200, br=3200, gadd=13.5, lay="Two layers", cover=40,
          vbar="N12", vs=200, hbar="N12", hs=200, xbar="N20", xs=400, Gr=2.5, Qr=1.2, Wr=-3.0, ecc=150, ceq=0.15, lcs=WIND_LCS)
p3 = dict(**WIND, loc="Solid panel – top supported at the panel joints only (rafters)", tw=200, fc=32, Hs=9000, hp=0, W=6000, top=TOP_OPTS[1], hb=1200,
          strip="Solid panel", b=1000, bt=1000, br=1000, gadd=0, lay="Central", vbar="N16", vs=200, hbar="N12", hs=200, hbbar="N16", hbs=400,
          Gr=2.5, Qr=1.2, Wr=-3.0, ecc=150, ceq=0.15, lcs=WIND_LCS)


if __name__ == "__main__":
    for nm_, dd in (("P1", p1), ("P2", p2), ("P3", p3)): build_tab(nm_, dd)
    wb._sheets = [wb[n] for n in ("P1", "P2", "P3", "Bars")]
    wb.active = 0
    from openpyxl.workbook.properties import CalcProperties
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    wb.save(OUT)
    print("saved", OUT)
