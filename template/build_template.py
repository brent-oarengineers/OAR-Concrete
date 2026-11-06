"""Builds template/OAR - RC Flexural Design.xlsx (OAR house calc style).

Run from the repo root:  python template/build_template.py
Needs openpyxl. The section-analysis cells call the OAR Concrete add-in (OAR.CONCRETE.ULT and
OAR.CONCRETE.CRACKED), so the add-in must be loaded in Excel. Every code check (φ, minimum strength,
crack control limits, bar spacing, shrinkage) is a visible worksheet formula.

One calc tab = one flexural element, slab or beam (Member type switch), both bending directions.
Example tabs: S1 slab strip (Brent's 25 Lutana X-strip, previous slab sheet), B1 wide band beam.
"""
import sys, os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule, CellIsRule, DataBarRule
from openpyxl.drawing.image import Image as XLImage
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.formula import ArrayFormula

HERE = os.path.dirname(os.path.abspath(__file__))
LOGO = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "assets", "oar-logo.png")
OUT = sys.argv[2] if len(sys.argv) > 2 else os.path.join(HERE, "OAR - RC Flexural Design.xlsx")

from oar_style import *  # house-style helpers (oar-calc-sheet skill)

wb = Workbook()
wb.remove(wb.active)

# =========================== DATA TAB: Bars ===========================
# AS/NZS 4671 nominal areas. L = mesh wires (Class L, e.g. SL92 = L9 @ 200). R = plain round (R250N).
from oar_style import BARS, BAR_NOTES as notes
D_ = wb.create_sheet("Bars")
setup_data_tab(D_, "Reinforcement – bar and mesh wire properties",
               "AS/NZS 4671 nominal values. S = 250 MPa deformed bar (D250N, ductility Class N). L = mesh wire (Class L; SL92 = L9 @ 200, RL1218 = L12 @ 200 main). Edit or add rows below; the calc tabs read this list.",
               {"B": 10, "C": 10, "D": 10, "E": 10, "F": 10, "G": 40})
for j, t in enumerate(["Bar", "db (mm)", "Ab (mm²)", "fsy (MPa)", "Class", "Note"]):
    c = D_.cell(4, 2 + j); c.value = t; c.font = font(8, True); c.fill = fill(HDR_FILL); c.border = BOX; c.alignment = CEN
BAR0 = 5; BAR_LAST = BAR0 + len(BARS) + 5
for i in range(BAR_LAST - BAR0 + 1):
    row = BARS[i] + [notes.get(BARS[i][0], "")] if i < len(BARS) else [None] * 6
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
wb.defined_names["EcFc"] = DefinedName("EcFc", attr_text="Bars!$I$5:$I$12")
wb.defined_names["EcVal"] = DefinedName("EcVal", attr_text="Bars!$J$5:$J$12")
wb.defined_names["BarList"] = DefinedName("BarList", attr_text=f"Bars!$B${BAR0}:$B${BAR_LAST}")
BARTAB = f"Bars!$B${BAR0}:$F${BAR_LAST}"
def barv(ref, col):  # col 2 db, 3 Ab, 4 fsy, 5 class
    return f"VLOOKUP({ref},{BARTAB},{col},FALSE)"

# =========================== DATA TAB: Crack Control ===========================
CC = wb.create_sheet("Crack Control")
setup_data_tab(CC, "Crack control – maximum steel stress (AS 3600-2018 incl. Amdt 2)",
               "σscr must not exceed the larger of Table A (largest bar diameter in the tension zone) and Table B (largest c/c spacing; "
               "bars under half the largest diameter ignored). Linear interpolation between rows; outside the table the end row is used.",
               {c: 9 for c in "BCDEFGHIJKLMNOP"})
def cc_table(r0, c0, title, head, rows):
    c = CC.cell(r0, c0); c.value = title; c.font = font(9, True, color=GREEN)
    for j, t in enumerate(head):
        x = CC.cell(r0 + 1, c0 + j); x.value = t; x.font = font(7, True); x.fill = fill(HDR_FILL); x.border = BOX
        x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    CC.row_dimensions[r0 + 1].height = 26
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            x = CC.cell(r0 + 2 + i, c0 + j); x.value = v; x.font = font(8, color="000000"); x.fill = fill(IN_FILL); x.border = BOX; x.alignment = CEN
    return r0 + 2, r0 + 1 + len(rows)
from openpyxl.utils import get_column_letter as _L
def cc_names(name, r1, r2, c0, ncol):
    wb.defined_names[name + "_X"] = DefinedName(name + "_X", attr_text=f"'Crack Control'!${_L(c0)}${r1}:${_L(c0)}${r2}")
    wb.defined_names[name] = DefinedName(name, attr_text=f"'Crack Control'!${_L(c0+1)}${r1}:${_L(c0+ncol)}${r2}")
r1, r2 = cc_table(5, 2, "Table 8.6.2.2(A) – beams, by bar diameter", ["db (mm)", "w' 0.2", "w' 0.3", "w' 0.4"],
                  [[10, 190, 265, 335], [12, 175, 245, 305], [16, 155, 215, 270], [20, 140, 195, 240], [24, 125, 175, 215],
                   [28, 115, 160, 200], [32, 105, 150, 185], [36, 100, 140, 175], [40, 90, 130, 165]])
cc_names("CC_BA", r1, r2, 2, 3)
r1, r2 = cc_table(5, 7, "Table 8.6.2.2(B) = 9.5.2.1(B) – by c/c spacing", ["s (mm)", "w' 0.2", "w' 0.3", "w' 0.4"],
                  [[50, 200, 300, 400], [100, 170, 270, 360], [150, 155, 245, 330], [200, 145, 225, 300], [250, 135, 210, 280], [300, 125, 200, 260]])
cc_names("CC_B", r1, r2, 7, 3)
r1, r2 = cc_table(18, 2, "Table 9.5.2.1(A) – slabs, by bar diameter and overall depth Ds",
                  ["db (mm)", "0.2, Ds ≤ 300", "0.2, Ds > 300", "0.3, Ds ≤ 300", "0.3, Ds > 300", "0.4, Ds ≤ 300", "0.4, Ds > 300"],
                  [[6, 210, 230, 290, 315, 365, 390], [8, 195, 210, 260, 290, 340, 360], [10, 180, 190, 240, 265, 315, 335],
                   [12, 165, 175, 225, 245, 290, 305], [16, 150, 155, 205, 215, 260, 270], [20, 140, 140, 190, 195, 235, 240],
                   [24, 125, 125, 175, 175, 215, 215]])
cc_names("CC_SA", r1, r2, 2, 6)
c = CC.cell(28, 2); c.value = ("Checked against RAM Concept 8.x (AS 3600-2018 Amdt 2): slab w' 0.3 mm – db 15.96 → 205.2, s 156.3 → 242.5, "
                               "s 230.8 → 215.8 MPa (linear interpolation). Note RAM takes s = b / (bars + 1); the calc tabs use b / bars.")
c.font = font(7, i=True, color=GREY)

# Add-in version check: the template records the add-in version it was built with (from the Excel project);
# an older or missing add-in gets a red message under the title bar and UPDATE ADD-IN in the status box.
import re as _re
_fsproj = open(os.path.join(HERE, "..", "src", "OAR.Concrete.Excel", "OAR.Concrete.Excel.fsproj"), encoding="utf-8").read()
TEMPLATE_VERSION = _re.search(r"<Version>([0-9]+\.[0-9]+\.[0-9]+)</Version>", _fsproj).group(1)
_maj, _min, _pat = (int(v) for v in TEMPLATE_VERSION.split("."))

# =========================== CALC TAB ===========================
def build_tab(name, d):
    ws = wb.create_sheet(name)
    N = {}  # named cells → absolute refs, used as «name» in formulas
    def nm(n, ref): N[n] = "$" + "".join(ch for ch in ref if ch.isalpha()) + "$" + "".join(ch for ch in ref if ch.isdigit())
    # depth of a single mat from the top face: D/2, or at the top / bottom cover (+ ligature for a beam)
    _lig = f'IF(«type»="Beam",{barv("«lig»", 2)},0)'
    _dbm = f'IFERROR({barv("«tmb»", 2)},0)'
    MATD = (f'IF(«layout»="Single mat – at top cover",«tcov»+{_lig}+{_dbm}/2,'
            f'IF(«layout»="Single mat – at bottom cover",«D»-«bcov»-{_lig}-{_dbm}/2,«D»/2))')
    def Fx(s):
        if not isinstance(s, str): return s
        s = s.replace("«LO»", 'LEFT(«lmode»,1)="L"').replace("«LEG»", '«wmax»="Legacy"').replace("«CEN»", 'LEFT(«layout»,6)="Single"').replace("«MATD»", MATD)
        for k in sorted(N, key=len, reverse=True): s = s.replace(f"«{k}»", N[k])
        assert "«" not in s, s
        return s
    pend = []  # (ref, formula) resolved at the end
    arrays = []  # ranges written as fixed-size array formulas
    def lab(ref, t, **k): return label(ws, ref, t, **k)
    def IN(ref, v=None, nf=None, al=CEN, name=None):
        c = inp(ws, ref, v, nf, al)
        if name: nm(name, ref.split(":")[0])
        return c
    def CA(ref, f, nf=None, al=CEN, name=None):
        c = calc(ws, ref, None, nf, al); pend.append((ref.split(":")[0], f))
        if name: nm(name, ref.split(":")[0])
        return c
    def U(ref, t): return unit(ws, ref, t)
    def sym(r, t):
        ws.merge_cells(f"B{r}:C{r}"); lab(f"B{r}", t, al=CEN)
    def desc(r, t):
        if isinstance(t, str) and t.startswith("="): pend.append((f"D{r}", t))
        else: ws[f"D{r}"].value = t
        ws[f"D{r}"].font = font(8, color=GREY); ws[f"D{r}"].alignment = LEFT

    setup_calc_tab(ws, LOGO, '="RC FLEXURAL MEMBER (AS 3600-2018)  –  "&IF(G3="","(ref)",G3)&IF(I3="","","  ·  "&I3)',
                   "MEMBER", name, "Copy this tab for additional slabs / beams.")
    ws["I3"].value = d["loc"]
    location_frame(ws, "Member")
    # version check cells (column V, beside the location frame, not printed; same-sheet so copied tabs keep working)
    ws.column_dimensions["U"].width = 3; ws.column_dimensions["V"].width = 12
    lab("V2", "Add-in check", sz=7, color=GREY, i=True)
    for rr, t, v in ((3, "template", TEMPLATE_VERSION),
                     (4, "add-in", '=IFERROR(SUBSTITUTE(OAR.CONCRETE.VERSION(),"OAR Concrete ",""),"")'),
                     (6, "template no.", _maj * 10000 + _min * 100 + _pat)):
        lab(f"W{rr}", t, sz=7, color=GREY); c = ws[f"V{rr}"]; c.value = v; c.font = font(7, color=GREY); c.alignment = LEFT
    p1 = 'FIND(".",$V$4)'; p2 = f'FIND(".",$V$4,{p1}+1)'
    ws["V5"].value = f'=IFERROR(LEFT($V$4,{p1}-1)*10000+MID($V$4,{p1}+1,{p2}-{p1}-1)*100+MID($V$4,{p2}+1,9),0)'
    ws["V7"].value = '=AND($V$4<>"",$V$5<$V$6)'
    lab("W5", "add-in no.", sz=7, color=GREY); lab("W7", "add-in older?", sz=7, color=GREY)
    for rr in (5, 7): ws[f"V{rr}"].font = font(7, color=GREY); ws[f"V{rr}"].alignment = LEFT
    nm("tver", "V3"); nm("aver", "V4"); nm("aold", "V7")
    c = ws["E7"]; c.font = font(7, i=True, color=GREY)
    # B7 normally holds the copy-tab note; it carries the warning when the add-in is missing or older than the template
    pend.append(("B7", '=IF(«aver»="","⚠ OAR Concrete add-in not loaded – results below are not calculated",IF(«aold»,"⚠ Add-in "&«aver»&" is older than this template ("&«tver»&") – update the add-in and restart Excel","Copy this tab for additional slabs / beams.   ·   Calculated with OAR Concrete "&«aver»))'))

    # ------------------------------------------------------------------ inputs (page 1, below the checks)
    r = 33; heading(ws, r, "Member and Concrete")
    lab("B34", "Member type"); IN("E34", d["type"], name="type"); lab("H34", "Slab: strip of width b, mats as bar @ spacing", sz=7, color=GREY)
    lab("B35", "Width b"); IN("E35", d["b"], "0", name="b"); U("F35", "mm"); lab("H35", "slab: 1000 = per metre width", sz=7, color=GREY)
    lab("B36", "Overall depth D"); IN("E36", d["D"], "0", name="D"); U("F36", "mm")
    lab("B37", "Concrete strength f'c"); IN("E37", d["fc"], "0", name="fc"); U("F37", "MPa")
    lab("B38", "Stress block (ULS)"); IN("E38:F38", d["model"], name="model")
    lab("B39", "Ec override"); IN("E39", d.get("ec"), "0", name="ecin"); U("F39", "MPa"); lab("H39", "blank = Table 3.1.2", sz=7, color=GREY)
    lab("B40", "Creep factor for SLS stresses, φcc"); IN("E40", d.get("phicc", 0), "0.0", name="phicc"); lab("H40", "0 = short-term n = Es / Ec", sz=7, color=GREY)

    lab("B41", "Reinforcement layout"); IN("E41:F41", d.get("layout", "Top & bottom"), name="layout")
    lab("H41", "single mat: Top column = mat + extras above, Bottom = extras below", sz=7, color=GREY)
    r = 42; heading(ws, r, "Reinforcement")
    table_header(ws, 43, [("E43", "Top"), ("F43", "Bottom")])
    pend.append(("E43", '=IF(«CEN»,"Mat / above","Top")')); pend.append(("F43", '=IF(«CEN»,"Below mat","Bottom")'))
    lab("H45", 'bar "None" = no mat on that face', sz=7, color=GREY)
    rows = [
        (44, "Cover (slab: to bar; beam: to ligature)", "cov", "0", "mm"),
        (45, '=IF(«type»="Slab","Main mat – bar","Main bars – bar")', "mb", None, ""),
        (46, '=IF(«type»="Slab","Main mat – spacing","Main bars – number")', "mn", "0", '=IF(«type»="Slab","mm","no.")'),
        (47, "Extra bars – number (within b)", "xn", "0", "no."),
        (48, "Extra bars – bar", "xb", None, ""),
        (49, "Extra bars – position", "xp", None, ""),
        (50, "Clear gap, layer 1 to layer 2", "gap", "0", "mm"),
    ]
    for rr, t, key, nf, u in rows:
        if t.startswith("="): pend.append((f"B{rr}", t)); ws[f"B{rr}"].font = font(9)
        else: lab(f"B{rr}", t)
        IN(f"E{rr}", d["top"][key], nf, name="t" + key); IN(f"F{rr}", d["bot"][key], nf, name="b" + key)
        if u.startswith("="): pend.append((f"G{rr}", u)); ws[f"G{rr}"].font = font(9, color=GREY)
        else: U(f"G{rr}", u)
    lab("H49", "Layer 2 = second row inside layer 1", sz=7, color=GREY)
    lab("B51", "Ligature bar (beam)"); IN("E51", d.get("lig", "N10"), name="lig"); lab("H51", "beam only; adds to cover", sz=7, color=GREY)
    lab("B52", "Slab supported on (αb, Cl 9.1.1)"); IN("E52:F52", d.get("support", "Beams or walls"), name="support")
    lab("B53", "Class L reinforcement, φ"); IN("E53:F53", d.get("lmode", "Whole section 0.65"), name="lmode")
    lab("H53", "L steel only: Class L As × 0.65/φ (judgement)", sz=7, color=GREY)

    r = 54; heading(ws, r, "Design Actions")
    table_header(ws, 55, [("E55", "Sagging +"), ("F55", "Hogging −")])
    lab("H55", "combination", sz=7, color=GREY)
    for rr, t, key, comb in ((56, "M*  (ULS)", "Mu", "1.2G + 1.5Q (or as governs)"),
                             (57, "Ms*  (crack control)", "Ms", "G + ψs Q  (e.g. G + 0.7Q)"),
                             (58, "Ms1*  (σs ≤ 0.8 fsy)", "Ms1", "G + Q")):
        lab(f"B{rr}", t); IN(f"E{rr}", d["sag"][key], "0.0", name="s" + key); IN(f"F{rr}", d["hog"][key], "0.0", name="h" + key)
        U(f"G{rr}", "kNm"); lab(f"H{rr}", comb, sz=7, color=GREY)
    note(ws, 59, "Enter magnitudes. Sagging = bottom face in tension, hogging = top face in tension.")

    r = 60; heading(ws, r, "Crack Control and Shrinkage")
    table_header(ws, 61, [("E61", "Sagging +"), ("F61", "Hogging −")])
    lab("B62", "σs.max override (Ms*)"); IN("E62", d.get("ovs"), "0", name="ovs"); IN("F62", d.get("ovh"), "0", name="ovh"); U("G62", "MPa")
    lab("H62", "blank = from the tables (calcs)", sz=7, color=GREY)
    lab("B63", "Shrinkage & temp.: degree of control"); IN("E63:F63", d.get("control", "Moderate"), name="control"); lab("H63", "slab only", sz=7, color=GREY)
    lab("B64", "Restraint factor (0 = free, 1 = fully restrained)"); IN("E64", d.get("restraint", 0.75), "0.00", name="restr")
    lab("B65", "Crack width limit w'max (σs tables)"); IN("E65:F65", d.get("wmax", "0.3 mm"), name="wmax")
    lab("H65", "Legacy = previous OAR sheet (2009 fits)", sz=7, color=GREY)

    # ------------------------------------------------------------------ calculations (page 2)
    # Reinforcement layers table → the add-in's bars argument (As | depth | fsy)
    r = 66; heading(ws, r, "Reinforcement Layers")
    note(ws, 67, "One row per layer across b; depth to the bar centre from the top face. Slab mat As = Ab × b / s. Beam cover includes the ligature.")
    table_header(ws, 68, [("B68:C68", "Layer"), ("D68", "Description"), ("E68", "As (mm²)"), ("F68", "depth (mm)"),
                          ("G68", "fsy (MPa)"), ("H68", "db (mm)"), ("I68", "Class"), ("J68:K68", "bars in b")])
    # effective cover to the outer bar = cover (+ ligature for a beam)
    LIG = f'IF(«type»="Beam",{barv("«lig»", 2)},0)'
    def area_main(f):
        a = f'IF(OR(«{f}mb»="",«{f}mb»="None",N(«{f}mn»)=0),0,IF(«type»="Slab",{barv(f"«{f}mb»", 3)}*«b»/«{f}mn»,«{f}mn»*{barv(f"«{f}mb»", 3)}))'
        return f'IF(«CEN»,0,{a})' if f == "b" else a
    def db(key): return f'IFERROR({barv(key, 2)},0)'
    layers = [
        (69, "Top – main", "t", "m"), (70, "Top – extra", "t", "x"), (71, "Bottom – main", "b", "m"), (72, "Bottom – extra", "b", "x")]
    cen_lab = {69: "Single mat", 70: "Extra – above mat", 71: "Bottom – main", 72: "Extra – below mat"}
    for rr, t, f, k in layers:
        ws.merge_cells(f"B{rr}:C{rr}"); lab(f"B{rr}", t, sz=8); pend.append((f"B{rr}", f'=IF(«CEN»,"{cen_lab[rr]}","{t}")'))
        if k == "m":
            pend.append((f"D{rr}", f'=IF(E{rr}=0,IF(AND(«CEN»,"{f}"="b"),"– (single mat above)","none"),«{f}mb»&IF(«type»="Slab"," @ "&«{f}mn»&" ("&TEXT({barv(f"«{f}mb»", 3)}*1000/«{f}mn»,"0")&" mm²/m)"," × "&«{f}mn»))'))
            area = area_main(f)
            dep_off = f'(«{f}cov»+{LIG}+{db(f"«{f}mb»")}/2)'
            nbars = f'IF(E{rr}=0,0,IF(«type»="Slab",«b»/«{f}mn»,«{f}mn»))'
        else:
            pend.append((f"D{rr}", f'=IF(E{rr}=0,"none",«{f}xn»&" × "&«{f}xb»&IF(«CEN»,IF("{f}"="t"," on top of mat"," under mat"),IF(«{f}xp»="Layer 2"," (layer 2)","")))'))
            area = f'IF(OR(N(«{f}xn»)=0,«{f}xb»="",«{f}xb»="None"),0,«{f}xn»*{barv(f"«{f}xb»", 3)})'
            dep_off = (f'(«{f}cov»+{LIG}+IF(«{f}xp»="Layer 2",{db(f"«{f}mb»")}+«{f}gap»+{db(f"«{f}xb»")}/2,{db(f"«{f}xb»")}/2))')
            nbars = f'IF(E{rr}=0,0,N(«{f}xn»))'
        ws[f"D{rr}"].font = font(8, color=GREY)
        CA(f"E{rr}", f"={area}", "0", name=f"As{f}{k}")
        normal = dep_off if f == "t" else f"«D»-{dep_off}"
        central = {("t", "m"): "«MATD»", ("t", "x"): f'«MATD»-{db("«tmb»")}/2-{db("«txb»")}/2',
                   ("b", "m"): "«D»/2", ("b", "x"): f'«MATD»+{db("«tmb»")}/2+{db("«bxb»")}/2'}[(f, k)]
        CA(f"F{rr}", f"=IF(«CEN»,{central},{normal})", "0.0", name=f"dep{f}{k}")
        barref = f"«{f}mb»" if k == "m" else f"«{f}xb»"
        CA(f"G{rr}", f'=IF(E{rr}=0,"",{barv(barref, 4)})', "0", name=f"fsy{f}{k}")
        CA(f"H{rr}", f'=IF(E{rr}=0,"",{barv(barref, 2)})', "0.0", name=f"db{f}{k}")
        CA(f"I{rr}", f'=IF(E{rr}=0,"",{barv(barref, 5)})', name=f"cl{f}{k}")
        CA(f"J{rr}:K{rr}", f"={nbars}", "0.0", name=f"n{f}{k}")
    lab("B73", "Σ", al=CEN); ws.merge_cells("B73:C73")
    desc(73, "totals: top | bottom")
    CA("E73", "=E69+E70", "0", name="Ast_top"); CA("F73", "=E71+E72", "0", name="Ast_bot")
    U("G73", "mm²")
    BARS_ARG = "$E$69:$G$72"

    # Material properties
    r = 75; heading(ws, r, "Concrete Properties")
    note(ws, 76, "α2, γ: Cl 8.1.3.  Ec: Table 3.1.2 (interpolated) unless entered.  f'ct.f = 0.6√f'c, Cl 3.1.1.3.  Es = 200 000 MPa, εcu = 0.003.")
    props = [
        (77, "α2", "0.85 − 0.0015 f'c ≥ 0.67", "=MAX(0.67,0.85-0.0015*«fc»)", "0.000", "", "a2"),
        (78, "γ", "0.97 − 0.0025 f'c ≥ 0.67", "=MAX(0.67,0.97-0.0025*«fc»)", "0.000", "", "gam"),
        (79, "Ec", '=IF(N(«ecin»)>0,"entered","Table 3.1.2, interpolated")',
         "=IF(N(«ecin»)>0,«ecin»,INDEX(EcVal,MIN(MATCH(«fc»,EcFc,1),7))+(«fc»-INDEX(EcFc,MIN(MATCH(«fc»,EcFc,1),7)))*(INDEX(EcVal,MIN(MATCH(«fc»,EcFc,1),7)+1)-INDEX(EcVal,MIN(MATCH(«fc»,EcFc,1),7)))/(INDEX(EcFc,MIN(MATCH(«fc»,EcFc,1),7)+1)-INDEX(EcFc,MIN(MATCH(«fc»,EcFc,1),7))))",
         "0", "MPa", "Ec"),
        (80, "Ec.eff", "Ec / (1 + φcc)  (SLS stresses)", "=«Ec»/(1+N(«phicc»))", "0", "MPa", "Eceff"),
        (81, "n", "Es / Ec.eff", "=200000/«Eceff»", "0.00", "", "nmod"),
        (82, "f'ct.f", "0.6 √f'c", "=0.6*SQRT(«fc»)", "0.00", "MPa", "fctf"),
    ]
    for rr, s_, dsc, f, nf, u, key in props:
        sym(rr, s_); desc(rr, dsc); CA(f"E{rr}", f, nf, name=key); U(f"F{rr}", u)

    # ULS strain compatibility
    r = 84; heading(ws, r, "Ultimate Strength – strain compatibility (OAR Concrete add-in)")
    note(ws, 85, "OAR.CONCRETE.ULT: plane sections, εcu = 0.003 at the extreme fibre, no concrete tension, steel elastic–plastic (Es 200 GPa), displaced concrete deducted at compression bars.")
    table_header(ws, 86, [("E86", "Sagging +"), ("F86", "Hogging −")])
    ULT = [("dn", "neutral axis depth", "0.0", "mm"), ("kuo", "dn / do", "0.000", ""), ("ku", "dn / d", "0.000", ""),
           ("do", "to outermost tension bar", "0.0", "mm"), ("d", "to tension resultant", "0.0", "mm"),
           ("εst", "strain, outermost tension bar", "0.0000", ""), ("σst", "stress, outermost tension bar", "0", "MPa"),
           ("σsc", "max compression bar stress", "0", "MPa"), ("Cc", "concrete compression", "0.0", "kN"),
           ("Cs", "compression steel (net)", "0.0", "kN"), ("T", "tension steel", "0.0", "kN"),
           ("z", "lever arm Muo / T", "0.0", "mm"), ("Muo", '=IF(«LO»,"ultimate moment, Class L As × kL","ultimate moment, N = 0")', "0.0", "kNm"),
           ("φ.L", "φ on reduced section (L steel only)", "0.000", ""), ("kL", "Class L area factor 0.65 / φ", "0.000", ""),
           ("εsL", "max Class L tensile strain", "0.0000", "")]
    U0 = 87
    for i, (s_, dsc, nf, u) in enumerate(ULT):
        rr = U0 + i; sym(rr, s_); desc(rr, dsc); calc(ws, f"E{rr}", None, nf); calc(ws, f"F{rr}", None, nf); U(f"G{rr}", u)
    U1 = U0 + len(ULT) - 1
    def ult(sense, model="«model»"): return f'OAR.CONCRETE.ULT("Rect",«b»:«D»,{BARS_ARG},«fc»,{model},"{sense}",,«Ec»,$I$69:$I$72,«lmode»)'
    pend.append((f"E{U0}", "=" + ult("Sagging"))); arrays.append(f"E{U0}:E{U1}")
    pend.append((f"F{U0}", "=" + ult("Hogging"))); arrays.append(f"F{U0}:F{U1}")
    nm("kuo_s", f"E{U0+1}"); nm("kuo_h", f"F{U0+1}"); nm("Muo_s", f"E{U0+12}"); nm("Muo_h", f"F{U0+12}")
    nm("phiE_s", f"E{U0+13}"); nm("phiE_h", f"F{U0+13}"); nm("kL_s", f"E{U0+14}"); nm("kL_h", f"F{U0+14}")
    nm("epsL_s", f"E{U0+15}"); nm("epsL_h", f"F{U0+15}")
    r = U1 + 1
    sym(r, "class"); desc(r, "tension face (any L bar → L)")
    CA(f"E{r}", '=IF(OR(«FIsm»="L",«FIsx»="L"),"L","N")', name="cls_s"); CA(f"F{r}", '=IF(OR(«FIhm»="L",«FIhx»="L"),"L","N")', name="cls_h")
    r += 1; sym(r, "φ"); desc(r, '=IF(«LO»,"Class N formula, Class L As × kL","Tbl 2.2.2 (Class L: 0.65)")')
    for col, s in (("E", "s"), ("F", "h")):
        CA(f"{col}{r}", f'=IF(«LO»,«phiE_{s}»,IF(«cls_{s}»="L",0.65,MIN(0.85,MAX(0.65,1.24-13*«kuo_{s}»/12))))', "0.000", name=f"phi_{s}")
    r += 1; sym(r, "φMuo"); desc(r, "design capacity")
    CA(f"E{r}", "=«phi_s»*«Muo_s»", "0.0", name="phiM_s"); CA(f"F{r}", "=«phi_h»*«Muo_h»", "0.0", name="phiM_h"); U(f"G{r}", "kNm")
    r += 1; sym(r, "Muo'"); desc(r, '=IF(«model»="Rectangular","Curvilinear block, for comparison","Rectangular block, for comparison")')
    other = '=INDEX(' + ult("Sagging", 'IF(«model»="Rectangular","Curvilinear","Rectangular")') + ',13,1)'
    CA(f"E{r}", other, "0.0", name="Muo2_s")
    CA(f"F{r}", other.replace('"Sagging"', '"Hogging"'), "0.0", name="Muo2_h"); U(f"G{r}", "kNm")
    r += 1; sym(r, "φMuo'"); desc(r, "same φ, for comparison (info)")
    CA(f"E{r}", "=«phi_s»*«Muo2_s»", "0.0", name="phiM2_s"); CA(f"F{r}", "=«phi_h»*«Muo2_h»", "0.0", name="phiM2_h"); U(f"G{r}", "kNm")
    ULT_END = r
    pend.append((f"B{ULT_END + 1}", '=IF(«LO»,"L steel only: Class L areas × kL = 0.65/φ, φ from the Class N formula on that section, iterated until φ and kL agree. Engineering judgement, not a Table 2.2.2 provision.","Whole section: φ = 0.65 if any tension-face bar is Class L (Table 2.2.2).")')); ws[f"B{ULT_END + 1}"].font = font(7, i=True, color=GREY)

    # Tension reinforcement for each direction: normally the bottom face (sagging) / top face (hogging).
    # A face with no steel uses the other face's layers (single layer at cover); a central mat serves both directions,
    # with the extras below it (sagging) or above it (hogging).
    r = ULT_END + 2; heading(ws, r, "Tension Reinforcement by Direction")
    r += 1; note(ws, r, "Used for φ class, minimum strength and crack control. A face with no steel takes the other face's layers; a single mat serves both directions.")
    r += 1
    table_header(ws, r, [(f"B{r}:C{r}", "Direction"), (f"D{r}", "Source"), (f"E{r}", "As (mm²)"), (f"F{r}", "depth (mm)"),
                         (f"G{r}", "fsy (MPa)"), (f"H{r}", "db (mm)"), (f"I{r}", "Class"), (f"J{r}:K{r}", "bars in b")])
    useO = {"s": "AND(NOT(«CEN»),$E$71+$E$72=0)", "h": "AND(NOT(«CEN»),$E$69+$E$70=0)"}
    rows_ = {"s": (71, 72, 69, 70, "b", "t"), "h": (69, 70, 71, 72, "t", "b")}
    for d_, lab_ in (("s", "Sagging"), ("h", "Hogging")):
        om, ox, pm, px, fo, fp = rows_[d_]
        for k, rr_own, rr_oth, rr_cen in (("m", om, pm, 69), ("x", ox, px, ox)):
            r += 1; ws.merge_cells(f"B{r}:C{r}"); lab(f"B{r}", f"{lab_} – {'main' if k == 'm' else 'extra'}", sz=8)
            own_t = {"s": "bottom", "h": "top"}[d_]; oth_t = {"s": "top", "h": "bottom"}[d_]
            kind = "main" if k == "m" else "extras"
            cen_t = "single mat" if k == "m" else ("extras below mat" if d_ == "s" else "extras above mat")
            pend.append((f"D{r}", f'=IF(«CEN»,"{cen_t}",IF({useO[d_]},"{oth_t} {kind} (no {own_t} steel)","{own_t} {kind}"))'))
            ws[f"D{r}"].font = font(8, color=GREY)
            for col, nf in (("E", "0"), ("F", "0.0"), ("G", "0"), ("H", "0.0"), ("I", None), ("J", "0.0")):
                ref = f"{col}{r}:K{r}" if col == "J" else f"{col}{r}"
                CA(ref, f'=IF(«CEN»,{col}{rr_cen},IF({useO[d_]},{col}{rr_oth},{col}{rr_own}))', nf, name=f"F{col}{d_}{k}")
            if k == "x":   # does this extra row sit in layer 1 (counts for spacing)?
                xp_own, xp_oth = f"«{fo}xp»", f"«{fp}xp»"
                nm(f"L1_{d_}", f"L{r}")
                ws[f"L{r}"].value = f'=OR(«CEN»,IF({useO[d_]},{xp_oth},{xp_own})="Layer 1")'
                ws[f"L{r}"].font = font(7, color=GREY); pend.append((f"L{r}", ws[f"L{r}"].value))
                cov_own, cov_oth = f"«{fo}cov»", f"«{fp}cov»"
                nm(f"cov_{d_}", f"M{r}")
                pend.append((f"M{r}", f'=IF({useO[d_]},{cov_oth},{cov_own})')); ws[f"M{r}"].font = font(7, color=GREY)
    TEN_END = r

    # Minimum strength
    r = TEN_END + 2; heading(ws, r, "Minimum Strength (Cl 8.1.6.1 / 9.1.1)")
    r += 1; note(ws, r, "(Muo)min = 1.2 Z f'ct.f (no prestress, N = 0). Deemed to comply if Ast ≥ αb (D/d)² f'ct.f / fsy · b d.")
    table_header(ws, r + 1, [("E%d" % (r + 1), "Sagging +"), ("F%d" % (r + 1), "Hogging −")])
    r += 2; sym(r, "Z"); desc(r, "b D² / 6, uncracked gross"); CA(f"E{r}", "=«b»*«D»^2/6/1E6", "0.00", name="Z"); CA(f"F{r}", "=«Z»", "0.00"); U(f"G{r}", "×10⁶ mm³")
    r += 1; sym(r, "(Muo)min"); desc(r, "1.2 Z f'ct.f"); CA(f"E{r}", "=1.2*«Z»*«fctf»", "0.0", name="Mmin"); CA(f"F{r}", "=«Mmin»", "0.0"); U(f"G{r}", "kNm")
    r += 1; sym(r, "Ast"); desc(r, "tension steel, total"); CA(f"E{r}", "=«FEsm»+«FEsx»", "0", name="Ast_s"); CA(f"F{r}", "=«FEhm»+«FEhx»", "0", name="Ast_h"); U(f"G{r}", "mm²")
    r += 1; sym(r, "d"); desc(r, "tension steel, from compr. face")
    CA(f"E{r}", "=IF(«Ast_s»=0,0,(«FEsm»*«FFsm»+«FEsx»*N(«FFsx»))/«Ast_s»)", "0.0", name="d_s")
    CA(f"F{r}", "=IF(«Ast_h»=0,0,«D»-(«FEhm»*«FFhm»+«FEhx»*N(«FFhx»))/«Ast_h»)", "0.0", name="d_h"); U(f"G{r}", "mm")
    r += 1; sym(r, "fsy"); desc(r, "tension steel (lowest)")
    CA(f"E{r}", '=MIN(IF(«FEsm»>0,«FGsm»,9999),IF(«FEsx»>0,«FGsx»,9999))', "0", name="fsy_s"); CA(f"F{r}", '=MIN(IF(«FEhm»>0,«FGhm»,9999),IF(«FEhx»>0,«FGhx»,9999))', "0", name="fsy_h"); U(f"G{r}", "MPa")
    r += 1; sym(r, "αb"); desc(r, '=IF(«type»="Beam","rectangular beam 0.20","slab on "&LOWER(«support»))')
    CA(f"E{r}", '=IF(«type»="Beam",0.2,IF(«support»="Columns",0.24,0.19))', "0.00", name="alb"); CA(f"F{r}", "=«alb»", "0.00")
    r += 1; sym(r, "Ast.min"); desc(r, "αb (D/d)² f'ct.f / fsy · b d")
    # deemed-to-comply only where the tension steel is on the tension face (or a central mat), not borrowed from the other face
    CA(f"E{r}", f'=IF(OR(«d_s»=0,{useO["s"]}),"",«alb»*(«D»/«d_s»)^2*«fctf»/«fsy_s»*«b»*«d_s»)', "0", name="Amin_s")
    CA(f"F{r}", f'=IF(OR(«d_h»=0,{useO["h"]}),"",«alb»*(«D»/«d_h»)^2*«fctf»/«fsy_h»*«b»*«d_h»)', "0", name="Amin_h"); U(f"G{r}", "mm²")
    MIN_END = r

    # Cracked section
    r = MIN_END + 2; heading(ws, r, "Cracked Section Stresses (OAR Concrete add-in)")
    r += 1; note(ws, r, "OAR.CONCRETE.CRACKED: elastic, no concrete tension, n = Es / Ec.eff, (n − 1) at compression bars. Stresses as calculated (not capped at fsy).")
    CR0 = r + 1
    table_header(ws, CR0, [(f"E{CR0}", "Ms* +"), (f"F{CR0}", "Ms* −"), (f"G{CR0}", "Ms1* +"), (f"H{CR0}", "Ms1* −")])
    CRK = [("dn", "neutral axis depth", "0.0", "mm"), ("Icr", "cracked second moment", "0.0", "×10⁶ mm⁴"),
           ("σscr", "tension steel stress (outermost)", "0", "MPa"), ("σsc", "compression steel stress", "0", "MPa"),
           ("σc", "concrete, extreme fibre", "0.0", "MPa")]
    C1 = CR0 + 1
    for i, (s_, dsc, nf, u) in enumerate(CRK):
        rr = C1 + i; sym(rr, s_); desc(rr, dsc)
        for col in "EFGH": calc(ws, f"{col}{rr}", None, nf)
        U(f"I{rr}", u)
    C2 = C1 + len(CRK) - 1
    def crk(m, sense): return f'=OAR.CONCRETE.CRACKED("Rect",«b»:«D»,{BARS_ARG},«Eceff»,N({m}),"{sense}")'
    for col, m, s in (("E", "«sMs»", "Sagging"), ("F", "«hMs»", "Hogging"), ("G", "«sMs1»", "Sagging"), ("H", "«hMs1»", "Hogging")):
        pend.append((f"{col}{C1}", crk(m, s))); arrays.append(f"{col}{C1}:{col}{C2}")
    nm("sig_s", f"E{C1+2}"); nm("sig_h", f"F{C1+2}"); nm("sig1_s", f"G{C1+2}"); nm("sig1_h", f"H{C1+2}")
    CRK_END = C2

    # Crack control limits
    r = CRK_END + 2; heading(ws, r, "Crack Control Limits")
    r += 1; note(ws, r, '=IF(«LEG»,"Legacy OAR limits (previous slab sheet): fits to AS 3600-2009 Tbl 8.6.1 / 9.4.1 – A by bar diameter, B by bar spacing, the larger applies.",'
                        'IF(«type»="Slab","Cl 9.5.2.1: σscr ≤ larger of Tbl 9.5.2.1(A) (largest db, by Ds) and 9.5.2.1(B) (largest c/c spacing), w\'max = "&«wmax»&".",'
                        '"Cl 8.6.2.2: σscr ≤ larger of Tbl 8.6.2.2(A) (largest db) and 8.6.2.2(B) (largest c/c spacing), w\'max = "&«wmax»&"."))')
    pend.append((f"B{r}", ws[f"B{r}"].value)); ws[f"B{r}"].value = None
    r += 1; note(ws, r, "Spacing ignores bars smaller than half the largest diameter on the face (Cl 9.5.2.1(b)). Tables on the Crack Control tab; linear interpolation, end rows beyond the table.")
    table_header(ws, r + 1, [(f"E{r+1}", "Sagging +"), (f"F{r+1}", "Hogging −")])
    r += 2; sym(r, "db"); desc(r, "largest tension bar")
    CA(f"E{r}", "=MAX(N(«FHsm»),N(«FHsx»))", "0.0", name="dbx_s"); CA(f"F{r}", "=MAX(N(«FHhm»),N(«FHhx»))", "0.0", name="dbx_h"); U(f"G{r}", "mm")
    r += 1; sym(r, "s"); desc(r, '=IF(«type»="Slab","b / (tension bars in layer 1)","clear width / (n − 1), layer 1")')
    def spacing(f, s_):
        cnt_m = f'IF(N(«FH{s_}m»)>=0.5*«dbx_{s_}»,N(«FJ{s_}m»),0)'
        cnt_x = f'IF(OR(NOT(«L1_{s_}»),N(«FH{s_}x»)<0.5*«dbx_{s_}»),0,N(«FJ{s_}x»))'
        n1 = f'({cnt_m}+{cnt_x})'
        return (f'=IF({n1}<=0,0,IF(«type»="Slab",«b»/{n1},IF({n1}<=1,0,(«b»-2*(«cov_{s_}»+{barv("«lig»", 2)})-«dbx_{s_}»)/({n1}-1))))')
    CA(f"E{r}", spacing("b", "s"), "0", name="sp_s"); CA(f"F{r}", spacing("t", "h"), "0", name="sp_h"); U(f"G{r}", "mm")
    def interp(xs, tbl, col, x):
        xc = f"MIN(MAX({x},INDEX({xs},1)),INDEX({xs},ROWS({xs})))"
        i = f"MIN(MATCH({xc},{xs},1),ROWS({xs})-1)"
        return (f"(INDEX({tbl},{i},{col})+({xc}-INDEX({xs},{i}))*(INDEX({tbl},{i}+1,{col})-INDEX({tbl},{i},{col}))"
                f"/(INDEX({xs},{i}+1)-INDEX({xs},{i})))")
    WC = 'MATCH(«wmax»,{"0.2 mm","0.3 mm","0.4 mm"},0)'
    r += 1; sym(r, "σA"); desc(r, '=IF(«LEG»,IF(AND(«type»="Slab",«D»<=300),"580 − 114 ln db","760 − 173 ln db"),IF(«type»="Slab","Tbl 9.5.2.1(A), Ds "&IF(«D»<=300,"≤","> ")&"300","Tbl 8.6.2.2(A)"))')
    for col, s_ in (("E", "s"), ("F", "h")):
        x = f"«dbx_{s_}»"
        leg = f'IF(AND(«type»="Slab",«D»<=300),580-114*LN({x}),760-173*LN({x}))'
        slab = interp("CC_SA_X", "CC_SA", f"2*({WC}-1)+1+(«D»>300)", x)
        beam = interp("CC_BA_X", "CC_BA", WC, x)
        CA(f"{col}{r}", f'=IF({x}=0,"",IF(«LEG»,{leg},IF(«type»="Slab",{slab},{beam})))', "0", name=f"sA_{s_}")
    U(f"G{r}", "MPa")
    r += 1; sym(r, "σB"); desc(r, '=IF(«LEG»,"400 − 0.8 s",IF(«type»="Slab","Tbl 9.5.2.1(B)","Tbl 8.6.2.2(B)"))')
    for col, s_ in (("E", "s"), ("F", "h")):
        x = f"«sp_{s_}»"
        CA(f"{col}{r}", f'=IF({x}=0,"",IF(«LEG»,400-0.8*{x},{interp("CC_B_X", "CC_B", WC, x)}))', "0", name=f"sB_{s_}")
    U(f"G{r}", "MPa")
    r += 1; sym(r, "σs.max"); desc(r, "adopted: input, or max(σA, σB)")
    for col, s_ in (("E", "s"), ("F", "h")):
        ov = "«ovs»" if s_ == "s" else "«ovh»"
        CA(f"{col}{r}", f'=IF(N({ov})>0,{ov},MAX(N(«sA_{s_}»),N(«sB_{s_}»)))', "0", name=f"smax_{s_}")
    U(f"G{r}", "MPa")
    r += 1; sym(r, "0.8 fsy"); desc(r, "limit for Ms1*")
    CA(f"E{r}", "=0.8*«fsy_s»", "0", name="s08_s"); CA(f"F{r}", "=0.8*«fsy_h»", "0", name="s08_h"); U(f"G{r}", "MPa")
    r += 1; sym(r, "s.max"); desc(r, '=IF(«type»="Slab","bar spacing ≤ min(2D, 300)","bar spacing ≤ 300")')
    CA(f"E{r}", '=IF(«type»="Slab",MIN(2*«D»,300),300)', "0", name="spmax"); CA(f"F{r}", "=«spmax»", "0"); U(f"G{r}", "mm")
    CC_END = r

    # Shrinkage and temperature (slabs)
    r = CC_END + 2; heading(ws, r, "Shrinkage and Temperature Reinforcement (slabs)")
    r += 1; note(ws, r, "As previous OAR slab sheet: ρ = 1.75 + restraint × (ρ.control − 1.75) ×10⁻³, ρ.control 1.75 minor / 3.5 moderate / 6.0 strong. Main mats only (extras may not run full length).")
    r += 1; sym(r, "ρ.ctrl"); desc(r, '="degree of control: "&LOWER(«control»)')
    CA(f"E{r}", '=IF(«control»="Strong",6,IF(«control»="Moderate",3.5,1.75))', "0.00", name="rctrl"); U(f"F{r}", "×10⁻³")
    r += 1; sym(r, "ρ"); desc(r, "1.75 + r (ρ.ctrl − 1.75)"); CA(f"E{r}", "=1.75+«restr»*(«rctrl»-1.75)", "0.00", name="rho"); U(f"F{r}", "×10⁻³")
    r += 1; sym(r, "As.min"); desc(r, "ρ b D"); CA(f"E{r}", "=«rho»/1000*«b»*«D»", "0", name="Ashr"); U(f"F{r}", "mm²")
    r += 1; sym(r, "As"); desc(r, "top + bottom main mats"); CA(f"E{r}", "=«Astm»+«Asbm»", "0", name="Asmats"); U(f"F{r}", "mm²")
    LAST = r + 1
    note(ws, LAST, "Shear, torsion, deflection and detailing (anchorage, laps, bar spacing minimum) are not covered by this calc.")

    # ------------------------------------------------------------------ summary and checks (top of page 1)
    heading(ws, 8, "Design Summary")
    lab("B9", "Adopted section", sz=7, color=GREY); lab("F9", "Max utilisation", sz=7, color=GREY); lab("I9", "Status", sz=7, color=GREY)
    ws.merge_cells("B10:D10"); ws.merge_cells("F10:G10"); ws.merge_cells("I10:K10")
    pend.append(("B10", '=IF(«type»="Slab",«D»&" slab",«b»&" × "&«D»&" beam")&"  ·  f\'c "&«fc»'))
    ws["B10"].font = font(12, True, color=GREEN); ws["B10"].alignment = LEFT
    CHK0, CHK1 = 15, 32
    c = ws["F10"]; c.value = f'=IF(COUNT(I{CHK0}:I{CHK1})=0,"–",MAX(I{CHK0}:I{CHK1}))'; c.font = font(12, True, color=GREEN); c.number_format = "0%"; c.alignment = LEFT
    c = ws["I10"]; c.value = (f'=IF($V$7,"UPDATE ADD-IN",IF(COUNTIF(E{U0}:F{U1},"#OAR*")+COUNTIF(E{C1}:H{C2},"#OAR*")>0,"INPUT ERROR",'
                              f'IF(COUNT(I{CHK0}:I{CHK1})=0,"PENDING",IF(OR(MAX(I{CHK0}:I{CHK1})>1,COUNTIF(J{CHK0}:J{CHK1},"NG")>0),"FAIL","PASS"))))')
    c.font = font(10, True, color=GREY); c.alignment = CEN
    for x in ws["I10:K10"][0]: x.border = BOX
    passfail_cf(ws, "I10:K10", "$I$10")
    ws.conditional_formatting.add("I10:K10", FormulaRule(formula=['$I$10="UPDATE ADD-IN"'], fill=PatternFill(bgColor="C55A11"), font=Font(color="FFFFFF", bold=True)))
    ws.conditional_formatting.add("B7", FormulaRule(formula=['OR($V$7,$V$4="")'], font=Font(color=RED, bold=True, italic=False)))
    def face(f): return (f'IF(«type»="Slab",«{f}mb»&"@"&«{f}mn»,«{f}mn»&«{f}mb»)&IF(N(«{f}xn»)>0," + "&«{f}xn»&«{f}xb»,"")')
    cen_sum = ('IF(«layout»="Single mat – central","Central ",IF(«layout»="Single mat – at top cover","Top mat ","Bottom mat "))&IF(«type»="Slab",«tmb»&"@"&«tmn»,«tmn»&«tmb»)&IF(N(«Astx»)>0," + "&«txn»&«txb»&" above","")&IF(N(«Asbx»)>0," + "&«bxn»&«bxb»&" below","")')
    pend.append(("B11", f'=IF(«CEN»,{cen_sum},"Top "&IF(N(«Astm»)+N(«Astx»)=0,"none",{face("t")})&"  ·  bottom "&IF(N(«Asbm»)+N(«Asbx»)=0,"none",{face("b")}))&"  ·  "&«model»&IF(«LO»,"  ·  Class L: L steel only","")'))
    ws["B11"].font = font(7, color=GREY)
    c = ws["F11"]
    c.value = (f'=IFERROR(IF(LEFT(E{U0},4)="#OAR",E{U0},IF(LEFT(F{U0},4)="#OAR",F{U0},IF(LEFT(E{C1},4)="#OAR",E{C1},IF(LEFT(F{C1},4)="#OAR",F{C1},'
               f'IF(COUNT(I{CHK0}:I{CHK1})=0,"","Governs: "&INDEX(D{CHK0}:D{CHK1},MATCH(MAX(I{CHK0}:I{CHK1}),I{CHK0}:I{CHK1},0))))))),"")')
    c.font = font(7, color=GREY)
    table_header(ws, 13, [("B13:C13", "Check"), ("D13", "Basis"), ("E13", "Actual"), ("F13", "Limit"), ("G13", "Unit"),
                          ("H13", "Ratio"), ("I13", "Utilisation"), ("J13:K13", "Status")])
    heading(ws, 14, "Strength", sz=8)
    def chk(rr, s_, basis, actual, limit, unit_, util, status=None, nf="0.0", ratio=None):
        sym(rr, s_); desc(rr, basis)
        CA(f"E{rr}", actual, nf); CA(f"F{rr}", limit, nf); U(f"G{rr}", unit_); ws[f"G{rr}"].alignment = CEN
        CA(f"H{rr}", ratio, "0.00"); CA(f"I{rr}", util, "0%")
        ws.merge_cells(f"J{rr}:K{rr}"); pend.append((f"J{rr}", status if status else f'=IF(I{rr}="","–",IF(I{rr}<=1,"OK","NG"))'))
        ws[f"J{rr}"].font = font(9, True, color=GREEN); ws[f"J{rr}"].alignment = CEN
    def ut(rr, blank): return f'=IF({blank},"",E{rr}/F{rr})'
    chk(15, "M*+ ≤ φMuo", "sagging, bottom tension (Cl 8.1)", "=IF(N(«sMu»)=0,\"\",«sMu»)", "=«phiM_s»", "kNm", ut(15, "N(«sMu»)=0"))
    chk(16, "M*− ≤ φMuo", "hogging, top tension (Cl 8.1)", "=IF(N(«hMu»)=0,\"\",«hMu»)", "=«phiM_h»", "kNm", ut(16, "N(«hMu»)=0"))
    for rr, s in ((17, "s"), (18, "h")):
        face = "bottom" if s == "s" else "top"
        chk(rr, f"Muo{'+' if s == 's' else '−'} ≥ (Muo)min", f"1.2Mcr, or Ast ≥ Ast.min ({face})", "=«Mmin»", f"=«Muo_{s}»", "kNm",
            f'=IF(«Ast_{s}»=0,"",MIN(«Mmin»/«Muo_{s}»,IF(N(«Amin_{s}»)>0,«Amin_{s}»/«Ast_{s}»,9)))',
            status=f'=IF(I{rr}="","–",IF(«Mmin»/«Muo_{s}»<=1,"OK",IF(I{rr}<=1,"OK*","NG")))', ratio=f'=IF(OR(«Ast_{s}»=0,N(«Amin_{s}»)=0),"",«Ast_{s}»/«Amin_{s}»)')
    chk(19, "εsL ≤ 1.5 %", '=IF(«LO»,"Class L strain at ULS (L steel only)","Class L strain at ULS (info)")',
        '=IF(COUNTIF($I$69:$I$72,"L")=0,"",MAX(«epsL_s»,«epsL_h»))', "=0.015", "", '=IF(OR(E19="",NOT(«LO»)),"",E19/F19)',
        status='=IF(E19="","–",IF(NOT(«LO»),"info",IF(I19<=1,"OK","NG")))', nf="0.0000")
    heading(ws, 20, "Serviceability – crack control", sz=8)
    chk(21, "σscr+ (Ms*)", "sagging, bottom bars ≤ σs.max", "=IF(N(«sMs»)=0,\"\",«sig_s»)", "=«smax_s»", "MPa", ut(21, "N(«sMs»)=0"), nf="0")
    chk(22, "σscr− (Ms*)", "hogging, top bars ≤ σs.max", "=IF(N(«hMs»)=0,\"\",«sig_h»)", "=«smax_h»", "MPa", ut(22, "N(«hMs»)=0"), nf="0")
    chk(23, "σscr.1+ (Ms1*)", "sagging ≤ 0.8 fsy", "=IF(N(«sMs1»)=0,\"\",«sig1_s»)", "=«s08_s»", "MPa", ut(23, "N(«sMs1»)=0"), nf="0")
    chk(24, "σscr.1− (Ms1*)", "hogging ≤ 0.8 fsy", "=IF(N(«hMs1»)=0,\"\",«sig1_h»)", "=«s08_h»", "MPa", ut(24, "N(«hMs1»)=0"), nf="0")
    chk(25, "s (sag)", '=IF(«type»="Slab","bar spacing ≤ min(2D, 300)","bar spacing ≤ 300 (layer 1)")', '=IF(«sp_s»=0,"",«sp_s»)', "=«spmax»", "mm",
        ut(25, "«sp_s»=0"), nf="0")
    chk(26, "s (hog)", '=IF(«type»="Slab","bar spacing ≤ min(2D, 300)","bar spacing ≤ 300 (layer 1)")', '=IF(«sp_h»=0,"",«sp_h»)', "=«spmax»", "mm",
        ut(26, "«sp_h»=0"), nf="0")
    chk(27, "As shrinkage", '=IF(«type»="Slab","main mats ≥ ρ b D (shrinkage & temp.)","slabs only")',
        '=IF(«type»="Slab",«Ashr»,"")', '=IF(«type»="Slab",«Asmats»,"")', "mm²", '=IF(«type»<>"Slab","",E27/F27)', nf="0")
    heading(ws, 28, "Information (not in max utilisation)", sz=8)
    for rr, s in ((29, "s"), (30, "h")):
        chk(rr, f"kuo{'+' if s == 's' else '−'}", f'="φ = "&TEXT(«phi_{s}»,"0.00")&IF(«LO»,"  (kL = "&TEXT(«kL_{s}»,"0.00")&")","  (Class "&«cls_{s}»&")")', f"=«kuo_{s}»", None, "", None, status="info", nf="0.000")
    for rr, s in ((31, "s"), (32, "h")):
        chk(rr, f"φMuo'{'+' if s == 's' else '−'}", '=IF(«model»="Rectangular","curvilinear block","rectangular block")&" (comparison)"',
            f"=«phiM2_{s}»", f"=«phiM_{s}»", "kNm", None, status="info", ratio=f"=IF(N(«phiM_{s}»)=0,\"\",«phiM2_{s}»/«phiM_{s}»)")
    for rr in range(29, 33): ws[f"J{rr}"].font = font(8, i=True, color=GREY)
    status_cf(ws, f"J{CHK0}:K{CHK1}"); util_cf(ws, f"I{CHK0}:I{CHK1}", f"I{CHK0}")

    # ------------------------------------------------------------------ resolve, arrays, validation, grey-outs
    for ref, f in pend: ws[ref].value = Fx(f)
    for ref in arrays:
        top = ref.split(":")[0]; ws[top].value = ArrayFormula(ref, ws[top].value)
    def dvl(formula, ref):
        v = DataValidation(type="list", formula1=formula, allow_blank=True); ws.add_data_validation(v); v.add(ref)
    dvl('"Slab,Beam"', "E34"); dvl('"Rectangular,Curvilinear,Parabolic"', "E38"); dvl("=BarList", "E45:F45"); dvl("=BarList", "E48:F48")
    dvl('"Layer 1,Layer 2"', "E49:F49"); dvl("=BarList", "E51"); dvl('"Beams or walls,Columns"', "E52"); dvl('"Minor,Moderate,Strong"', "E63"); dvl('"Whole section 0.65,L steel only"', "E53"); dvl('"0.2 mm,0.3 mm,0.4 mm,Legacy"', "E65"); dvl('"Top & bottom,Single mat – central,Single mat – at top cover,Single mat – at bottom cover"', "E41")
    rules = [("E51", '«type»<>"Beam"', "B51"),
             ("E52:F52", '«type»<>"Slab"', "B52"),
             ("E63:F64", '«type»<>"Slab"', "B63:B64"),
             ("E50", '«txp»<>"Layer 2"', "E50"), ("F50", '«bxp»<>"Layer 2"', "F50"),
             ("E47:E50", "N(«txn»)=0", "E47"), ("F47:F50", "N(«bxn»)=0", "F47"),
             ("E44", 'AND(«CEN»,«layout»<>"Single mat – at top cover")', "E44"),
             ("F44", 'AND(«CEN»,«layout»<>"Single mat – at bottom cover")', "F44"), ("F45:F46", "«CEN»", "F45"), ("E49:F50", "«CEN»", "B49:B50")]
    for ref, cond, labref in rules:
        f = Fx(cond)
        ws.conditional_formatting.add(ref, FormulaRule(formula=[f], fill=PatternFill(bgColor=NA_FILL), font=Font(color=NA_FONT)))
        ws.conditional_formatting.add(labref, FormulaRule(formula=[f], font=Font(color=NA_FONT)))
    for rr in range(15, LAST + 1):
        if ws.row_dimensions[rr].height is None: ws.row_dimensions[rr].height = 14
    ws.print_area = f"A1:K{LAST}"
    ws.row_breaks.append(Break(id=65))
    ws.row_breaks.append(Break(id=ULT_END + 1))
    return ws, N


# ---------------- examples ----------------
# S1: Brent's previous slab strip sheet example (25 Lutana – X strips): 175 slab, SL92 T&B, 3 N16 extra top.
s1 = dict(loc="25 Lutana – X strips", type="Slab", b=1000, D=175, fc=32, model="Rectangular",
          top=dict(cov=40, mb="L9", mn=200, xn=3, xb="N16", xp="Layer 1", gap=25),
          bot=dict(cov=40, mb="L9", mn=200, xn=0, xb="N16", xp="Layer 1", gap=25),
          sag=dict(Mu=8.7, Ms=6.7, Ms1=6.0), hog=dict(Mu=37.3, Ms=28.6, Ms1=25.6), control="Moderate", restraint=0.75)
# B1: Helensvale wide band beam (previous beam sheet), top steel increased with a second layer.
b1 = dict(loc="Helensvale – wide beam 3", type="Beam", b=1200, D=600, fc=40, model="Rectangular", lig="N12",
          top=dict(cov=50, mb="N24", mn=7, xn=4, xb="N24", xp="Layer 2", gap=25),
          bot=dict(cov=50, mb="N24", mn=6, xn=0, xb="N20", xp="Layer 1", gap=25),
          sag=dict(Mu=180, Ms=100, Ms1=170), hog=dict(Mu=943.8, Ms=726, Ms1=726))
# S2: pool floor / slab on ground with one central mat (SL92 = L9 @ 200) and N12 extras under the mat
s2 = dict(loc="Pool floor – single central mat", type="Slab", b=1000, D=150, fc=32, model="Rectangular", layout="Single mat – central",
          top=dict(cov=50, mb="L9", mn=200, xn=0, xb="N12", xp="Layer 1", gap=25),
          bot=dict(cov=50, mb="None", mn=200, xn=5, xb="N12", xp="Layer 1", gap=25),
          sag=dict(Mu=6.5, Ms=4.0, Ms1=4.5), hog=dict(Mu=3.0, Ms=1.8, Ms1=2.0), control="Minor", restraint=0.75)
EXAMPLES = (("S1", s1), ("B1", b1), ("S2", s2))
if __name__ == "__main__":
    for nm_, dd in EXAMPLES: build_tab(nm_, dd)
    wb._sheets = [wb[n] for n in ("S1", "B1", "S2", "Bars", "Crack Control")]
    wb.active = 0
    from openpyxl.workbook.properties import CalcProperties
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    wb.save(OUT)
    print("saved", OUT)
