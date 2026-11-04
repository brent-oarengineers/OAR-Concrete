"""Builds template/RC Flexural Design - Base.xlsx (OAR house calc style).

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
OUT = sys.argv[2] if len(sys.argv) > 2 else os.path.join(HERE, "RC Flexural Design - Base.xlsx")

# ---------------- house-style helpers (oar-calc-sheet skill) ----------------
F = "Arial"; GREEN = "465A52"; GREY = "595959"; TEXT = "262626"; IN_FILL = "F3EEE9"; CALC_FILL = "E6ECE9"
HDR_FILL = "DCD5CF"; RULE = "B8ADA5"; HAIR = "8C8C8C"; BAR = "9DB3A9"; RED = "C00000"; NA_FILL = "F2F2F2"; NA_FONT = "BFBFBF"
AMBER = "B7791F"; FAIL_FILL = "F6E1E1"
def font(sz=9, b=False, i=False, color=TEXT): return Font(name=F, size=sz, bold=b, italic=i, color=color)
hair = Side(style="hair", color=HAIR); BOX = Border(left=hair, right=hair, top=hair, bottom=hair)
UNDER = Border(bottom=Side(style="thin", color=RULE)); fill = lambda c: PatternFill("solid", fgColor=c)
CEN = Alignment(horizontal="center", vertical="center"); LEFT = Alignment(horizontal="left", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")
COLS = {"A": 1.5, "B": 8, "C": 9, "D": 26, "E": 10, "F": 10, "G": 11, "H": 10, "I": 10, "J": 11, "K": 7}


def setup_calc_tab(ws, logo_path, title_formula, ref_label, ref_default, note_text):
    ws.sheet_view.showGridLines = False; ws.sheet_properties.tabColor = GREEN
    for k, v in COLS.items(): ws.column_dimensions[k].width = v
    for r in range(1, 5): ws.row_dimensions[r].height = 17
    ws.row_dimensions[5].height = 6; ws.row_dimensions[6].height = 22; ws.row_dimensions[7].height = 14
    img = XLImage(logo_path); img.width, img.height = 282, 40; ws.add_image(img, "B1")
    def lab(ref, t): c = ws[ref]; c.value = t; c.font = font(7, color=GREY); c.alignment = RIGHT
    def box(ref, v=None, al=CEN, nf=None, merge=None):
        if merge: ws.merge_cells(merge)
        c = ws[ref]; c.value = v; c.font = font(9, color="000000"); c.alignment = al
        for row in (ws[merge] if merge else ((c,),)):
            for x in row: x.border = BOX; x.fill = fill(IN_FILL)
        if nf: c.number_format = nf
    lab("F1", "JOB No."); box("G1", "26-000"); lab("H1", "DATE"); box("I1", None, nf="dd-mmm-yy", merge="I1:K1")
    lab("F2", "PROJECT"); box("G2", "Project name / address", LEFT, merge="G2:K2")
    lab("F3", ref_label); box("G3", ref_default); lab("H3", "LOCATION"); box("I3", None, LEFT, merge="I3:K3")
    lab("F4", "BY"); box("G4", "BO"); lab("H4", "CHECKED"); box("I4", None, merge="I4:K4")
    c = ws["B4"]; c.value = "OAR Consulting Engineers Pty Ltd  ·  22 Martin Street, Nerang QLD 4211  ·  (07) 5527 2288"; c.font = font(7, color=GREY)
    ws.merge_cells("B6:K6"); c = ws["B6"]; c.value = title_formula; c.font = font(11, True, color="FFFFFF"); c.alignment = LEFT
    for x in ws["B6:K6"][0]: x.fill = fill(GREEN)
    c = ws["B7"]; c.value = note_text; c.font = font(7, i=True, color=GREY); lab("H7", "Legend:")
    c = ws["I7"]; c.value = "Input"; c.font = font(7, color="000000"); c.fill = fill(IN_FILL); c.border = BOX; c.alignment = CEN
    c = ws["J7"]; c.value = "Calculated"; c.font = font(7); c.fill = fill(CALC_FILL); c.border = BOX; c.alignment = CEN
    ws.page_setup.paperSize = 9; ws.page_setup.orientation = "portrait"; ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.page_margins.left = ws.page_margins.right = ws.page_margins.top = 0.5; ws.page_margins.bottom = 0.6
    ws.oddFooter.left.text = "OAR Consulting Engineers Pty Ltd"; ws.oddFooter.center.text = "&A"; ws.oddFooter.right.text = "Page &P of &N"
    for p in (ws.oddFooter.left, ws.oddFooter.center, ws.oddFooter.right): p.font = "Arial"; p.size = 7; p.color = GREY


def location_frame(ws, element):
    for col, w in zip("LMNOPQRST", (4, 8, 7, 9, 8, 9, 26, 6.5, 6.5)): ws.column_dimensions[col].width = w
    c = ws["M6"]; c.value = f"{element} Location  (not printed)"; c.font = font(10, True, color=GREEN); c.alignment = Alignment(vertical="bottom")
    for x in ws["M6:T6"][0]: x.border = UNDER
    dash = Side(style="dashed", color=RULE)
    for row in ws["M8:T40"]:
        for x in row:
            x.fill = fill("FAF8F6")
            x.border = Border(left=dash if x.column == 13 else None, right=dash if x.column == 20 else None,
                              top=dash if x.row == 8 else None, bottom=dash if x.row == 40 else None)
    c = ws["M8"]; c.value = f"Paste {element.lower()} location screenshot / plan extract here\n(right-click > Paste as Picture, then resize to fit this frame)"
    c.font = font(9, i=True, color="A6A6A6"); c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.merge_cells("M8:T40")


def heading(ws, r, text, sz=10, last="K"):
    ws.row_dimensions[r].height = 20 if sz == 10 else 14
    c = ws[f"B{r}"]; c.value = text; c.font = font(sz, True, color=GREEN)
    for x in ws[f"B{r}:{last}{r}"][0]: x.border = UNDER
def note(ws, r, text): c = ws[f"B{r}"]; c.value = text; c.font = font(7, i=True, color=GREY)
def label(ws, ref, text, sz=9, al=LEFT, color=TEXT, b=False, i=False):
    c = ws[ref]; c.value = text; c.font = font(sz, b, i, color); c.alignment = al; return c
def unit(ws, ref, text, al=LEFT): return label(ws, ref, text, 9, al, GREY)
def _boxrange(ws, ref, fc):
    if ":" in ref:
        ws.merge_cells(ref)
        for row in ws[ref]:
            for x in row: x.fill = fill(fc); x.border = BOX
        return ws[ref.split(":")[0]]
    return ws[ref]
def inp(ws, ref, v=None, nf=None, al=CEN):
    c = _boxrange(ws, ref, IN_FILL); c.value = v; c.font = font(9, color="000000"); c.fill = fill(IN_FILL); c.border = BOX; c.alignment = al
    if nf: c.number_format = nf
    return c
def calc(ws, ref, f, nf=None, al=CEN):
    c = _boxrange(ws, ref, CALC_FILL); c.value = f; c.font = font(9); c.fill = fill(CALC_FILL); c.border = BOX; c.alignment = al
    if nf: c.number_format = nf
    return c
def table_header(ws, r, items, box=False):
    for ref, t in items:
        if ":" in ref: ws.merge_cells(ref)
        for row in (ws[ref] if ":" in ref else ((ws[ref],),)):
            for x in row:
                x.fill = fill(HDR_FILL)
                if box: x.border = BOX
        c = ws[ref.split(":")[0]]; c.value = t; c.font = font(8, True); c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[r].height = 16
def status_cf(ws, rng):
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"NG"'], fill=PatternFill(bgColor=RED), font=Font(color="FFFFFF", bold=True)))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"OK*"'], font=Font(color=AMBER, bold=True)))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"OK"'], font=Font(color=GREEN, bold=True)))
def util_cf(ws, rng, first):
    ws.conditional_formatting.add(rng, DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color=BAR, showValue=True))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({first}),{first}>1)"], font=Font(color=RED, bold=True)))
def passfail_cf(ws, rng, anchor):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{anchor}="PASS"'], fill=PatternFill(bgColor=CALC_FILL), font=Font(color=GREEN, bold=True)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{anchor}="FAIL"'], fill=PatternFill(bgColor=FAIL_FILL), font=Font(color=RED, bold=True)))
def setup_data_tab(ws, title, note_text, widths):
    ws.sheet_view.showGridLines = False; ws.sheet_properties.tabColor = RULE; ws.column_dimensions["A"].width = 1.5
    for k, v in widths.items(): ws.column_dimensions[k].width = v
    c = ws["B2"]; c.value = title; c.font = font(11, True, color=GREEN)
    c = ws["B3"]; c.value = note_text; c.font = font(7, i=True, color=GREY)
    ws.page_setup.paperSize = 9; ws.page_setup.orientation = "portrait"; ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.oddFooter.left.text = "OAR Consulting Engineers Pty Ltd"; ws.oddFooter.center.text = "&A"; ws.oddFooter.right.text = "Page &P of &N"
    for p_ in (ws.oddFooter.left, ws.oddFooter.center, ws.oddFooter.right): p_.font = "Arial"; p_.size = 7; p_.color = GREY


wb = Workbook()
wb.remove(wb.active)

# =========================== DATA TAB: Bars ===========================
# AS/NZS 4671 nominal areas. L = mesh wires (Class L, e.g. SL92 = L9 @ 200). R = plain round (R250N).
BARS = [
    ["N10", 10, 80, 500, "N"], ["N12", 12, 110, 500, "N"], ["N16", 16, 200, 500, "N"], ["N20", 20, 310, 500, "N"],
    ["N24", 24, 450, 500, "N"], ["N28", 28, 620, 500, "N"], ["N32", 32, 800, 500, "N"], ["N36", 36, 1020, 500, "N"],
    ["N40", 40, 1260, 500, "N"], ["R10", 10, 78, 250, "N"], ["R12", 12, 113, 250, "N"],
    ["L6", 6.0, 28.3, 500, "L"], ["L7", 6.75, 35.8, 500, "L"], ["L8", 7.6, 45.4, 500, "L"], ["L9", 8.6, 58.1, 500, "L"],
    ["L10", 9.5, 70.9, 500, "L"], ["L11", 10.65, 89.1, 500, "L"], ["L12", 11.9, 111.2, 500, "L"],
]
D_ = wb.create_sheet("Bars")
setup_data_tab(D_, "Reinforcement – bar and mesh wire properties",
               "AS/NZS 4671 nominal values. L = mesh wire (Class L; SL92 = L9 @ 200, RL1218 = L12 @ 200 main). Edit or add rows below; the calc tabs read this list.",
               {"B": 10, "C": 10, "D": 10, "E": 10, "F": 10, "G": 40})
for j, t in enumerate(["Bar", "db (mm)", "Ab (mm²)", "fsy (MPa)", "Class", "Note"]):
    c = D_.cell(4, 2 + j); c.value = t; c.font = font(8, True); c.fill = fill(HDR_FILL); c.border = BOX; c.alignment = CEN
BAR0 = 5; BAR_LAST = BAR0 + len(BARS) + 5
notes = {"R10": "plain round R250N", "L9": "SL92 mesh wire", "L12": "RL1218 main wire"}
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

# =========================== CALC TAB ===========================
def build_tab(name, d):
    ws = wb.create_sheet(name)
    N = {}  # named cells → absolute refs, used as «name» in formulas
    def nm(n, ref): N[n] = "$" + "".join(ch for ch in ref if ch.isalpha()) + "$" + "".join(ch for ch in ref if ch.isdigit())
    def Fx(s):
        if not isinstance(s, str): return s
        s = s.replace("«LO»", 'LEFT(«lmode»,1)="L"')
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
    c = ws["E7"]; c.value = '=IFERROR(OAR.CONCRETE.VERSION(),"OAR Concrete add-in not loaded")'; c.font = font(7, i=True, color=GREY)

    # ------------------------------------------------------------------ inputs (page 1, below the checks)
    r = 33; heading(ws, r, "Member and Concrete")
    lab("B34", "Member type"); IN("E34", d["type"], name="type"); lab("H34", "Slab: strip of width b, mats as bar @ spacing", sz=7, color=GREY)
    lab("B35", "Width b"); IN("E35", d["b"], "0", name="b"); U("F35", "mm"); lab("H35", "slab: 1000 = per metre width", sz=7, color=GREY)
    lab("B36", "Overall depth D"); IN("E36", d["D"], "0", name="D"); U("F36", "mm")
    lab("B37", "Concrete strength f'c"); IN("E37", d["fc"], "0", name="fc"); U("F37", "MPa")
    lab("B38", "Stress block (ULS)"); IN("E38:F38", d["model"], name="model")
    lab("B39", "Ec override"); IN("E39", d.get("ec"), "0", name="ecin"); U("F39", "MPa"); lab("H39", "blank = Table 3.1.2", sz=7, color=GREY)
    lab("B40", "Creep factor for SLS stresses, φcc"); IN("E40", d.get("phicc", 0), "0.0", name="phicc"); lab("H40", "0 = short-term n = Es / Ec", sz=7, color=GREY)

    r = 42; heading(ws, r, "Reinforcement")
    table_header(ws, 43, [("E43", "Top"), ("F43", "Bottom")])
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
    lab("H62", "blank = legacy table (calcs)", sz=7, color=GREY)
    lab("B63", "Shrinkage & temp.: degree of control"); IN("E63:F63", d.get("control", "Moderate"), name="control"); lab("H63", "slab only", sz=7, color=GREY)
    lab("B64", "Restraint factor (0 = free, 1 = fully restrained)"); IN("E64", d.get("restraint", 0.75), "0.00", name="restr")

    # ------------------------------------------------------------------ calculations (page 2)
    # Reinforcement layers table → the add-in's bars argument (As | depth | fsy)
    r = 66; heading(ws, r, "Reinforcement Layers")
    note(ws, 67, "One row per layer across b; depth to the bar centre from the top face. Slab mat As = Ab × b / s. Beam cover includes the ligature.")
    table_header(ws, 68, [("B68:C68", "Layer"), ("D68", "Description"), ("E68", "As (mm²)"), ("F68", "depth (mm)"),
                          ("G68", "fsy (MPa)"), ("H68", "db (mm)"), ("I68", "Class"), ("J68:K68", "bars in b")])
    # effective cover to the outer bar = cover (+ ligature for a beam)
    LIG = f'IF(«type»="Beam",{barv("«lig»", 2)},0)'
    def area_main(f): return f'IF(«type»="Slab",{barv(f"«{f}mb»", 3)}*«b»/«{f}mn»,«{f}mn»*{barv(f"«{f}mb»", 3)})'
    def db(key): return f'IFERROR({barv(key, 2)},0)'
    layers = [
        (69, "Top – main", "t", "m"), (70, "Top – extra", "t", "x"), (71, "Bottom – main", "b", "m"), (72, "Bottom – extra", "b", "x")]
    for rr, t, f, k in layers:
        ws.merge_cells(f"B{rr}:C{rr}"); lab(f"B{rr}", t, sz=8)
        if k == "m":
            pend.append((f"D{rr}", f'=«{f}mb»&IF(«type»="Slab"," @ "&«{f}mn»&" ("&TEXT({barv(f"«{f}mb»", 3)}*1000/«{f}mn»,"0")&" mm²/m)"," × "&«{f}mn»)'))
            area = area_main(f)
            dep_off = f'(«{f}cov»+{LIG}+{db(f"«{f}mb»")}/2)'
            nbars = f'IF(«type»="Slab",«b»/«{f}mn»,«{f}mn»)'
        else:
            pend.append((f"D{rr}", f'=IF(N(«{f}xn»)=0,"none",«{f}xn»&" × "&«{f}xb»&IF(«{f}xp»="Layer 2"," (layer 2)",""))'))
            area = f'IF(N(«{f}xn»)=0,0,«{f}xn»*{barv(f"«{f}xb»", 3)})'
            dep_off = (f'(«{f}cov»+{LIG}+IF(«{f}xp»="Layer 2",{db(f"«{f}mb»")}+«{f}gap»+{db(f"«{f}xb»")}/2,{db(f"«{f}xb»")}/2))')
            nbars = f'N(«{f}xn»)'
        ws[f"D{rr}"].font = font(8, color=GREY)
        CA(f"E{rr}", f"={area}", "0", name=f"As{f}{k}")
        CA(f"F{rr}", f"={dep_off}" if f == "t" else f"=«D»-{dep_off}", "0.0", name=f"dep{f}{k}")
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
    CA(f"E{r}", '=IF(OR(«clbm»="L",«clbx»="L"),"L","N")', name="cls_s"); CA(f"F{r}", '=IF(OR(«cltm»="L",«cltx»="L"),"L","N")', name="cls_h")
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

    # Minimum strength
    r = ULT_END + 2; heading(ws, r, "Minimum Strength (Cl 8.1.6.1 / 9.1.1)")
    r += 1; note(ws, r, "(Muo)min = 1.2 Z f'ct.f (no prestress, N = 0). Deemed to comply if Ast ≥ αb (D/d)² f'ct.f / fsy · b d.")
    table_header(ws, r + 1, [("E%d" % (r + 1), "Sagging +"), ("F%d" % (r + 1), "Hogging −")])
    r += 2; sym(r, "Z"); desc(r, "b D² / 6, uncracked gross"); CA(f"E{r}", "=«b»*«D»^2/6/1E6", "0.00", name="Z"); CA(f"F{r}", "=«Z»", "0.00"); U(f"G{r}", "×10⁶ mm³")
    r += 1; sym(r, "(Muo)min"); desc(r, "1.2 Z f'ct.f"); CA(f"E{r}", "=1.2*«Z»*«fctf»", "0.0", name="Mmin"); CA(f"F{r}", "=«Mmin»", "0.0"); U(f"G{r}", "kNm")
    r += 1; sym(r, "Ast"); desc(r, "tension face, total"); CA(f"E{r}", "=«Ast_bot»", "0", name="Ast_s"); CA(f"F{r}", "=«Ast_top»", "0", name="Ast_h"); U(f"G{r}", "mm²")
    r += 1; sym(r, "d"); desc(r, "tension face, from compr. face")
    CA(f"E{r}", "=IF(«Ast_s»=0,0,(«Asbm»*«depbm»+«Asbx»*«depbx»)/«Ast_s»)", "0.0", name="d_s")
    CA(f"F{r}", "=IF(«Ast_h»=0,0,«D»-(«Astm»*«deptm»+«Astx»*«deptx»)/«Ast_h»)", "0.0", name="d_h"); U(f"G{r}", "mm")
    r += 1; sym(r, "fsy"); desc(r, "tension face (lowest)")
    CA(f"E{r}", "=MIN(«fsybm»,IF(«Asbx»>0,«fsybx»,«fsybm»))", "0", name="fsy_s"); CA(f"F{r}", "=MIN(«fsytm»,IF(«Astx»>0,«fsytx»,«fsytm»))", "0", name="fsy_h"); U(f"G{r}", "MPa")
    r += 1; sym(r, "αb"); desc(r, '=IF(«type»="Beam","rectangular beam 0.20","slab on "&LOWER(«support»))')
    CA(f"E{r}", '=IF(«type»="Beam",0.2,IF(«support»="Columns",0.24,0.19))', "0.00", name="alb"); CA(f"F{r}", "=«alb»", "0.00")
    r += 1; sym(r, "Ast.min"); desc(r, "αb (D/d)² f'ct.f / fsy · b d")
    CA(f"E{r}", "=IF(«d_s»=0,0,«alb»*(«D»/«d_s»)^2*«fctf»/«fsy_s»*«b»*«d_s»)", "0", name="Amin_s")
    CA(f"F{r}", "=IF(«d_h»=0,0,«alb»*(«D»/«d_h»)^2*«fctf»/«fsy_h»*«b»*«d_h»)", "0", name="Amin_h"); U(f"G{r}", "mm²")
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
    r += 1; note(ws, r, "Legacy OAR limits (previous slab strip sheet): fits to AS 3600-2009 Tbl 8.6.1 / 9.4.1 – A by bar diameter, B by bar spacing, the larger applies.")
    r += 1; note(ws, r, "AS 3600-2018 Amdt 2 tabulates σs by crack width (Tbl 8.6.2.2 / 9.5.2.1): enter σs.max in the inputs to use those values instead.")
    table_header(ws, r + 1, [(f"E{r+1}", "Sagging +"), (f"F{r+1}", "Hogging −")])
    r += 2; sym(r, "db"); desc(r, "largest bar, tension face")
    CA(f"E{r}", "=MAX(N(«dbbm»),N(«dbbx»))", "0.0", name="dbx_s"); CA(f"F{r}", "=MAX(N(«dbtm»),N(«dbtx»))", "0.0", name="dbx_h"); U(f"G{r}", "mm")
    r += 1; sym(r, "s"); desc(r, '=IF(«type»="Slab","b / (bars in layer 1 of the face)","clear width / (n − 1), layer 1")')
    def spacing(f):
        n1 = f'(«n{f}m»+IF(«{f}xp»="Layer 2",0,«n{f}x»))'
        return (f'=IF(«type»="Slab",«b»/{n1},IF({n1}<=1,0,(«b»-2*(«{f}cov»+{barv("«lig»", 2)})-«db{f}m»)/({n1}-1)))')
    CA(f"E{r}", spacing("b"), "0", name="sp_s"); CA(f"F{r}", spacing("t"), "0", name="sp_h"); U(f"G{r}", "mm")
    r += 1; sym(r, "σA"); desc(r, '=IF(AND(«type»="Slab",«D»<=300),"slab D ≤ 300: 580 − 114 ln db","760 − 173 ln db")')
    for col, s in (("E", "s"), ("F", "h")):
        CA(f"{col}{r}", f'=IF(«dbx_{s}»=0,"",IF(AND(«type»="Slab",«D»<=300),580-114*LN(«dbx_{s}»),760-173*LN(«dbx_{s}»)))', "0", name=f"sA_{s}")
    U(f"G{r}", "MPa")
    r += 1; sym(r, "σB"); desc(r, "400 − 0.8 s")
    for col, s in (("E", "s"), ("F", "h")): CA(f"{col}{r}", f'=IF(«sp_{s}»=0,"",400-0.8*«sp_{s}»)', "0", name=f"sB_{s}")
    U(f"G{r}", "MPa")
    r += 1; sym(r, "σs.max"); desc(r, "adopted: input, or max(σA, σB)")
    for col, s in (("E", "s"), ("F", "h")):
        ov = "«ovs»" if s == "s" else "«ovh»"
        CA(f"{col}{r}", f'=IF(N({ov})>0,{ov},MAX(N(«sA_{s}»),N(«sB_{s}»)))', "0", name=f"smax_{s}")
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
    c = ws["I10"]; c.value = (f'=IF(COUNTIF(E{U0}:F{U1},"#OAR*")+COUNTIF(E{C1}:H{C2},"#OAR*")>0,"INPUT ERROR",'
                              f'IF(COUNT(I{CHK0}:I{CHK1})=0,"PENDING",IF(OR(MAX(I{CHK0}:I{CHK1})>1,COUNTIF(J{CHK0}:J{CHK1},"NG")>0),"FAIL","PASS")))')
    c.font = font(10, True, color=GREY); c.alignment = CEN
    for x in ws["I10:K10"][0]: x.border = BOX
    passfail_cf(ws, "I10:K10", "$I$10")
    def face(f): return (f'IF(«type»="Slab",«{f}mb»&"@"&«{f}mn»,«{f}mn»&«{f}mb»)&IF(N(«{f}xn»)>0," + "&«{f}xn»&«{f}xb»,"")')
    pend.append(("B11", f'="Top "&{face("t")}&"  ·  bottom "&{face("b")}&"  ·  "&«model»&IF(«LO»,"  ·  Class L: L steel only","")'))
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
            f'=IF(«Ast_{s}»=0,"",MIN(«Mmin»/«Muo_{s}»,IF(«Ast_{s}»>0,«Amin_{s}»/«Ast_{s}»,9)))',
            status=f'=IF(I{rr}="","–",IF(«Mmin»/«Muo_{s}»<=1,"OK",IF(I{rr}<=1,"OK*","NG")))', ratio=f'=IF(«Ast_{s}»=0,"",«Ast_{s}»/«Amin_{s}»)')
    chk(19, "εsL ≤ 1.5 %", '=IF(«LO»,"Class L strain at ULS (L steel only)","Class L strain at ULS (info)")',
        '=IF(COUNTIF($I$69:$I$72,"L")=0,"",MAX(«epsL_s»,«epsL_h»))', "=0.015", "", '=IF(OR(E19="",NOT(«LO»)),"",E19/F19)',
        status='=IF(E19="","–",IF(NOT(«LO»),"info",IF(I19<=1,"OK","NG")))', nf="0.0000")
    heading(ws, 20, "Serviceability – crack control", sz=8)
    chk(21, "σscr+ (Ms*)", "sagging, bottom bars ≤ σs.max", "=IF(N(«sMs»)=0,\"\",«sig_s»)", "=«smax_s»", "MPa", ut(21, "N(«sMs»)=0"), nf="0")
    chk(22, "σscr− (Ms*)", "hogging, top bars ≤ σs.max", "=IF(N(«hMs»)=0,\"\",«sig_h»)", "=«smax_h»", "MPa", ut(22, "N(«hMs»)=0"), nf="0")
    chk(23, "σscr.1+ (Ms1*)", "sagging ≤ 0.8 fsy", "=IF(N(«sMs1»)=0,\"\",«sig1_s»)", "=«s08_s»", "MPa", ut(23, "N(«sMs1»)=0"), nf="0")
    chk(24, "σscr.1− (Ms1*)", "hogging ≤ 0.8 fsy", "=IF(N(«hMs1»)=0,\"\",«sig1_h»)", "=«s08_h»", "MPa", ut(24, "N(«hMs1»)=0"), nf="0")
    chk(25, "s bottom", '=IF(«type»="Slab","bar spacing ≤ min(2D, 300)","bar spacing ≤ 300 (layer 1)")', '=IF(«sp_s»=0,"",«sp_s»)', "=«spmax»", "mm",
        ut(25, "«sp_s»=0"), nf="0")
    chk(26, "s top", '=IF(«type»="Slab","bar spacing ≤ min(2D, 300)","bar spacing ≤ 300 (layer 1)")', '=IF(«sp_h»=0,"",«sp_h»)', "=«spmax»", "mm",
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
    dvl('"Layer 1,Layer 2"', "E49:F49"); dvl("=BarList", "E51"); dvl('"Beams or walls,Columns"', "E52"); dvl('"Minor,Moderate,Strong"', "E63"); dvl('"Whole section 0.65,L steel only"', "E53")
    rules = [("E51", '«type»<>"Beam"', "B51"),
             ("E52:F52", '«type»<>"Slab"', "B52"),
             ("E63:F64", '«type»<>"Slab"', "B63:B64"),
             ("E50", '«txp»<>"Layer 2"', "E50"), ("F50", '«bxp»<>"Layer 2"', "F50"),
             ("E47:E50", "N(«txn»)=0", "E47"), ("F47:F50", "N(«bxn»)=0", "F47")]
    for ref, cond, labref in rules:
        f = Fx(cond)
        ws.conditional_formatting.add(ref, FormulaRule(formula=[f], fill=PatternFill(bgColor=NA_FILL), font=Font(color=NA_FONT)))
        ws.conditional_formatting.add(labref, FormulaRule(formula=[f], font=Font(color=NA_FONT)))
    for rr in range(15, LAST + 1):
        if ws.row_dimensions[rr].height is None: ws.row_dimensions[rr].height = 14
    ws.print_area = f"A1:K{LAST}"
    ws.row_breaks.append(Break(id=64))
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
EXAMPLES = (("S1", s1), ("B1", b1))
if __name__ == "__main__":
    for nm_, dd in EXAMPLES: build_tab(nm_, dd)
    wb._sheets = [wb[n] for n in ("S1", "B1", "Bars")]
    wb.active = 0
    from openpyxl.workbook.properties import CalcProperties
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    wb.save(OUT)
    print("saved", OUT)
