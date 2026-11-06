"""OAR house calc style (oar-calc-sheet skill): shared by the template builders."""
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule, CellIsRule, DataBarRule
from openpyxl.drawing.image import Image as XLImage

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



# AS/NZS 4671 nominal areas. S = 250 MPa deformed (D250N), L = mesh wires (Class L), R = plain round (R250N).
BARS = [
    ["None", 0, 0, 500, "N"], ["N10", 10, 80, 500, "N"], ["N12", 12, 110, 500, "N"], ["N16", 16, 200, 500, "N"], ["N20", 20, 310, 500, "N"],
    ["N24", 24, 450, 500, "N"], ["N28", 28, 620, 500, "N"], ["N32", 32, 800, 500, "N"], ["N36", 36, 1020, 500, "N"],
    ["N40", 40, 1260, 500, "N"],
    ["S10", 10, 80, 250, "N"], ["S12", 12, 110, 250, "N"], ["S16", 16, 200, 250, "N"], ["S20", 20, 310, 250, "N"],
    ["S24", 24, 450, 250, "N"], ["S28", 28, 620, 250, "N"], ["S32", 32, 800, 250, "N"],
    ["R10", 10, 78, 250, "N"], ["R12", 12, 113, 250, "N"],
    ["L6", 6.0, 28.3, 500, "L"], ["L7", 6.75, 35.8, 500, "L"], ["L8", 7.6, 45.4, 500, "L"], ["L9", 8.6, 58.1, 500, "L"],
    ["L10", 9.5, 70.9, 500, "L"], ["L11", 10.65, 89.1, 500, "L"], ["L12", 11.9, 111.2, 500, "L"],
]
BAR_NOTES = {"None": "no bars (single-layer sections)", "S12": "deformed 250 MPa (D250N), N-bar areas", "R10": "plain round R250N", "L9": "SL92 mesh wire", "L12": "RL1218 main wire"}
