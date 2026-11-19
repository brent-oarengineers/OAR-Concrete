"""Independent check of template/OAR - Tilt Panel Design.xlsx (needs LibreOffice and openpyxl).

    python tools/check_panel_template.py

1. LibreOffice recalculates the template. LibreOffice cannot run the add-in, so every OAR.CONCRETE.* cell is
   replaced by a value from the small section routines below (rectangular block, strain compatibility, cracked
   section - written separately from the F# engine), and the workbook is recalculated until the stubs settle.
2. Every tab is then recomputed from its inputs by `panel()` below (design moments, capacity, SLS bar stresses)
   and compared with the sheet's results.
This checks the worksheet formulas; the engine itself is checked by tests/OAR.Concrete.Tests.
"""
import math, os, re, shutil, subprocess, sys, tempfile
import openpyxl
from openpyxl.worksheet.formula import ArrayFormula

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "..", "template", "OAR - Tilt Panel Design.xlsx")
ECU, ES = 0.003, 200000.0


# ------------------------------------------------------------------ section routines (rect b x D, depth y from the top face)
def a2(fc): return max(0.67, 0.85 - 0.0015 * fc)
def gam(fc): return max(0.67, 0.97 - 0.0025 * fc)
def a1(fc): return min(0.85, max(0.72, 1.0 - 0.003 * fc))


def flip(bars, D, top): return bars if top else [(A, D - y, fy, es, c) for A, y, fy, es, c in bars]


def forces(b, D, bars, fc, dn):
    """N (N) and M about mid-depth (Nmm), top in compression, rectangular block."""
    g = gam(fc); a = min(g * dn, D)
    C = a2(fc) * fc * a * b
    n, m = C, C * (D / 2 - a / 2)
    for A, y, fy, es, _ in bars:
        e = ECU * (dn - y) / dn
        s = max(-fy, min(fy, es * e)) - (a2(fc) * fc if y <= a and e > 0 else 0.0)
        n += s * A; m += s * A * (D / 2 - y)
    return n, m


def solve_dn(b, D, bars, fc, N):
    lo, hi = 1e-6 * D, D / gam(fc)
    for _ in range(200):
        mid = 0.5 * (lo + hi)
        if forces(b, D, bars, fc, mid)[0] < N: lo = mid
        else: hi = mid
    return 0.5 * (lo + hi)


def phib(bars, kuo, dn, D):
    anyL_tension = any(c == "L" and y > dn for _, y, _, _, c in bars)
    return 0.65 if any(c == "L" for *_, c in bars) else min(0.85, max(0.65, 1.24 - 13 * kuo / 12)), anyL_tension


def axis(b, D, bars0, fc, top=True):
    bars = flip(bars0, D, top)
    do = max(y for _, y, *_ in bars)
    _, ys, fy, es, _ = max(bars, key=lambda t: t[1])
    dnb = do * ECU / (ECU + fy / es)
    nb, mb = forces(b, D, bars, fc, dnb)
    dn0 = solve_dn(b, D, bars, fc, 0.0); _, m0 = forces(b, D, bars, fc, dn0)
    kuo = dn0 / do
    pb = 0.65 if any(c == "L" for *_, c in bars) else min(0.85, max(0.65, 1.24 - 13 * kuo / 12))
    As = sum(t[0] for t in bars)
    nuo = (a1(fc) * fc * (b * D - As) + sum(A * min(f, e_ * ECU) for A, _, f, e_, _ in bars)) / 1000
    return dict(Nuo=nuo, Nub=nb / 1000, Mub=mb / 1e6, Muo=m0 / 1e6, kuo=kuo, phib=pb, do=do,
                list=[nuo, -sum(A * f for A, _, f, _, _ in bars) / 1000, nb / 1000, mb / 1e6, m0 / 1e6, kuo, pb, do, 0, 0, 0.65])


def column(b, D, bars0, fc, N, M, phi0):
    top = M >= 0
    ax = axis(b, D, bars0, fc, top); bars = flip(bars0, D, top)
    dlim = D / gam(fc); nlim, mlim = forces(b, D, bars, fc, dlim)
    if N * 1000 > nlim:
        k = (ax["Nuo"] - N) / (ax["Nuo"] - nlim / 1000); mu = k * mlim / 1e6
    else:
        dn = solve_dn(b, D, bars, fc, N * 1000); mu = forces(b, D, bars, fc, dn)[1] / 1e6
    phi = phi0 if N >= ax["Nub"] else phi0 + (ax["phib"] - phi0) * (1 - N / ax["Nub"])
    return phi, abs(mu), phi * abs(mu)


def ult0(b, D, bars0, fc, top):
    bars = flip(bars0, D, top)
    dn = solve_dn(b, D, bars, fc, 0.0); mu = forces(b, D, bars, fc, dn)[1] / 1e6
    lt = any(c == "L" and y > dn for _, y, _, _, c in bars)
    do = max(y for _, y, *_ in bars)
    phi = 0.65 if lt else min(0.85, max(0.65, 1.24 - 13 * (dn / do) / 12))
    return abs(mu), phi


def icr(b, D, bars0, ec, top):
    bars = flip(bars0, D, top); n = ES / ec
    def fm(x): return b * x * x / 2 + sum(((n - 1) if y < x else n) * A * (x - y) for A, y, *_ in bars)
    lo, hi = 1e-6, D
    for _ in range(200):
        mid = 0.5 * (lo + hi)
        if fm(mid) < 0: lo = mid
        else: hi = mid
    x = 0.5 * (lo + hi)
    return (b * x ** 3 / 3 + sum(((n - 1) if y < x else n) * A * (x - y) ** 2 for A, y, *_ in bars)) / 1e6


def cracked(b, D, bars0, ec, top, M, N):
    """Cracked elastic section at M (kNm) and N (kN, compression ≥ 0): returns dn mm, σst MPa (largest tension), σc MPa."""
    bars = flip(bars0, D, top); n = ES / ec
    def F(x): return b * x * x / 2 + sum(((n - 1) if y < x else n) * A * (x - y) for A, y, *_ in bars)
    def Mo(x): return b * x * x / 2 * (D / 2 - x / 3) + sum(((n - 1) if y < x else n) * A * (x - y) * (D / 2 - y) for A, y, *_ in bars)
    e = M * 1e6 / (N * 1e3) if N > 0 else None
    g = (lambda x: F(x)) if e is None else (lambda x: Mo(x) - e * F(x))
    lo, hi = 1e-6, 50 * D
    for _ in range(300):
        mid = 0.5 * (lo + hi)
        if (g(mid) < 0) == (g(lo) < 0): lo = mid
        else: hi = mid
    x = 0.5 * (lo + hi)
    k = M * 1e6 / Mo(x) if e is None else N * 1e3 / F(x)     # Ec κ
    st = max(n * k * (y - x) for A, y, *_ in bars)
    return x, st, k * x


# ------------------------------------------------------------------ LibreOffice recalculation with stubs
def recalc(path, outdir):
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", outdir, path],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return openpyxl.load_workbook(os.path.join(outdir, os.path.basename(path)), data_only=True)


def num(v):
    try: return float(v)
    except (TypeError, ValueError): return None


def cellref(s): return s.replace("$", "")


def stub_values(wsF, wsV):
    """{cell: value} for every add-in cell of one tab, from the last recalculated values."""
    out = {}
    def bars(r0):
        rows = []
        for r in range(r0, r0 + 4):
            A = num(wsV[f"L{r}"].value) or 0
            if A > 0: rows.append((A, num(wsV[f"M{r}"].value), num(wsV[f"N{r}"].value), num(wsV[f"P{r}"].value) or ES, wsV[f"Q{r}"].value or "N"))
        return rows
    fc = num(wsV["E35"].value); phi0 = num(wsV["E38"].value); ec = None
    for row in wsF.iter_rows():
        for c in row:
            f = c.value.text if isinstance(c.value, ArrayFormula) else c.value
            if not (isinstance(f, str) and f.startswith("=") and "OAR.CONCRETE." in f): continue
            if "VERSION" in f: continue
            m = re.search(r'"Rect",\$([A-Z])\$(\d+):\$([A-Z])\$(\d+),\$L\$(\d+)', f)
            b = num(wsV[f"{m.group(1)}{m.group(2)}"].value); D = num(wsV[f"{m.group(3)}{m.group(4)}"].value)
            bs = bars(int(m.group(5)))
            if ec is None: ec = next(num(wsV[f"E{r}"].value) for r in range(70, 90) if wsV[f"B{r}"].value == "Ec")
            ref = c.coordinate
            col, row_ = re.match(r"([A-Z]+)(\d+)", ref).groups(); row_ = int(row_)
            if "COLUMNAXIS(" in f:
                top = ",90," in f
                for k, v in enumerate(axis(b, D, bs, fc, top)["list"]): out[f"{col}{row_ + k}"] = v
            elif f.startswith("=IF(OR(") and "OAR.CONCRETE.COLUMN(" in f:      # top support φMu
                mm = re.search(r"COLUMN\([^)]*?,([A-J]\d+),([A-J]\d+),0,", f)
                Nn, Mm = num(wsV[mm.group(1)].value), num(wsV[mm.group(2)].value)
                out[ref] = "" if (Nn is None or not Mm) else column(b, D, bs, fc, Nn, Mm, phi0)[2]
            elif "OAR.CONCRETE.COLUMN(" in f:                                   # capacity array
                mm = re.search(r"\$D\$(\d+),\$G\$(\d+)", f)
                Nn, Mm = num(wsV[f"D{mm.group(1)}"].value), num(wsV[f"G{mm.group(2)}"].value)
                if Nn is None: vals = [""] * 10
                else:
                    phi, mu, pmu = column(b, D, bs, fc, Nn, Mm, phi0); vals = [phi, mu, pmu, 0, 0, 0, 0, 0, 0, 0]
                for k, v in enumerate(vals): out[f"{col}{row_ + k}"] = v
            elif "OAR.CONCRETE.ULT(" in f:
                band_on = wsV[ref].value not in (None, "")  # value only matters when the band is active
                mu, phi = ult0(b, D, bs, fc, '"Sagging"' in f) if bs else (0, 0)
                out[ref] = (mu if ",13,1)" in f else phi)
            elif "OAR.CONCRETE.CRACKED(" in f:
                mm = re.search(r',([EF]\d+),"(Sagging|Hogging)",MAX\(0,N\((\$[A-Z]+\$\d+)\)\)', f)
                M = num(wsV[mm.group(1)].value); Nn = max(0, num(wsV[cellref(mm.group(3))].value) or 0)
                if not M: out[ref] = ""
                else:
                    dn_, st_, sc_ = cracked(b, D, bs, ec, mm.group(2) == "Sagging", M, Nn)
                    out[ref] = dn_ if ",1,1)" in f else (st_ if ",3,1)" in f else sc_)
    return out


def run_stubbed(template, iters=4):
    tmp = tempfile.mkdtemp()
    work = os.path.join(tmp, "panel.xlsx"); shutil.copy(template, work)
    wbF = openpyxl.load_workbook(template)
    vals = {}
    for it in range(iters):
        wb = openpyxl.load_workbook(template)
        for ws in wb.worksheets:
            if ws.title == "Bars": continue
            ws["V4"].value = "9.9.9"
            for ref, v in vals.get(ws.title, {}).items():
                c = ws[ref]
                c.value = v
            # cells of array formulas not overwritten yet: blank them so LibreOffice does not see the add-in
            for row in ws.iter_rows():
                for c in row:
                    f = c.value.text if isinstance(c.value, ArrayFormula) else c.value
                    if isinstance(f, str) and f.startswith("=") and "OAR.CONCRETE." in f and c.coordinate != "V4": c.value = 0
        wb.save(work)
        out = os.path.join(tmp, f"out{it}"); os.makedirs(out)
        wbV = recalc(work, out)
        new = {ws.title: stub_values(wbF[ws.title], wbV[ws.title]) for ws in wbF.worksheets if ws.title != "Bars"}
        if new == vals: break
        vals = new
    run_stubbed.last = os.path.join(out, 'panel.xlsx')
    return wbV


# ------------------------------------------------------------------ independent member checks from the inputs
def panel(ws):
    g = lambda r: ws[r].value
    tw, fc, phi0 = g("E34"), g("E35"), g("E38")
    bw = g("E71")   # strip width: 1000 (unit width) or the isolated leg / head width
    Ec = next(ws[f"E{r}"].value for r in range(70, 90) if ws[f"B{r}"].value == "Ec")
    def rows(r0):
        out = []
        for r in range(r0, r0 + 4):
            A = g(f"L{r}") or 0
            if A > 0: out.append((A, g(f"M{r}"), g(f"N{r}"), g(f"P{r}"), g(f"Q{r}")))
        return out
    bars, hbars = rows(250), rows(256)
    two = g("E42") == "Two layers"
    hx = {True: ult0(1000, tw, hbars, fc, True), False: ult0(1000, tw, hbars, fc, False)}
    out = {"cases": []}
    for i in range(6):
        lr = 55 + i
        N_ = g(f"E{lr}")
        if N_ in (None, ""): out["cases"].append(None); continue
        M2 = g(f"F{lr}") or 0; MT = g(f"G{lr}") or 0; Mx = g(f"H{lr}")
        Md = (-1 if M2 < 0 else 1) * max(abs(M2), 0.05 * tw / 1000 * N_)   # second order from the analysis
        phi, mu, pmu = column(bw, tw, bars, fc, N_, Md, phi0)
        top = abs(MT) / column(bw, tw, bars, fc, N_, MT, phi0)[2] if MT else None
        xr = None
        if Mx not in (None, ""):
            mu_x, ph_x = hx[Mx >= 0]; xr = abs(Mx) / (mu_x * ph_x)
        out["cases"].append(dict(Md=abs(Md), ratio=abs(Md) / pmu, top=top, xr=xr))
    # SLS cracked bar stresses (independent of the stubs: recomputed here)
    Ns = max(0, g("E63") or 0)
    out["sig"] = [cracked(bw, tw, bars, Ec, top, abs(Ms), Ns)[1] if Ms else None
                  for Ms, top in ((g("E64"), True), (g("E65"), False))]
    return out


def main():
    wb = run_stubbed(TEMPLATE)
    bad = 0
    def cmp(what, sheet, ref, tol=2e-3):
        nonlocal bad
        ok = (sheet in (None, "") and ref is None) or (sheet not in (None, "") and ref is not None and abs(sheet - ref) <= tol * max(1, abs(ref)))
        if not ok: bad += 1
        print(f"  {'ok ' if ok else 'BAD'} {what:<26} sheet {sheet!s:>12.10}   ref {ref!s:>12.10}")
    for tab in [n for n in wb.sheetnames if n != "Bars"]:
        ws = wb[tab]
        ref = panel(ws)
        print(f"{tab}: {ws['B10'].value}  ->  {ws['I10'].value}, max {ws['F10'].value}")
        lab = [(r, ws[f"B{r}"].value) for r in range(60, 200)]
        def after(start, name): return next(r for r, v in lab if r > start and v == name)
        RM = after(0, "|M*d|"); TR = after(after(0, "M*y sup"), "ratio"); XR = after(after(0, "M*x"), "ratio"); ST = after(0, "σst")
        for i, c in enumerate(ref["cases"]):
            if c is None: continue
            col = "EFGHIJ"[i]
            cmp(f"LC{i + 1} |M*d|", ws[f"{col}{RM}"].value, c["Md"])
            cmp(f"LC{i + 1} ratio", ws[f"{col}{RM + 1}"].value, c["ratio"])
            cmp(f"LC{i + 1} support ratio", ws[f"{col}{TR}"].value, c["top"])
            cmp(f"LC{i + 1} M*x ratio", ws[f"{col}{XR}"].value, c["xr"])
        cmp("SLS σst Ms+", ws[f"E{ST}"].value, ref["sig"][0])
        cmp("SLS σst Ms−", ws[f"F{ST}"].value, ref["sig"][1])
        for r in range(15, 30):
            print(f"     {str(ws[f'B{r}'].value or ''):<16} {str(ws[f'D{r}'].value or '')[:60]:<60} {ws[f'E{r}'].value!s:>10.8} {ws[f'F{r}'].value!s:>10.8} {ws[f'I{r}'].value!s:>8.6} {ws[f'J{r}'].value}")
    print("ALL OK" if bad == 0 else f"{bad} mismatches")
    return bad


if __name__ == "__main__":
    sys.exit(1 if main() else 0)
