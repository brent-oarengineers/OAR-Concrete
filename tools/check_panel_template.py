"""Independent check of template/OAR - Tilt Panel Design.xlsx (needs LibreOffice and openpyxl).

    python tools/check_panel_template.py

1. LibreOffice recalculates the template. LibreOffice cannot run the add-in, so every OAR.CONCRETE.* cell is
   replaced by a value from the small section routines below (rectangular block, strain compatibility, cracked
   section - written separately from the F# engine), and the workbook is recalculated until the stubs settle.
2. Every tab is then recomputed from its inputs by `panel()` below (statics, magnifier, shear, deflection, band)
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
    fc = num(wsV["E37"].value); phi0 = num(wsV["E40"].value); ec = None
    for row in wsF.iter_rows():
        for c in row:
            f = c.value.text if isinstance(c.value, ArrayFormula) else c.value
            if not (isinstance(f, str) and f.startswith("=") and "OAR.CONCRETE." in f): continue
            if "VERSION" in f: continue
            m = re.search(r'"Rect",\$([A-Z])\$(\d+):\$([A-Z])\$(\d+),\$L\$(\d+)', f)
            b = num(wsV[f"{m.group(1)}{m.group(2)}"].value); D = num(wsV[f"{m.group(3)}{m.group(4)}"].value)
            bs = bars(int(m.group(5)))
            if ec is None: ec = num(wsV["E122"].value)
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
                mm = re.search(r"\$D\$(\d+),\$J\$(\d+)", f)
                Nn, Mm = num(wsV[f"D{mm.group(1)}"].value), num(wsV[f"J{mm.group(2)}"].value)
                if Nn is None: vals = [""] * 10
                else:
                    phi, mu, pmu = column(b, D, bs, fc, Nn, Mm, phi0); vals = [phi, mu, pmu, 0, 0, 0, 0, 0, 0, 0]
                for k, v in enumerate(vals): out[f"{col}{row_ + k}"] = v
            elif "OAR.CONCRETE.ULT(" in f:
                band_on = wsV[ref].value not in (None, "")  # value only matters when the band is active
                mu, phi = ult0(b, D, bs, fc, '"Sagging"' in f) if bs else (0, 0)
                out[ref] = (mu if ",13,1)" in f else phi)
            elif "OAR.CONCRETE.CRACKED(" in f:
                mm = re.search(r"ABS\((\$[A-Z]+\$\d+)\)", f)
                M = num(wsV[cellref(mm.group(1))].value)
                out[ref] = "" if not M else icr(b, D, bs, ec, M >= 0)
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


# ------------------------------------------------------------------ independent panel calc from the inputs
def panel(ws):
    g = lambda r: ws[r].value
    tw, fc, phi0, rho = g("E36"), g("E37"), g("E40"), g("E41")
    Hs, hp, W, top, hb, k = g("E45"), g("E46") or 0, g("E47"), g("E48"), g("E49"), g("E50")
    b, bt, br, gadd = g("E54"), g("E55"), g("E56"), g("E57") or 0
    Gr, Qr, Wr, e, ceq = g("E69"), g("E70"), g("E71"), g("E72") or 0, g("E73") or 0
    two = g("E60") == "Two layers"; band = top.startswith("At panel joints") and g("E42") != "Manual (RFEM)"
    Ec = g("E122"); L = Hs / 1000; gs = rho * tw / 1000 * b / 1000
    bars = []
    for r in range(278, 282):
        A = g(f"L{r}") or 0
        if A > 0: bars.append((A, g(f"M{r}"), g(f"N{r}"), g(f"P{r}"), g(f"Q{r}")))
    bbars = []
    for r in range(284, 288):
        A = g(f"L{r}") or 0
        if A > 0: bbars.append((A, g(f"M{r}"), g(f"N{r}"), g(f"P{r}"), g(f"Q{r}")))
    r_ = 0.3 * tw; Le = k * Hs; sl = Le / r_
    axp, axn = axis(b, tw, bars, fc, True), axis(b, tw, bars, fc, False)
    res = []
    pe = ceq * rho * tw / 1000
    def wp(io, qz):
        c = "E" if io == "i" else "F"
        return qz * (g(f"{c}80") * g(f"{c}82") * g(f"{c}81") * g(f"{c}78") - g(f"{c}83") * g(f"{c}79"))
    qzu, qzs = g("E75") or 0, g("F75") or 0
    qzpu = g("E76") if g("E76") not in (None, "") else qzu; qzps = g("F76") if g("F76") not in (None, "") else qzs
    pin_u, pout_u, pin_s, pout_s = wp("i", qzu), wp("o", qzu), wp("i", qzs), wp("o", qzs)
    ppin_u, ppout_u = qzpu * g("E84"), qzpu * g("F84")
    sls_in = abs(pin_s) >= abs(pout_s)
    expect_p = {93: (pin_u, ppin_u), 94: (pin_u, ppin_u), 95: (pout_u, ppout_u), 96: (pout_u, ppout_u), 97: (pe, pe),
                98: (pin_s if sls_in else pout_s, qzps * (g("E84") if sls_in else g("F84")))}
    for i in range(7):
        lr = 92 + i
        gG = g(f"E{lr}")
        if gG in (None, ""): res.append(None); continue
        gQ, gW = g(f"F{lr}") or 0, g(f"G{lr}") or 0
        p, pp = g(f"H{lr}"), g(f"I{lr}")
        p = pe if isinstance(p, str) else (p or 0); pp = pe if isinstance(pp, str) else (pp or 0)
        w = p * bt / 1000; wpp = pp * bt / 1000
        Proof = (gG * Gr + gQ * Qr + gW * Wr) * br / 1000
        Ptop = Proof + gG * (gadd + gs * hp / 1000)
        MT = -wpp * (hp / 1000) ** 2 / 2 - Proof * e / 1000
        # max |M| by brute force along the span
        xs = [L * j / 2000 for j in range(2001)]
        Ms = [w * x * (L - x) / 2 + MT * x / L for x in xs]
        j = max(range(len(xs)), key=lambda j: abs(Ms[j]))
        x, M2 = xs[j], Ms[j]
        Nst = Ptop + gG * gs * (L - x)
        Rb = w * L / 2 + MT / L; Rt = w * L - Rb + wpp * hp / 1000
        V = max(abs(Rb), abs(Rb - w * L), abs(wpp * hp / 1000))
        res.append(dict(w=w, Ptop=Ptop, MT=MT, x=x, M2=M2, N=Nst, gG=gG, p=p, Rb=Rb, Rt=Rt, V=V))
    out = {"cases": [], "p": {r: (v, (g(f"H{r}"), g(f"I{r}"))) for r, v in expect_p.items()}}
    for i in range(6):
        c = res[i]
        if c is None: out["cases"].append(None); continue
        N_, M2 = c["N"], c["M2"]
        perm = c["gG"] * (Gr * br / 1000 + gadd + gs * (hp / 1000 + L - c["x"]))
        bdin = g("E101")
        bd = 0 if (N_ <= 0 or (sl <= 40 and N_ <= abs(M2) / (2 * tw / 1000))) else (bdin if bdin not in (None, "") else min(1, max(0, perm / N_)))
        ax = axp if M2 >= 0 else axn
        Nc = math.pi ** 2 / Le ** 2 * 182 * ax["do"] * phi0 * ax["Mub"] * 1e6 / (1 + bd) / 1000
        km = 0.6 if c["p"] == 0 else 1.0
        db = 1 if (N_ <= 0 or sl <= 25) else (99 if N_ >= Nc else max(1, km / (1 - N_ / Nc)))
        Md = (-1 if M2 < 0 else 1) * db * max(abs(M2), 0.05 * tw / 1000 * N_)
        phi, mu, pmu = column(b, tw, bars, fc, N_, Md, phi0)
        rt = None
        if abs(c["MT"]) > 0:
            rt = abs(c["MT"]) / column(b, tw, bars, fc, c["Ptop"], c["MT"], phi0)[2]
        bandr = None
        if band:
            q = c["Rt"] / (bt / 1000); Mb = q * (W / 1000) ** 2 / 8
            mu_b, ph_b = ult0(hb, tw, bbars, fc, Mb >= 0)
            bandr = abs(Mb) / (mu_b * ph_b)
        out["cases"].append(dict(Md=abs(Md), pmu=pmu, ratio=abs(Md) / pmu, db=db, top=rt, band=bandr, V=c["V"], Rt=c["Rt"] / (bt / 1000)))
    # deflection
    s = res[6]
    if s:
        dd = bars[1][1] if two else tw / 2
        Ast = (bars[1][0] + (bars[3][0] if len(bars) > 3 else 0)) if two else sum(t[0] for t in bars)
        Asc = (bars[0][0] + (bars[2][0] if len(bars) > 2 else 0)) if two else 0
        rw, rcw = Ast / (b * dd), Asc / (b * dd)
        scs = max(0, (2.5 * rw - 0.8 * rcw) / (1 + 50 * rw) * ES * g("H100"))
        Ns = s["Ptop"] + s["gG"] * gs * L / 2
        Ms_ = s["M2"]; Ig = b * tw ** 3 / 12; Z = b * tw ** 2 / 6
        Mcr = max(0, Z * (0.6 * math.sqrt(fc) - scs + Ns * 1000 / (b * tw)) / 1e6)
        Imax = (1 if rw >= 0.005 else 0.6) * Ig
        Icr_ = icr(b, tw, bars, Ec, Ms_ >= 0) * 1e6
        Ief = Imax if abs(Ms_) <= Mcr else min(Imax, Icr_ / (1 - (1 - Icr_ / Ig) * (Mcr / Ms_) ** 2))
        d1 = (5 * s["w"] * Hs ** 4 / 384 + s["MT"] * 1e6 * Hs ** 2 / 16) / (Ec * Ief)
        Ncr = math.pi ** 2 * Ec * Ief / Le ** 2 / 1000
        d2 = d1 / (1 - max(0, Ns) / Ncr)
        db_ = 0
        if band:
            qs = s["Rt"] / (bt / 1000); Msb = qs * (W / 1000) ** 2 / 8
            Igb = hb * tw ** 3 / 12; Mcrb = max(0, hb * tw ** 2 / 6 * (0.6 * math.sqrt(fc) - scs) / 1e6)
            Icrb = icr(hb, tw, bbars, Ec, Msb >= 0) * 1e6
            Imb = (1 if sum(t[0] for t in bbars) / (hb * dd) >= 0.005 else 0.6) * Igb
            Iefb = Imb if abs(Msb) <= Mcrb else min(Imb, Icrb / (1 - (1 - Icrb / Igb) * (Mcrb / Msb) ** 2))
            db_ = 5 * qs * W ** 4 / (384 * Ec * Iefb)
        out["defl"] = abs(d2) + abs(db_) / 2
    return out


def main():
    wb = run_stubbed(TEMPLATE)
    bad = 0
    def cmp(tab, what, sheet, ref, tol=2e-3):
        nonlocal bad
        ok = (sheet is None and ref is None) or (sheet not in (None, "") and ref is not None and abs(sheet - ref) <= tol * max(1, abs(ref)))
        if not ok: bad += 1
        print(f"  {'ok ' if ok else 'BAD'} {what:<26} sheet {sheet!s:>12.10}   ref {ref!s:>12.10}")
    for tab in [n for n in wb.sheetnames if n != "Bars"]:
        ws = wb[tab]
        ref = panel(ws)
        print(f"{tab}: {ws['B10'].value}  ->  {ws['I10'].value}, max {ws['F10'].value}")
        for i, c in enumerate(ref["cases"]):
            if c is None: continue
            col = "EFGHIJ"[i]
            cmp(tab, f"LC{i + 1} |M*d|", ws[f"{col}182"].value, c["Md"])
            cmp(tab, f"LC{i + 1} δb", ws[f"H{162 + i}"].value, c["db"])
            cmp(tab, f"LC{i + 1} ratio", ws[f"{col}183"].value, c["ratio"])
            if c["top"] is not None: cmp(tab, f"LC{i + 1} top ratio", ws[f"{col}190"].value, c["top"])
            cmp(tab, f"LC{i + 1} V*", ws[f"{col}195"].value, c["V"])
            if c["band"] is not None: cmp(tab, f"LC{i + 1} band ratio", ws[f"{col}214"].value, c["band"])
        cmp(tab, "SLS deflection", ws["E233"].value, ref.get("defl"))
        for r, ((pv, ppv), (ps, pps)) in ref["p"].items():
            cmp(tab, f"row {r} p", ps, pv); cmp(tab, f"row {r} pp", pps, ppv)
        for r in range(15, 34):
            print(f"     {str(ws[f'B{r}'].value or ''):<16} {str(ws[f'D{r}'].value or '')[:70]:<70} {ws[f'E{r}'].value!s:>10.8} {ws[f'F{r}'].value!s:>10.8} {ws[f'I{r}'].value!s:>8.6} {ws[f'J{r}'].value}")
    print("ALL OK" if bad == 0 else f"{bad} mismatches")
    return bad


if __name__ == "__main__":
    sys.exit(1 if main() else 0)
