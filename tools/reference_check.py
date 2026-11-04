"""Independent re-implementation of the OAR Concrete engine, for cross-checking.

Written separately from the F# code on purpose: it integrates the concrete stress with scipy.quad
over a rectangle (y = depth from the top face) and solves the neutral axis with brentq, where the
engine cuts polygons / circles into strips and bisects. If the two agree, both are right.

    python tools/reference_check.py          # prints the benchmark values used in tests/ and README

Sign convention: compression positive. Units N, mm, MPa (results kN, kN·m).
"""
import math
from scipy import integrate, optimize

ES = 200000.0
ECU = 0.003
FC_T = [20, 25, 32, 40, 50, 65, 80, 100]
EC_T = [24000, 26700, 30100, 32800, 34800, 37400, 39600, 42200]


def ec_table(fc):
    for i in range(len(FC_T) - 1):
        if FC_T[i] <= fc <= FC_T[i + 1]:
            return EC_T[i] + (fc - FC_T[i]) / (FC_T[i + 1] - FC_T[i]) * (EC_T[i + 1] - EC_T[i])
    raise ValueError(fc)


def alpha2(fc): return max(0.67, 0.85 - 0.0015 * fc)
def gamma(fc): return max(0.67, 0.97 - 0.0025 * fc)


def law(model, fc, ec):
    """stress(eps), break strain"""
    if model == "rect":
        e1 = ECU * (1 - gamma(fc))
        return (lambda e: alpha2(fc) * fc if e >= e1 else 0.0), e1
    if model == "para":
        fm = 0.9 * fc
        return (lambda e: 0.0 if e <= 0 else (fm * (2 * e / 0.002 - (e / 0.002) ** 2) if e < 0.002 else fm)), 0.002
    if model == "curv":   # Brent's AS3600StressCurve
        scp = 0.9 * fc
        ecp = 4.11 * scp ** 0.75 / ec
        n = ec / (ec - scp / ecp)
        kd = max(0.67 + scp / 62, 1.0)
        def s(e):
            if e <= 0: return 0.0
            eta = e / ecp
            k = 1.0 if e <= ecp else kd
            return scp * n * eta / (n - 1 + eta ** (n * k))
        return s, ecp
    raise ValueError(model)


def ult_rect(b, D, layers, fc, model="rect", ec=None, sag=True):
    """layers: [(As, depth_from_top, fsy)]. sag=True: top in compression. Returns dict."""
    ec = ec or ec_table(fc)
    sig, ebrk = law(model, fc, ec)
    # depth measured from the compression face
    L = [(a, (d if sag else D - d), fy) for a, d, fy in layers if a > 0]

    def forces(dn):
        eps = lambda y: ECU * (dn - y) / dn
        top = min(dn, D)
        ybrk = dn * (1 - ebrk / ECU)
        pts = [ybrk] if 0 < ybrk < top else None
        cc = integrate.quad(lambda y: b * sig(eps(y)), 0, top, points=pts, limit=200)[0]
        mc = integrate.quad(lambda y: b * sig(eps(y)) * (D / 2 - y), 0, top, points=pts, limit=200)[0]
        n, m, t, tm = cc, mc, 0.0, 0.0
        for a, y, fy in L:
            e = eps(y)
            s = max(-fy, min(fy, ES * e))
            f = (s - (sig(e) if e > 0 else 0)) * a
            n += f; m += f * (D / 2 - y)
            if e < 0: t -= f; tm -= f * y
        return n, m, t, tm, cc

    dn = optimize.brentq(lambda x: forces(x)[0], 1e-6 * D, 1e3 * D, xtol=1e-12, rtol=1e-14)
    n, m, t, tm, cc = forces(dn)
    do = max(y for a, y, fy in L if y > dn)
    return dict(dn=dn, kuo=dn / do, do=do, d=tm / t, Mu=m / 1e6, T=t / 1e3, Cc=cc / 1e3)


def cracked_rect(b, D, layers, ec, M, sag=True):
    """Closed form: b x²/2 + Σ(compression) (n−1)As(x − y) − Σ(tension) n As (y − x) = 0."""
    n = ES / ec
    L = [(a, (d if sag else D - d)) for a, d, fy in layers if a > 0]
    def f(x):
        v = b * x * x / 2
        for a, y in L:
            v += ((n - 1) if y < x else n) * a * (x - y)
        return v
    x = optimize.brentq(f, 1e-9, D, xtol=1e-12)
    icr = b * x ** 3 / 3 + sum((((n - 1) if y < x else n) * a * (y - x) ** 2) for a, y in L)
    ymax = max(y for a, y in L)
    sig_st = n * M * 1e6 * (ymax - x) / icr
    return dict(dn=x, Icr=icr, sig_st=sig_st, sig_c=M * 1e6 * x / icr)


def l_steel_only(b, D, layers, classes, fc, sag=True):
    """Class L areas x 0.65/phi, phi from the Class N formula on the reduced section, iterated."""
    phi = 0.85
    for _ in range(50):
        lay = [(a * (0.65 / phi if c == "L" else 1.0), d, fy) for (a, d, fy), c in zip(layers, classes)]
        r = ult_rect(b, D, lay, fc, sag=sag)
        new = min(0.85, max(0.65, 1.24 - 13 * r["kuo"] / 12))
        if abs(new - phi) < 1e-10: break
        phi = new
    return dict(phi=phi, k=0.65 / phi, Mu=r["Mu"], kuo=r["kuo"], phiMu=phi * r["Mu"])


def pile_brent():
    """Brent's Circular_Pile_AS3600.py example: D450, 4/N20 (310), cover 100 + lig 10, f'c 40, Ec 30100"""
    import importlib.util, os, sys, io, contextlib
    here = os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(here, "Circular_Pile_AS3600.py")
    if not os.path.exists(path):
        return None
    spec = importlib.util.spec_from_file_location("pile", path)
    mod = importlib.util.module_from_spec(spec)
    with contextlib.redirect_stdout(io.StringIO()):
        spec.loader.exec_module(mod)
    return dict(Mu=mod.Calc.Mu() / 1e6, kuo=mod.Calc.kuo(), dn=mod.Calc.dNA())


if __name__ == "__main__":
    print("Singly reinforced 1000x200, 550 mm2 @ 164, f'c 32:")
    for m in ("rect", "para", "curv"):
        r = ult_rect(1000, 200, [(550, 164, 500)], 32, m)
        print(f"  {m}: dn {r['dn']:.4f}  kuo {r['kuo']:.5f}  Muo {r['Mu']:.4f}")
    print("Doubly reinforced 300x600, 3N28 top @ 62 (1860), 4N32 bottom @ 538 (3200), f'c 40:")
    lay = [(1860, 62, 500), (3200, 538, 500)]
    for m in ("rect", "para", "curv"):
        for sag in (True, False):
            r = ult_rect(300, 600, lay, 40, m, sag=sag)
            print(f"  {m} {'sag' if sag else 'hog'}: dn {r['dn']:.4f}  kuo {r['kuo']:.5f}  d {r['d']:.3f}  Muo {r['Mu']:.4f}")
    print("Cracked, same beams:")
    r = cracked_rect(1000, 200, [(550, 164, 500)], 30100, 20)
    print(f"  slab M 20: dn {r['dn']:.4f}  Icr {r['Icr']:.1f}  sig_st {r['sig_st']:.4f}  sig_c {r['sig_c']:.4f}")
    for sag in (True, False):
        r = cracked_rect(300, 600, lay, 32800, 250, sag=sag)
        print(f"  beam M 250 {'sag' if sag else 'hog'}: dn {r['dn']:.4f}  Icr {r['Icr']:.1f}  sig_st {r['sig_st']:.4f}  sig_c {r['sig_c']:.4f}")
    S1 = [(290.5, 44.3, 500), (600, 48, 500), (290.5, 130.7, 500)]
    for sag in (True, False):
        r = l_steel_only(1000, 175, S1, ["L", "N", "L"], 32, sag=sag)
        print(f"S1 L steel only {'sag' if sag else 'hog'}: phi {r['phi']:.5f}  k {r['k']:.5f}  Muo {r['Mu']:.4f}  phiMuo {r['phiMu']:.4f}")
    p = pile_brent()
    if p: print(f"Brent's pile (curvilinear): Mu {p['Mu']:.4f}  kuo {p['kuo']:.5f}  dn {p['dn']:.4f}")
