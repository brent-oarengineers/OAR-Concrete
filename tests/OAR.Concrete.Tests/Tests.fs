/// Console test runner (no test framework needed): dotnet run --project tests/OAR.Concrete.Tests
/// Expected values are hand calculations or come from tools/reference_check.py (an independent
/// scipy implementation) and Brent's Circular_Pile_AS3600.py.
module Tests

open System
open OAR.Concrete
open OAR.Concrete.Materials

let mutable failures = 0
let mutable count = 0

/// relative tolerance on |expected| (absolute below 1)
let check name (expected: float) (actual: float) (tol: float) =
    count <- count + 1
    let ok = abs (actual - expected) <= tol * max 1.0 (abs expected)
    if not ok then failures <- failures + 1
    printfn "%s %-62s expected %12.4f  got %12.4f" (if ok then "PASS" else "FAIL") name expected actual

let bar a depth fsy = Section.barAt a depth fsy 0.0 EsDefault
let rc fc model = concrete fc None model

[<EntryPoint>]
let main argv =
    // ---------------- materials ----------------
    check "Ec f'c 32 (Table 3.1.2)" 30100.0 (ecTable312 32.0) 1e-9
    check "Ec f'c 36 (interpolated 32-40)" 31450.0 (ecTable312 36.0) 1e-9
    check "alpha2 f'c 32 = 0.85 - 0.0015 f'c" 0.802 (alpha2 32.0) 1e-12
    check "gamma f'c 32 = 0.97 - 0.0025 f'c" 0.89 (gamma 32.0) 1e-12
    check "alpha2 f'c 100 floor 0.67" 0.70 (alpha2 100.0) 1e-12
    check "gamma f'c 125 floor 0.67 (formula only)" 0.67 (gamma 125.0) 1e-12

    // ---------------- singly reinforced slab strip: closed form ----------------
    // 1000 x 200, 550 mm² at d = 164, f'c 32, fsy 500
    let slab = Section.rectangle 1000.0 200.0 [ bar 550.0 164.0 500.0 ]
    let r = Ultimate.muo slab (rc 32.0 Rectangular) 90.0
    let dnHand = 550.0 * 500.0 / (0.802 * 32.0 * 0.89 * 1000.0)
    check "Rect block  dn = Ast fsy / (a2 f'c g b)" dnHand r.Dn 1e-7
    check "Rect block  Muo = T (d - g dn / 2)" (275.0 * (164.0 - 0.89 * dnHand / 2.0) / 1000.0) r.Mu 1e-7
    check "Rect block  kuo = dn / do" (dnHand / 164.0) r.Kuo 1e-7
    check "Rect block  Cc = T" 275.0 r.Cc 1e-7
    // other stress-strain laws (reference_check.py)
    let r = Ultimate.muo slab (rc 32.0 Parabolic) 90.0
    check "Parabolic   dn (reference_check)" 12.2768 r.Dn 2e-5
    check "Parabolic   Muo (reference_check)" 43.7335 r.Mu 2e-5
    let r = Ultimate.muo slab (rc 32.0 Curvilinear) 90.0
    check "Curvilinear dn (reference_check)" 12.6065 r.Dn 2e-5
    check "Curvilinear Muo (reference_check)" 43.5819 r.Mu 2e-5

    // ---------------- doubly reinforced beam, both directions ----------------
    // 300 x 600, 3N28 top at 62 (1860 mm²), 4N32 bottom at 538 (3200 mm²), f'c 40
    let beam = Section.rectangle 300.0 600.0 [ bar 1860.0 62.0 500.0; bar 3200.0 538.0 500.0 ]
    let cases =
        [ Rectangular, 90.0, 130.2298, 767.3464
          Rectangular, -90.0, 77.7833, 460.7481
          Parabolic, 90.0, 129.0374, 772.1905
          Parabolic, -90.0, 75.7264, 462.6211
          Curvilinear, 90.0, 134.5312, 765.9069
          Curvilinear, -90.0, 77.2281, 460.2367 ]
    for model, theta, dn, mu in cases do
        let r = Ultimate.muo beam (rc 40.0 model) theta
        let tag = sprintf "Beam %A %s" model (if theta > 0.0 then "sag" else "hog")
        check (tag + "  dn") dn r.Dn 2e-5
        check (tag + "  Muo") mu r.Mu 2e-5
    let r = Ultimate.muo beam (rc 40.0 Rectangular) 90.0
    check "Beam rect sag  kuo = dn / 538" (130.2298 / 538.0) r.Kuo 2e-5
    check "Beam rect sag  N = 0" 0.0 r.N 1e-6

    // ---------------- cracked elastic: closed form (transformed section) ----------------
    let c = Service.analyse slab 30100.0 90.0 20.0 0.0
    let n = 200000.0 / 30100.0
    let rho = 550.0 / (1000.0 * 164.0)
    let k = sqrt (2.0 * rho * n + (rho * n) ** 2.0) - rho * n
    let x = k * 164.0
    let icr = 1000.0 * x ** 3.0 / 3.0 + n * 550.0 * (164.0 - x) ** 2.0
    check "Cracked slab  dn = kd" x c.Dn 1e-6
    check "Cracked slab  Icr = b x^3/3 + n Ast (d - x)^2 (x10^6)" (icr / 1e6) (c.Icr / 1e6) 1e-5
    check "Cracked slab  sigma_s = n M (d - x) / Icr" (n * 20e6 * (164.0 - x) / icr) c.SigSt 1e-5
    check "Cracked slab  sigma_c = M x / Icr" (20e6 * x / icr) c.SigC 1e-5
    let c = Service.analyse beam 32800.0 90.0 250.0 0.0
    check "Cracked beam sag  dn (reference_check, (n-1)Asc)" 191.8733 c.Dn 1e-5
    check "Cracked beam sag  sigma_st" 164.6820 c.SigSt 1e-5
    check "Cracked beam sag  Icr (x10^6)" 3203.9461 (c.Icr / 1e6) 1e-5
    let c = Service.analyse beam 32800.0 -90.0 250.0 0.0
    check "Cracked beam hog  dn" 144.2897 c.Dn 1e-5
    check "Cracked beam hog  sigma_st" 276.7181 c.SigSt 1e-5
    check "Cracked beam hog  sigma_c" 16.6318 c.SigC 1e-5

    // ---------------- Brent's circular pile (Circular_Pile_AS3600.py) ----------------
    // D 450, 4/N20 (310 mm²) on PCD 2 x (225 - 100 - 10) = 230, f'c 40, Ec 30100, curvilinear curve
    let pile = Section.circular 450.0 (Section.circularBars 450.0 230.0 4 310.0 500.0)
    let cv = concrete 40.0 (Some 30100.0) Curvilinear
    let r = Ultimate.muo pile cv 90.0
    check "Pile curvilinear  Mu (Brent's script 106.345)" 106.3453 r.Mu 5e-5
    check "Pile curvilinear  dn (Brent's script 85.81)" 85.8100 r.Dn 5e-5
    check "Pile curvilinear  kuo (Brent's script 0.25238)" 0.25238 r.Kuo 5e-5
    // exact circle vs a 720-sided polygon of equal area
    let nSides = 720
    let rEq = 225.0 * sqrt (Math.PI / (float nSides / 2.0 * sin (2.0 * Math.PI / float nSides)))
    let poly =
        [| for i in 0 .. nSides - 1 ->
               let a = 2.0 * Math.PI * float i / float nSides
               { Geometry.X = 225.0 + rEq * cos a; Geometry.Y = -225.0 + rEq * sin a } |]
    let pileP = { pile with Regions = [ { Shape = Geometry.Polygon poly; IsVoid = false } ] }
    let rp = Ultimate.muo pileP cv 90.0
    check "Pile as polygon  Mu = circle Mu" r.Mu rp.Mu 2e-4
    // rotating the bending direction of a symmetric pile changes nothing (8 bars)
    let pile8 = Section.circular 600.0 (Section.circularBars 600.0 460.0 8 450.0 500.0)
    let rr = rc 40.0 Rectangular
    let m0 = (Ultimate.muo pile8 rr 90.0).Mu
    check "Pile 8 bars  Mu at 112.5 deg = Mu at 90 deg (bars between)" m0 (Ultimate.muo pile8 rr 112.5).Mu 0.03
    check "Pile 8 bars  Mu at 135 deg = Mu at 90 deg" m0 (Ultimate.muo pile8 rr 135.0).Mu 1e-6

    // ---------------- M-N curve (column sheet groundwork) ----------------
    let col = Section.rectangle 400.0 400.0 [ bar 1240.0 60.0 500.0; bar 1240.0 340.0 500.0 ]
    let pts = Interaction.curve col (rc 40.0 Rectangular) 90.0 30 0.65 Design.ClassN
    let nuo = (alpha1 40.0 * 40.0 * (160000.0 - 2480.0) + 2480.0 * 500.0) / 1000.0
    check "M-N  squash Nuo = a1 f'c Ac + fsy As (kN)" nuo (pts |> List.last).N 1e-9
    check "M-N  pure tension = -fsy As (kN)" -1240.0 (pts |> List.head).N 1e-9
    let pb = pts |> List.find (fun p -> p.N = 0.0)
    check "M-N  pure bending point = Muo" (Ultimate.muo col (rc 40.0 Rectangular) 90.0).Mu pb.M 1e-9
    // a point with N: equilibrium solve returns the asked N
    let r = Ultimate.analyse col (rc 40.0 Rectangular) 90.0 1500.0
    check "Ultimate with N = 1500 kN  internal N" 1500.0 r.N 1e-6

    // ---------------- Class L: whole section vs L steel only (reference_check.l_steel_only) ----------------
    // S1: 1000 x 175, f'c 32: L9 mesh @ 200 top (290.5 @ 44.3), 3 N16 top (600 @ 48), L9 mesh bottom (290.5 @ 130.7)
    let s1 = Section.rectangle 1000.0 175.0 [ bar 290.5 44.3 500.0; bar 600.0 48.0 500.0; bar 290.5 130.7 500.0 ]
    let flags = [ true; false; true ]
    let c32 = rc 32.0 Rectangular
    let w = Flexure.capacity s1 c32 -90.0 flags Flexure.WholeSection
    check "S1 hog whole section  phi = 0.65 (Class L in tension)" 0.65 w.Phi 1e-12
    check "S1 hog whole section  Muo unreduced" 56.4291 w.Ult.Mu 2e-5
    let l = Flexure.capacity s1 c32 -90.0 flags Flexure.LSteelOnly
    check "S1 hog L steel only  phi (Class N formula)" 0.85 l.Phi 1e-9
    check "S1 hog L steel only  kL = 0.65 / phi" (0.65 / 0.85) l.KL 1e-9
    check "S1 hog L steel only  Muo reduced section" 52.2262 l.Ult.Mu 2e-5
    check "S1 hog L steel only  phi Muo" 44.3923 (l.Phi * l.Ult.Mu) 2e-5
    check "S1 hog L steel only  eps L = 0.003 (130.7 - dn) / dn" (0.003 * (130.7 - l.Ult.Dn) / l.Ult.Dn) l.EpsL 1e-9
    let l = Flexure.capacity s1 c32 90.0 flags Flexure.LSteelOnly
    check "S1 sag L steel only  Muo reduced section" 28.5262 l.Ult.Mu 2e-5
    // phi below 0.85, so k and phi iterate: 300 x 600, 2400 N @ 538 + 1200 L @ 490, f'c 40
    let it = Section.rectangle 300.0 600.0 [ bar 2400.0 538.0 500.0; bar 1200.0 490.0 500.0 ]
    let l = Flexure.capacity it (rc 40.0 Rectangular) 90.0 [ false; true ] Flexure.LSteelOnly
    check "Iterated  phi (reference 0.832671)" 0.8326711 l.Phi 1e-6
    check "Iterated  kL = 0.65 / phi" 0.7806203 l.KL 1e-6
    check "Iterated  phi Muo (reference 606.4303)" 606.4303 (l.Phi * l.Ult.Mu) 2e-5
    let n = Flexure.capacity slab (rc 32.0 Rectangular) 90.0 [ false ] Flexure.LSteelOnly
    check "No Class L bars  kL = 1, phi from kuo" (Design.phiBending Design.ClassN n.Ult.Kuo) n.Phi 1e-12

    printfn ""
    printfn "%d checks, %d failed" count failures
    if failures = 0 then 0 else 1
