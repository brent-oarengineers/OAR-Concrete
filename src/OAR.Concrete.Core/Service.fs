/// Cracked elastic section analysis for serviceability (bar stresses for crack control, Cl 8.6 / 9.5).
///
/// Linear elastic concrete with no tension, linear elastic steel (no yield cap: a stress above fsy
/// is reported as calculated). Bars in compression use (Es − Ec) to allow for the concrete they
/// displace, the same as the familiar (n − 1) Asc. The modular ratio is Es / Ec, so a long-term check
/// is made by passing an effective Ec = Ec / (1 + φcc).
module OAR.Concrete.Service

open System
open OAR.Concrete.Materials
open OAR.Concrete.Geometry
open OAR.Concrete.Section

[<Literal>]
let private Strips = 400

/// Section stiffness terms for a neutral axis at depth dn and unit curvature:
/// force, moment about the gross centroid, moment about the neutral axis (N, N·mm per 1/mm)
let private unitActions (sec: Section) (ec: float) (d: Dir) (dn: float) =
    let sMin, sTop = Section.extent sec d
    let sNa = sTop - dn
    let sRef = sOf d (centroid sec)
    let mutable n = 0.0
    let mutable mRef = 0.0
    let mutable mNa = 0.0
    for st in strips sec d (max sMin sNa) sTop [] Strips do
        let y = st.S - sNa
        let own = st.M.A * st.H * st.H / 12.0 // the strip's own second moment
        n <- n + ec * y * st.M.A
        mRef <- mRef + ec * (y * (st.S - sRef) * st.M.A + own)
        mNa <- mNa + ec * (y * y * st.M.A + own)
    for b in sec.Bars do
        let p = { X = b.X; Y = b.Y }
        let s = sOf d p
        let y = s - sNa
        let e = if y > 0.0 && inConcrete sec p then b.Steel.Es - ec else b.Steel.Es
        n <- n + e * y * b.Area
        mRef <- mRef + e * y * (s - sRef) * b.Area
        mNa <- mNa + e * y * y * b.Area
    n, mRef, mNa

type Cracked =
    { /// neutral axis depth from the extreme compression fibre, mm
      Dn: float
      /// cracked second moment of area about the neutral axis, in concrete units, mm⁴
      Icr: float
      /// largest tensile bar stress, MPa (positive)
      SigSt: float
      /// largest compressive bar stress, MPa
      SigSc: float
      /// concrete stress at the extreme compression fibre, MPa
      SigC: float
      /// curvature, 1/mm
      Kappa: float }

/// Cracked section under moment mKNm (kN·m, magnitude, compression on the θ side) and axial force
/// nKN (kN, compression +, 0 for slabs and beams). ec = concrete modulus to use, MPa.
let analyse (sec: Section) (ec: float) (thetaDeg: float) (mKNm: float) (nKN: float) =
    if ec <= 0.0 then failwith "Ec must be positive"
    if nKN < 0.0 then failwith "cracked analysis with axial tension is not covered (v0.1)"
    let d = dirOf thetaDeg
    let sMin, sTop = Section.extent sec d
    let h = sTop - sMin
    let m = abs mKNm * 1e6
    let nn = nKN * 1000.0
    // equation for dn: N = 0 → force term zero; N > 0 → moment / force = eccentricity
    let g =
        if nn = 0.0 then fun dn -> let f, _, _ = unitActions sec ec d dn in f
        else
            let e = m / nn
            fun dn -> let f, mr, _ = unitActions sec ec d dn in mr - e * f
    let mutable lo = 1e-6 * h
    let mutable hi = 1e4 * h
    let glo, ghi = g lo, g hi
    if Math.Sign glo = Math.Sign ghi then
        failwith "no cracked neutral axis found (section uncracked in compression or no tension steel)"
    let mutable i = 0
    while hi / lo - 1.0 > 1e-11 && i < 200 do
        let mid = sqrt (lo * hi)
        if Math.Sign(g mid) = Math.Sign glo then lo <- mid else hi <- mid
        i <- i + 1
    let dn = sqrt (lo * hi)
    let f, mRef, mNa = unitActions sec ec d dn
    let kappa = if nn = 0.0 then (if mRef = 0.0 then 0.0 else m / mRef) else nn / f
    let sNa = sTop - dn
    let barStress (b: Bar) = b.Steel.Es * kappa * (sOf d { X = b.X; Y = b.Y } - sNa)
    let stresses = sec.Bars |> List.map barStress
    { Dn = dn
      Icr = mNa / ec
      SigSt = stresses |> List.fold (fun acc s -> max acc -s) 0.0
      SigSc = stresses |> List.fold (fun acc s -> max acc s) 0.0
      SigC = ec * kappa * min dn h
      Kappa = kappa }
