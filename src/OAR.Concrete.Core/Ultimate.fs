/// Ultimate strength by strain compatibility (AS 3600-2018 Cl 8.1.2, 8.1.3, 10.6.2).
///
/// Plane sections remain plane. The extreme compression fibre is at εcu = 0.003 and the neutral axis
/// lies at depth dn, measured at right angles to the neutral axis. For a chosen direction θ the solver
/// finds dn so that the internal force equals the applied axial force N (0 for a beam or slab),
/// then sums moments. Concrete in tension carries nothing. Bars in compression have the concrete
/// they displace subtracted (σs − σc) × As.
///
/// The column sheet will use the same functions: N ≠ 0 for a point on the M–N curve, and any θ for
/// biaxial bending.
module OAR.Concrete.Ultimate

open System
open OAR.Concrete.Materials
open OAR.Concrete.Geometry
open OAR.Concrete.Section

/// Forces and moments for one neutral axis position. Forces N, moments N·mm, compression +.
type State =
    { Dn: float
      /// s of the extreme compression fibre and of the neutral axis
      STop: float
      SNa: float
      /// concrete compression force, N
      Cc: float
      /// net force in bars in compression (σs − σc) As, N
      Cs: float
      /// force in bars in tension (positive number), N
      T: float
      /// total axial force, N (compression +)
      N: float
      /// moments about the gross centroid, N·mm: Mx = ∫σ (y − yc) dA, My = ∫σ (x − xc) dA
      Mx: float
      My: float
      /// moment about the neutral axis direction (positive = compression on the θ side), N·mm
      Mu: float
      /// distance from the extreme compression fibre to the outermost tension bar, mm (NaN if none)
      Do: float
      /// distance from the extreme compression fibre to the resultant tension force, mm (NaN if none)
      D: float
      /// strain and stress magnitude at the outermost tension bar
      EpsSt: float
      SigSt: float
      /// largest compressive stress in a bar, MPa (0 if none)
      SigSc: float }

/// Evaluate the section for a neutral axis depth dn (mm) in direction θ
let evaluate (sec: Section) (c: Concrete) (thetaDeg: float) (dn: float) (nStrips: int) =
    let d = dirOf thetaDeg
    let sMin, sTop = Section.extent sec d
    let sNa = sTop - dn
    let strain s = EpsCu * (s - sNa) / dn
    let cen = centroid sec
    let sRef = sOf d cen
    // --- concrete compression zone
    let lo = max sMin sNa
    let breaks = breakStrains c |> List.map (fun e -> sNa + e * dn / EpsCu)
    let mutable cc = 0.0
    let mutable mx = 0.0
    let mutable my = 0.0
    for st in strips sec d lo sTop breaks nStrips do
        let sigma = stress c (strain st.S)
        cc <- cc + sigma * st.M.A
        mx <- mx + sigma * (st.M.Sy - cen.Y * st.M.A)
        my <- my + sigma * (st.M.Sx - cen.X * st.M.A)
    // --- bars
    let mutable cs = 0.0
    let mutable t = 0.0
    let mutable tMoment = 0.0 // Σ T·s for the tension resultant
    let mutable sOuter = Double.PositiveInfinity
    let mutable epsOuter = 0.0
    let mutable sigOuter = 0.0
    let mutable sigSc = 0.0
    for b in sec.Bars do
        let p = { X = b.X; Y = b.Y }
        let s = sOf d p
        let eps = strain s
        let sigS = steelStress b.Steel eps
        let displaced = if eps > 0.0 && inConcrete sec p then stress c eps else 0.0
        let f = (sigS - displaced) * b.Area
        if eps > 0.0 then
            cs <- cs + f
            sigSc <- max sigSc sigS
        elif eps < 0.0 then
            t <- t - f
            tMoment <- tMoment - f * s
            if s < sOuter then
                sOuter <- s
                epsOuter <- -eps
                sigOuter <- -sigS
        mx <- mx + f * (b.Y - cen.Y)
        my <- my + f * (b.X - cen.X)
    let n = cc + cs - t
    let doo = if Double.IsInfinity sOuter then nan else sTop - sOuter
    let dd = if t > 0.0 then sTop - tMoment / t else nan
    { Dn = dn
      STop = sTop
      SNa = sNa
      Cc = cc
      Cs = cs
      T = t
      N = n
      Mx = mx
      My = my
      Mu = mx * d.Uy + my * d.Ux
      Do = doo
      D = dd
      EpsSt = epsOuter
      SigSt = sigOuter
      SigSc = sigSc }

/// Depth of the section in direction θ, mm
let depthIn (sec: Section) (thetaDeg: float) =
    let lo, hi = Section.extent sec (dirOf thetaDeg)
    hi - lo

/// Find dn (mm) so the section carries axial force nApplied (N, compression +) in direction θ.
/// Bisection on log(dn): the internal force increases with dn.
let solveDn (sec: Section) (c: Concrete) (thetaDeg: float) (nApplied: float) (nStrips: int) =
    let h = depthIn sec thetaDeg
    let f dn = (evaluate sec c thetaDeg dn nStrips).N - nApplied
    let mutable lo = 1e-6 * h
    let mutable hi = 1e4 * h
    let flo = f lo
    let fhi = f hi
    if flo > 0.0 then failwith "axial force is below the section's tension capacity"
    if fhi < 0.0 then failwith "axial force exceeds the section's compression capacity"
    let mutable i = 0
    while hi / lo - 1.0 > 1e-11 && i < 200 do
        let mid = sqrt (lo * hi)
        if f mid < 0.0 then lo <- mid else hi <- mid
        i <- i + 1
    sqrt (lo * hi)

/// Results reported to the sheet (kN, kN·m, mm, MPa)
type Result =
    { Dn: float
      Kuo: float
      Ku: float
      Do: float
      D: float
      EpsSt: float
      SigSt: float
      SigSc: float
      /// kN
      Cc: float
      Cs: float
      T: float
      /// lever arm Mu / T, mm
      Z: float
      /// kN
      N: float
      /// kN·m
      Mu: float
      Mx: float
      My: float }

let toResult (st: State) =
    { Dn = st.Dn
      Kuo = st.Dn / st.Do
      Ku = st.Dn / st.D
      Do = st.Do
      D = st.D
      EpsSt = st.EpsSt
      SigSt = st.SigSt
      SigSc = st.SigSc
      Cc = st.Cc / 1000.0
      Cs = st.Cs / 1000.0
      T = st.T / 1000.0
      Z = (if st.T > 0.0 then st.Mu / st.T else nan)
      N = st.N / 1000.0
      Mu = st.Mu / 1e6
      Mx = st.Mx / 1e6
      My = st.My / 1e6 }

/// Ultimate capacity in direction θ with axial force nKN (kN, compression +). 0 for slabs and beams.
let analyse (sec: Section) (c: Concrete) (thetaDeg: float) (nKN: float) =
    if sec.Bars |> List.forall (fun b -> b.Area <= 0.0) then failwith "no reinforcement"
    let dn = solveDn sec c thetaDeg (nKN * 1000.0) DefaultStrips
    evaluate sec c thetaDeg dn DefaultStrips |> toResult

/// Muo for bending only, θ = 90 (sagging, top in compression) or −90 (hogging)
let muo (sec: Section) (c: Concrete) (thetaDeg: float) = analyse sec c thetaDeg 0.0
