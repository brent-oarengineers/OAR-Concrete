/// Columns: capacity at a given axial force in any direction, biaxial bending by strain compatibility,
/// moment contours at constant N, axis properties for slenderness, and typical bar layouts.
///
/// Axes (sheet convention): x across the width b, y up the depth D, origin at the section centre.
/// Mx bends about the x axis (compression on the +y / top face when positive),
/// My bends about the y axis (compression on the +x face when positive).
/// In the engine θ is the direction of maximum compression: θ = 90° gives Mx, θ = 0° gives My.
module OAR.Concrete.Column

open System
open OAR.Concrete.Materials
open OAR.Concrete.Geometry
open OAR.Concrete.Section

// ---------------------------------------------------------------------------------------------
// Bar layouts (coordinates from the section centre, mm)
// ---------------------------------------------------------------------------------------------

/// Rectangular perimeter layout: nx bars along each b face (top and bottom, corners included),
/// ny bars along each D face (corners included). edge = distance from the face to the bar centre.
/// Order: top face left → right, right face down, bottom face right → left, left face up.
let rectLayout (b: float) (depth: float) (edge: float) (nx: int) (ny: int) =
    if nx < 2 || ny < 2 then failwith "a rectangular layout needs at least 2 bars on each face (corners)"
    let x0, x1 = -b / 2.0 + edge, b / 2.0 - edge
    let y0, y1 = -depth / 2.0 + edge, depth / 2.0 - edge
    if x1 < x0 || y1 < y0 then failwith "cover too large for the section"
    let lerp a c i n = a + (c - a) * float i / float (n - 1)
    [ for i in 0 .. nx - 1 -> lerp x0 x1 i nx, y1 ]
    @ [ for j in 1 .. ny - 2 -> x1, lerp y1 y0 j ny ]
    @ [ for i in 0 .. nx - 1 -> lerp x1 x0 i nx, y0 ]
    @ [ for j in 1 .. ny - 2 -> x0, lerp y0 y1 j ny ]

/// Circular layout: n bars equally spaced on a circle of radius D/2 − edge, first bar at the top
let circLayout (diameter: float) (edge: float) (n: int) =
    if n < 3 then failwith "a circular layout needs at least 3 bars"
    let r = diameter / 2.0 - edge
    if r <= 0.0 then failwith "cover too large for the section"
    [ for k in 0 .. n - 1 ->
          let a = Math.PI / 2.0 - 2.0 * Math.PI * float k / float n
          r * cos a, r * sin a ]

// ---------------------------------------------------------------------------------------------
// Axis properties
// ---------------------------------------------------------------------------------------------

type AxisProps =
    { /// squash load α1 f'c Ac + Σ fsy As (Ac net of bars), kN
      Nuo: float
      /// pure tension Σ fsy As (negative), kN
      Nuot: float
      /// balanced point (outermost tension bar at yield), kN and kN·m
      Nub: float
      Mub: float
      /// pure bending, kN·m, its kuo and φ
      Muo: float
      Kuo: float
      PhiB: float
      /// extreme compression fibre to outermost tension bar, mm
      Do: float }

let axisProps (sec: Section) (c: Concrete) (thetaDeg: float) (duct: Design.Ductility) =
    let d = dirOf thetaDeg
    let bend = Ultimate.muo sec c thetaDeg
    let outer = sec.Bars |> List.minBy (fun b -> sOf d { X = b.X; Y = b.Y })
    let dnb = Design.dnBalanced bend.Do outer.Steel.Fsy outer.Steel.Es
    let bal = Ultimate.evaluate sec c thetaDeg dnb DefaultStrips
    { Nuo = Design.squash c.Fc (gross sec).A (sec.Bars |> List.map (fun b -> b.Area, b.Steel.Fsy)) / 1000.0
      Nuot = -(sec.Bars |> List.sumBy (fun b -> b.Area * b.Steel.Fsy)) / 1000.0
      Nub = bal.N / 1000.0
      Mub = bal.Mu / 1e6
      Muo = bend.Mu
      Kuo = bend.Kuo
      PhiB = Design.phiBending duct bend.Kuo
      Do = bend.Do }

// ---------------------------------------------------------------------------------------------
// Capacity at an axial force
// ---------------------------------------------------------------------------------------------

type Capacity =
    { /// direction of maximum compression, degrees
      Theta: float
      /// neutral axis depth, mm (NaN on the straight line to the squash load)
      Dn: float
      /// kN (compression +)
      N: float
      /// unfactored capacity: components and resultant, kN·m
      Mx: float
      My: float
      Mu: float
      Phi: float
      Nub: float
      Nuo: float }

/// Capacity at axial force nKN with maximum compression in direction θ.
/// Above the strain-compatibility point with dn = depth / γ, the capacity runs in a straight line
/// to the squash load Nuo (Cl 10.6.2.2), where M = 0.
let atThetaAx (sec: Section) (c: Concrete) (ax: AxisProps) (thetaDeg: float) (nKN: float) (phi0: float) (duct: Design.Ductility) =
    let nKN = if abs (nKN - ax.Nuo) < 1e-9 * ax.Nuo then ax.Nuo elif abs (nKN - ax.Nuot) < 1e-9 * abs ax.Nuot then ax.Nuot else nKN
    if nKN > ax.Nuo then failwithf "N* = %.0f kN exceeds the squash load Nuo = %.0f kN" nKN ax.Nuo
    if nKN < ax.Nuot then failwithf "tension N = %.0f kN exceeds the tension capacity %.0f kN" nKN ax.Nuot
    let phi =
        if nKN >= 0.0 then Design.phiCompression phi0 ax.PhiB nKN ax.Nub
        else Design.phiTension duct ax.PhiB nKN ax.Nuot
    let h = Ultimate.depthIn sec thetaDeg
    // near pure tension (dn → 0): straight line from the all-bars-yielded point to the smallest dn the solver uses
    let lo = Ultimate.evaluate sec c thetaDeg (1e-6 * h) DefaultStrips
    if nKN * 1000.0 <= lo.N then
        let cen = centroid sec
        let tx = -(sec.Bars |> List.sumBy (fun b -> b.Area * b.Steel.Fsy * (b.Y - cen.Y))) / 1e6
        let ty = -(sec.Bars |> List.sumBy (fun b -> b.Area * b.Steel.Fsy * (b.X - cen.X))) / 1e6
        let k = if lo.N / 1000.0 - ax.Nuot = 0.0 then 0.0 else (nKN - ax.Nuot) / (lo.N / 1000.0 - ax.Nuot)
        let mx = tx + k * (lo.Mx / 1e6 - tx)
        let my = ty + k * (lo.My / 1e6 - ty)
        { Theta = thetaDeg; Dn = nan; N = nKN; Mx = mx; My = my; Mu = sqrt (mx * mx + my * my)
          Phi = phi; Nub = ax.Nub; Nuo = ax.Nuo }
    else
    let dnLim = h / gamma c.Fc
    let lim = Ultimate.evaluate sec c thetaDeg dnLim DefaultStrips
    let nLim = lim.N / 1000.0
    if nKN > nLim then
        let k = (ax.Nuo - nKN) / (ax.Nuo - nLim)
        let mx, my = k * lim.Mx / 1e6, k * lim.My / 1e6
        { Theta = thetaDeg; Dn = nan; N = nKN; Mx = mx; My = my; Mu = sqrt (mx * mx + my * my)
          Phi = phi; Nub = ax.Nub; Nuo = ax.Nuo }
    else
        let dn = Ultimate.solveDn sec c thetaDeg (nKN * 1000.0) DefaultStrips
        let st = Ultimate.evaluate sec c thetaDeg dn DefaultStrips
        let mx, my = st.Mx / 1e6, st.My / 1e6
        { Theta = thetaDeg; Dn = dn; N = nKN; Mx = mx; My = my; Mu = sqrt (mx * mx + my * my)
          Phi = phi; Nub = ax.Nub; Nuo = ax.Nuo }

let atTheta (sec: Section) (c: Concrete) (thetaDeg: float) (nKN: float) (phi0: float) (duct: Design.Ductility) =
    atThetaAx sec c (axisProps sec c thetaDeg duct) thetaDeg nKN phi0 duct

/// M–N curve for bending direction θ on a fixed grid of 40 axial forces: 6 from pure tension to 0,
/// 14 from 0 to Nub, 20 from Nub to the squash load. M is about the axis normal to θ (Mx for 90°, My for 0°).
let mnCurve (sec: Section) (c: Concrete) (thetaDeg: float) (phi0: float) (duct: Design.Ductility) =
    let ax = axisProps sec c thetaDeg duct
    let d = dirOf thetaDeg
    let seg a b k incl = [ for i in 0 .. k - 1 -> a + (b - a) * float i / float (if incl then k - 1 else k) ]
    let ns = seg ax.Nuot 0.0 6 false @ seg 0.0 ax.Nub 14 false @ seg ax.Nub ax.Nuo 20 true
    [ for n in ns ->
          let r = atThetaAx sec c ax thetaDeg n phi0 duct
          n, r.Mx * d.Uy + r.My * d.Ux, r.Phi ]

let private wrap180 (a: float) =
    let r = (a + 180.0) % 360.0
    (if r < 0.0 then r + 360.0 else r) - 180.0

let private angleOf (mx: float) (my: float) = Math.Atan2(my, mx) * 180.0 / Math.PI

/// Biaxial capacity by strain compatibility: finds the neutral axis direction θ for which the capacity
/// moment (Mx, My) at N* points the same way as the applied moment (mx, my). The capacity is the
/// point on the moment contour at N* along the load direction (radial check, constant N).
let biaxial (sec: Section) (c: Concrete) (nKN: float) (mx: float) (my: float) (phi0: float) (duct: Design.Ductility) =
    if mx = 0.0 && my = 0.0 then atTheta sec c 90.0 nKN phi0 duct
    else
        let target = angleOf mx my
        let cap t = atTheta sec c t nKN phi0 duct
        let g t =
            let r = cap t
            wrap180 (angleOf r.Mx r.My - target)
        // ψ = atan2(My, Mx) is close to 90° − θ; bracket the root around that, widening as needed
        let t0 = 90.0 - target
        let g0 = g t0
        if g0 = 0.0 then cap t0
        else
            let mutable lo = nan
            let mutable hi = nan
            let mutable step = 5.0
            while Double.IsNaN lo && step <= 90.0 do
                for t in [ t0 - step; t0 + step ] do
                    if Double.IsNaN lo then
                        let gt = g t
                        if Math.Sign gt <> Math.Sign g0 && abs gt < 90.0 then
                            lo <- min t0 t
                            hi <- max t0 t
                step <- step + 5.0
            if Double.IsNaN lo then failwith "biaxial: no neutral axis direction found for this load direction"
            let glo = g lo
            let mutable i = 0
            while hi - lo > 1e-7 && i < 80 do
                let mid = 0.5 * (lo + hi)
                if Math.Sign(g mid) = Math.Sign glo then lo <- mid else hi <- mid
                i <- i + 1
            cap (0.5 * (lo + hi))

/// Moment capacity contour at axial force N: nPoints directions θ around the section (closed loop)
let contour (sec: Section) (c: Concrete) (nKN: float) (nPoints: int) (phi0: float) (duct: Design.Ductility) =
    let n = max 8 nPoints
    [ for k in 0 .. n -> atTheta sec c (90.0 - 360.0 * float k / float n) nKN phi0 duct ]
