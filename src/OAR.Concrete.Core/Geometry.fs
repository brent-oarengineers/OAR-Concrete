/// Section geometry: concrete regions (polygons, circles, voids) and the area integrals the
/// analysis needs. Everything is in mm, global axes x (right) and y (up).
///
/// The analysis works with a direction u (unit vector pointing to the most compressed side) and the
/// coordinate s = p · u of any point p. Lines of equal s are parallel to the neutral axis. The concrete
/// is cut into thin bands between two values of s; for each band we need its area and centroid.
module OAR.Concrete.Geometry

open System

type Pt = { X: float; Y: float }

type Shape =
    /// Vertices in order around the boundary (either direction)
    | Polygon of Pt[]
    /// Centre and radius
    | Circle of Pt * float

type Region = { Shape: Shape; IsVoid: bool }

/// Direction of maximum compression
type Dir = { Ux: float; Uy: float }

/// θ in degrees, measured from +x: 90 = top fibre in compression (sagging), −90 = bottom (hogging)
let dirOf (thetaDeg: float) =
    let t = thetaDeg * Math.PI / 180.0
    let ux, uy = cos t, sin t
    // tidy round-off so 90° gives exactly (0, 1)
    let tidy v = if abs v < 1e-14 then 0.0 else v
    { Ux = tidy ux; Uy = tidy uy }

let sOf (d: Dir) (p: Pt) = p.X * d.Ux + p.Y * d.Uy

/// Area and first moments of an area: A, ∫x dA, ∫y dA
type Moments = { A: float; Sx: float; Sy: float }

let zeroM = { A = 0.0; Sx = 0.0; Sy = 0.0 }
let addM (a: Moments) (b: Moments) = { A = a.A + b.A; Sx = a.Sx + b.Sx; Sy = a.Sy + b.Sy }
let scaleM k (a: Moments) = { A = k * a.A; Sx = k * a.Sx; Sy = k * a.Sy }

/// Shoelace formulas; returns positive area whatever the vertex order
let polygonMoments (pts: Pt[]) =
    let n = pts.Length
    if n < 3 then zeroM
    else
        let mutable a = 0.0
        let mutable sx = 0.0
        let mutable sy = 0.0
        for i in 0 .. n - 1 do
            let p = pts.[i]
            let q = pts.[(i + 1) % n]
            let cross = p.X * q.Y - q.X * p.Y
            a <- a + cross
            sx <- sx + (p.X + q.X) * cross
            sy <- sy + (p.Y + q.Y) * cross
        let m = { A = a / 2.0; Sx = sx / 6.0; Sy = sy / 6.0 }
        if m.A < 0.0 then scaleM -1.0 m else m

/// Sutherland–Hodgman clip of a polygon to the half-plane f(p) ≥ 0 (f linear)
let private clipHalf (pts: Pt[]) (f: Pt -> float) =
    let n = pts.Length
    if n = 0 then pts
    else
        let out = Collections.Generic.List<Pt>(n + 2)
        for i in 0 .. n - 1 do
            let cur = pts.[i]
            let prev = pts.[(i + n - 1) % n]
            let fc = f cur
            let fp = f prev
            let cut () =
                let t = fp / (fp - fc)
                { X = prev.X + t * (cur.X - prev.X); Y = prev.Y + t * (cur.Y - prev.Y) }
            if fc >= 0.0 then
                if fp < 0.0 then out.Add(cut ())
                out.Add cur
            elif fp >= 0.0 then
                out.Add(cut ())
        out.ToArray()

/// Part of a polygon with s1 ≤ s ≤ s2
let private polygonBand (pts: Pt[]) (d: Dir) (s1: float) (s2: float) =
    let a = clipHalf pts (fun p -> sOf d p - s1)
    let b = clipHalf a (fun p -> s2 - sOf d p)
    polygonMoments b

/// Circle of radius r: area with (s − sc) ≤ h, and its first moment about the centre line
let private circleBelow (r: float) (h: float) =
    let h = max -r (min r h)
    let root = sqrt (max 0.0 (r * r - h * h))
    let area = r * r * Math.Acos(-h / r) + h * root
    let moment = -(2.0 / 3.0) * root * root * root
    area, moment

/// Part of a circle with s1 ≤ s ≤ s2: exact (circular segment formulas, Brent's hand derivation)
let private circleBand (c: Pt) (r: float) (d: Dir) (s1: float) (s2: float) =
    let sc = sOf d c
    // coordinate across the band direction: w = p · v, v = u rotated +90°
    let vx, vy = -d.Uy, d.Ux
    let wc = c.X * vx + c.Y * vy
    let a2, q2 = circleBelow r (s2 - sc)
    let a1, q1 = circleBelow r (s1 - sc)
    let a = a2 - a1
    if a <= 0.0 then zeroM
    else
        let ss = sc * a + (q2 - q1) // ∫ s dA
        let sw = wc * a // ∫ w dA (symmetric about the centre)
        { A = a; Sx = d.Ux * ss + vx * sw; Sy = d.Uy * ss + vy * sw }

/// Signed moments (voids negative) of the band s1 ≤ s ≤ s2 of a region
let band (reg: Region) (d: Dir) (s1: float) (s2: float) =
    let m =
        match reg.Shape with
        | Polygon pts -> polygonBand pts d s1 s2
        | Circle(c, r) -> circleBand c r d s1 s2
    if reg.IsVoid then scaleM -1.0 m else m

/// Signed moments of the whole region
let whole (reg: Region) =
    let m =
        match reg.Shape with
        | Polygon pts -> polygonMoments pts
        | Circle(c, r) ->
            let a = Math.PI * r * r
            { A = a; Sx = c.X * a; Sy = c.Y * a }
    if reg.IsVoid then scaleM -1.0 m else m

/// Range of s covered by a region
let extent (reg: Region) (d: Dir) =
    match reg.Shape with
    | Polygon pts ->
        let ss = pts |> Array.map (sOf d)
        Array.min ss, Array.max ss
    | Circle(c, r) ->
        let sc = sOf d c
        sc - r, sc + r

/// Is a point inside the region's outline
let contains (reg: Region) (p: Pt) =
    match reg.Shape with
    | Circle(c, r) -> (p.X - c.X) ** 2.0 + (p.Y - c.Y) ** 2.0 <= r * r * (1.0 + 1e-9)
    | Polygon pts ->
        // crossing-number test; points on the boundary count as inside (bars at cover are well inside anyway)
        let n = pts.Length
        let mutable inside = false
        for i in 0 .. n - 1 do
            let a = pts.[i]
            let b = pts.[(i + 1) % n]
            if (a.Y > p.Y) <> (b.Y > p.Y) then
                let xCross = a.X + (p.Y - a.Y) * (b.X - a.X) / (b.Y - a.Y)
                if p.X < xCross then inside <- not inside
        inside
