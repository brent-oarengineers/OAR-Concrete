/// Custom section outlines for the column sheet: a polygon given corner by corner, where any edge can be a circular
/// arc (a rise / sagitta), plus a bar layout that follows the outline at a set distance from the face – so bars can
/// follow a curved face. Coordinates are mm, x right and y up, any origin.
module OAR.Concrete.Outline

open System
open OAR.Concrete.Geometry
open OAR.Concrete.Section

/// One corner and the edge from it to the next corner (the last corner joins back to the first).
/// Rise: 0 = straight edge; otherwise the edge is a circular arc whose mid-point sits Rise from the chord,
/// + bulging out of the section (convex), − into it (concave).
/// Bars: intermediate bars along that edge, between the corner bars.
type Corner =
    { X: float
      Y: float
      Rise: float
      Bars: int }

/// Chords per arc when an arc is turned into a polygon (area error under 0.01 % for a typical soffit)
let ArcSegments = 48

type private Edge =
    | Line of p: Pt * q: Pt
    /// centre, radius, start angle, sweep (radians, signed), convex (bulges out of the section)
    | Arc of c: Pt * r: float * a0: float * sweep: float * convex: bool

let private pt (x: float) (y: float) : Pt = { Pt.X = x; Pt.Y = y }
let private sub (a: Pt) (b: Pt) = (pt (a.X - b.X) (a.Y - b.Y))
let private add (a: Pt) (b: Pt) = (pt (a.X + b.X) (a.Y + b.Y))
let private mul (k: float) (a: Pt) = (pt (k * a.X) (k * a.Y))
let private len (a: Pt) = sqrt (a.X * a.X + a.Y * a.Y)
let private dot (a: Pt) (b: Pt) = a.X * b.X + a.Y * b.Y

let private wrap (a: float) =
    let mutable r = a
    while r > Math.PI do r <- r - 2.0 * Math.PI
    while r <= -Math.PI do r <- r + 2.0 * Math.PI
    r

let private check (cs: Corner list) =
    if cs.Length < 3 then failwith "a custom outline needs at least 3 corners"
    cs

/// +1 when the corners run anticlockwise, −1 when clockwise
let private orientation (cs: Corner list) =
    let pts = cs |> List.map (fun c -> (pt (c.X) (c.Y))) |> Array.ofList
    let n = pts.Length
    let a = Seq.sum (seq { for i in 0 .. n - 1 -> pts.[i].X * pts.[(i + 1) % n].Y - pts.[(i + 1) % n].X * pts.[i].Y })
    if abs a < 1e-9 then failwith "the outline has no area (corners in a line?)"
    float (sign a)

/// Outward unit normal of the chord p → q
let private outward (o: float) (p: Pt) (q: Pt) =
    let d = sub q p
    let l = len d
    if l < 1e-9 then failwith "two consecutive corners are at the same point"
    mul (o / l) (pt (d.Y) (-d.X))

let private edges (cs: Corner list) =
    let cs = check cs
    let o = orientation cs
    let arr = Array.ofList cs
    let n = arr.Length
    [ for i in 0 .. n - 1 ->
          let p = (pt (arr.[i].X) (arr.[i].Y))
          let q = (pt (arr.[(i + 1) % n].X) (arr.[(i + 1) % n].Y))
          let rise = arr.[i].Rise
          if abs rise < 1e-9 then Line(p, q)
          else
              let nOut = outward o p q
              let chord = len (sub q p)
              let s = abs rise
              if s >= chord / 2.0 * 0.999 then failwithf "corner %d: the rise %.0f must be less than half the chord (%.0f)" (i + 1) rise (chord / 2.0)
              let r = (chord * chord / 4.0 + s * s) / (2.0 * s)
              let mid = mul 0.5 (add p q)
              let bulge = add mid (mul rise nOut)
              // centre on the far side of the chord from the bulge
              let c = sub bulge (mul (float (sign rise) * r) nOut)
              let ang (a: Pt) = Math.Atan2(a.Y - c.Y, a.X - c.X)
              let a0 = ang p
              let sweep = wrap (ang bulge - a0) + wrap (ang q - ang bulge)
              Arc(c, r, a0, sweep, rise > 0.0) ]

let private at (c: Pt) (r: float) (a: float) = (pt (c.X + r * cos a) (c.Y + r * sin a))

/// The outline as polygon vertices (arcs as ArcSegments chords), in the corners' coordinates
let polygon (cs: Corner list) : Pt list =
    [ for e in edges cs do
          match e with
          | Line(p, _) -> yield p
          | Arc(c, r, a0, sw, _) ->
              for k in 0 .. ArcSegments - 1 -> at c r (a0 + sw * float k / float ArcSegments) ]

/// Outward normal at the start and at the end of an edge
let private normals (o: float) (e: Edge) =
    match e with
    | Line(p, q) ->
        let n = outward o p q
        n, n
    | Arc(_, _, a0, sw, convex) ->
        // radial direction, pointing out of the section: away from the centre when convex
        let k = if convex then 1.0 else -1.0
        let rad (a: float) = (pt (k * cos a) (k * sin a))
        rad a0, rad (a0 + sw)

/// Bars following the outline at `edge` from the faces (cover + ligature + db/2): one at each corner (on the
/// bisector of the two faces) and each edge's intermediate bars equally spaced along that edge (along the arc for
/// a curved edge). Returns (x, y) in the corners' coordinates, corner bar then that edge's bars, corner by corner.
let layout (cs: Corner list) (edge: float) : (float * float) list =
    let es = Array.ofList (edges cs)
    let o = orientation cs
    let arr = Array.ofList cs
    let n = es.Length
    let corner i =
        let v = (pt (arr.[i].X) (arr.[i].Y))
        let _, nIn = normals o es.[(i + n - 1) % n] // end of the edge arriving
        let nOut, _ = normals o es.[i] // start of the edge leaving
        let k = 1.0 + dot nIn nOut
        if k < 1e-6 then failwithf "corner %d folds back on itself" (i + 1)
        sub v (mul (edge / k) (add nIn nOut))
    let corners = Array.init n corner
    [ for i in 0 .. n - 1 do
          let a, b = corners.[i], corners.[(i + 1) % n]
          yield a.X, a.Y
          let m = max 0 arr.[i].Bars
          match es.[i] with
          | Line _ ->
              for j in 1 .. m do
                  let t = float j / float (m + 1)
                  yield a.X + t * (b.X - a.X), a.Y + t * (b.Y - a.Y)
          | Arc(c, r, a0, sw, convex) ->
              let r' = if convex then r - edge else r + edge
              if r' <= 0.0 then failwithf "corner %d: the arc is too tight for the bars" (i + 1)
              let angA, angB = Math.Atan2(a.Y - c.Y, a.X - c.X), Math.Atan2(b.Y - c.Y, b.X - c.X)
              // go the same way round as the edge
              let start = a0 + wrap (angA - a0)
              let span =
                  let d = wrap (angB - angA)
                  if sign d <> sign sw && abs d > 1e-9 then d + float (sign sw) * 2.0 * Math.PI else d
              for j in 1 .. m do
                  let p = at c r' (start + span * float j / float (m + 1))
                  yield p.X, p.Y ]

/// Gross properties of a polygon: area, centroid, second moments about centroidal x and y axes (mm, mm², mm⁴)
type Props =
    { A: float
      Xc: float
      Yc: float
      Ix: float
      Iy: float
      XMin: float
      XMax: float
      YMin: float
      YMax: float }

let props (pts: Pt list) =
    let p = Array.ofList pts
    let n = p.Length
    let mutable a, sx, sy, ixx, iyy = 0.0, 0.0, 0.0, 0.0, 0.0
    for i in 0 .. n - 1 do
        let p0, p1 = p.[i], p.[(i + 1) % n]
        let cr = p0.X * p1.Y - p1.X * p0.Y
        a <- a + cr
        sx <- sx + (p0.X + p1.X) * cr
        sy <- sy + (p0.Y + p1.Y) * cr
        ixx <- ixx + (p0.Y * p0.Y + p0.Y * p1.Y + p1.Y * p1.Y) * cr
        iyy <- iyy + (p0.X * p0.X + p0.X * p1.X + p1.X * p1.X) * cr
    let s = float (sign a)
    let area = s * a / 2.0
    let xc, yc = sx / (3.0 * a), sy / (3.0 * a)
    { A = area
      Xc = xc
      Yc = yc
      Ix = s * ixx / 12.0 - area * yc * yc
      Iy = s * iyy / 12.0 - area * xc * xc
      XMin = p |> Array.map (fun q -> q.X) |> Array.min
      XMax = p |> Array.map (fun q -> q.X) |> Array.max
      YMin = p |> Array.map (fun q -> q.Y) |> Array.min
      YMax = p |> Array.map (fun q -> q.Y) |> Array.max }

/// The outline as a Section, shifted so the top-left of its bounding box is (0, 0) with y up – the same frame as a
/// rectangle (bar depth from the top, x from the left)
let section (cs: Corner list) (bars: Bar list) : Section =
    let pts = polygon cs
    let pr = props pts
    let shifted = pts |> List.map (fun p -> (pt (p.X - pr.XMin) (p.Y - pr.YMax))) |> Array.ofList
    { Regions = [ { Shape = Polygon shifted; IsVoid = false } ]; Bars = bars }
