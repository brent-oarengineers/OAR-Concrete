/// A reinforced concrete section: concrete regions plus individual bars, and the strip integration
/// shared by the ultimate (strain compatibility) and cracked elastic analyses.
///
/// The same section type serves slabs, beams and (later) rectangular or circular columns: a slab strip
/// is a b × D rectangle with one bar per layer carrying that layer's total area.
module OAR.Concrete.Section

open System
open OAR.Concrete.Materials
open OAR.Concrete.Geometry

type Bar =
    { X: float
      Y: float
      /// mm²
      Area: float
      Steel: Steel }

type Section =
    { Regions: Region list
      Bars: Bar list }

// ---------------------------------------------------------------------------------------------
// Builders. Sheet convention: x from the left edge, depth measured DOWN from the top face.
// Internally y = −depth, so the top face is y = 0.
// ---------------------------------------------------------------------------------------------

/// Bar from sheet values: area mm², depth from the top face mm, fsy MPa, x from the left edge mm
let barAt (area: float) (depth: float) (fsy: float) (x: float) (es: float) =
    { X = x; Y = -depth; Area = area; Steel = { Fsy = fsy; Es = es } }

/// b × D rectangle, top-left corner at (0, 0)
let rectangle (b: float) (depth: float) (bars: Bar list) =
    if b <= 0.0 || depth <= 0.0 then failwith "b and D must be positive"
    let pts = [| { X = 0.0; Y = 0.0 }; { X = b; Y = 0.0 }; { X = b; Y = -depth }; { X = 0.0; Y = -depth } |]
    { Regions = [ { Shape = Polygon pts; IsVoid = false } ]; Bars = bars }

/// Circle of diameter D, its bounding square's top-left corner at (0, 0)
let circular (diameter: float) (bars: Bar list) =
    if diameter <= 0.0 then failwith "D must be positive"
    let r = diameter / 2.0
    { Regions = [ { Shape = Circle({ X = r; Y = -r }, r); IsVoid = false } ]; Bars = bars }

/// n bars evenly on a circle (pitch circle diameter pcd), first bar at the top, bounding-square coordinates
let circularBars (diameter: float) (pcd: float) (n: int) (areaEach: float) (fsy: float) =
    let r = diameter / 2.0
    [ for k in 0 .. n - 1 do
          let a = Math.PI / 2.0 + 2.0 * Math.PI * float k / float n
          yield { X = r + pcd / 2.0 * cos a; Y = -r + pcd / 2.0 * sin a; Area = areaEach; Steel = steel fsy } ]

// ---------------------------------------------------------------------------------------------
// Section properties
// ---------------------------------------------------------------------------------------------

/// Gross concrete area and first moments (voids subtracted, bars ignored)
let gross (sec: Section) = sec.Regions |> List.fold (fun acc r -> addM acc (whole r)) zeroM

/// Centroid of the gross concrete section: moments are taken about this point
let centroid (sec: Section) =
    let m = gross sec
    if m.A <= 0.0 then failwith "section has no concrete area"
    { X = m.Sx / m.A; Y = m.Sy / m.A }

/// Range of s over the solid concrete
let extent (sec: Section) (d: Dir) =
    let solids = sec.Regions |> List.filter (fun r -> not r.IsVoid)
    if solids.IsEmpty then failwith "section has no concrete region"
    let es = solids |> List.map (fun r -> Geometry.extent r d)
    (es |> List.map fst |> List.min), (es |> List.map snd |> List.max)

/// Is the point in concrete (inside a solid region and not in a void)
let inConcrete (sec: Section) (p: Pt) =
    let inSolid = sec.Regions |> List.exists (fun r -> not r.IsVoid && contains r p)
    let inVoid = sec.Regions |> List.exists (fun r -> r.IsVoid && contains r p)
    inSolid && not inVoid

/// Total steel area, mm²
let steelArea (sec: Section) = sec.Bars |> List.sumBy (fun b -> b.Area)

// ---------------------------------------------------------------------------------------------
// Strip integration
// ---------------------------------------------------------------------------------------------

/// One band of concrete: thickness h (in s), signed moments, s of its centroid
type Strip = { H: float; M: Moments; S: float }

/// Default number of strips over the compression zone (accuracy better than 0.01 % on the tests)
[<Literal>]
let DefaultStrips = 200

/// Cut the concrete between s1 and s2 into about n strips, with a strip boundary at each value in
/// `breaks` (where a stress-strain law has a step or kink, e.g. the edge of the rectangular block)
let strips (sec: Section) (d: Dir) (s1: float) (s2: float) (breaks: float list) (n: int) =
    if s2 <= s1 then [||]
    else
        let cuts =
            (s1 :: s2 :: (breaks |> List.filter (fun b -> b > s1 && b < s2)))
            |> List.sort
            |> List.toArray
        let total = s2 - s1
        [| for i in 0 .. cuts.Length - 2 do
               let a, b = cuts.[i], cuts.[i + 1]
               let len = b - a
               if len > total * 1e-12 then
                   let k = max 2 (int (Math.Ceiling(float n * len / total)))
                   for j in 0 .. k - 1 do
                       let lo = a + len * float j / float k
                       let hi = a + len * float (j + 1) / float k
                       let m = sec.Regions |> List.fold (fun acc r -> addM acc (band r d lo hi)) zeroM
                       if abs m.A > 1e-12 then
                           let s = (m.Sx * d.Ux + m.Sy * d.Uy) / m.A
                           yield { H = hi - lo; M = m; S = s } |]
