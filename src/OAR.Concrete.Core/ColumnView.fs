/// Column graphics: the pictures behind OAR > Column graphics (section and bars, M–N interaction about each axis,
/// moment contour at the N* of each load case). Everything here is plain data (lines, polygons and text in pixel
/// coordinates, y down) so it can be tested without Excel; the add-in window only paints it.
/// Same drawing model as OAR Steel / OAR Timber (Review module). Units: mm, kN, kN·m.
module OAR.Concrete.ColumnView

open System
open OAR.Concrete.Materials
open OAR.Concrete.Geometry
open OAR.Concrete.Section

type Align =
    | Left
    | Centre
    | Right

/// A drawing instruction. Coordinates in pixels, y down. Colours are "#RRGGBB".
type Prim =
    | PLine of x1: float * y1: float * x2: float * y2: float * colour: string * width: float * dashed: bool
    /// stroke "" = no outline
    | PPolygon of pts: (float * float) list * fill: string * opacity: float * stroke: string * width: float
    | PPolyline of pts: (float * float) list * colour: string * width: float
    /// y is the vertical centre of the text
    | PText of x: float * y: float * text: string * colour: string * size: float * align: Align * bold: bool

type Scene = { Width: float; Height: float; Prims: Prim list }

/// One ULS load case as the sheet checks it: N* and the design moments (magnified, minimum eccentricity applied)
type ColCase =
    { Name: string
      Description: string
      N: float
      Mx: float
      My: float }

type ColumnData =
    { Title: string
      Section: Section
      Concrete: Concrete
      Duct: Design.Ductility
      Phi0: float
      /// bar names in the order of Section.Bars (may be shorter, "" = unknown)
      BarNames: string list
      Cases: ColCase list }

type Item =
    | SectionItem
    /// 'x' = bending about the x axis (Mx), 'y' = about the y axis (My)
    | MNItem of axis: char
    /// moment contour at the N* of case i (0-based)
    | ContourItem of case: int

type View =
    { Unfactored: bool
      Phi: bool
      Loads: bool
      KeyPoints: bool }

let defaultView = { Unfactored = true; Phi = true; Loads = true; KeyPoints = true }

/// The list on the left of the window: (group, label, tooltip, item)
let items (d: ColumnData) =
    [ yield "Section", "Section and bars", "The section to scale with every bar (typical and custom)", SectionItem
      yield "Interaction (M–N)", "About x (Mx)", "Uniaxial M–N curve about the x axis, both directions, with the design actions", MNItem 'x'
      yield "Interaction (M–N)", "About y (My)", "Uniaxial M–N curve about the y axis, both directions, with the design actions", MNItem 'y'
      for i, c in List.indexed d.Cases ->
          "Moment contour at N*", sprintf "%s  (N* %.0f kN)" c.Name c.N, (if c.Description = "" then c.Name else c.Description), ContourItem i ]

// ---------------------------------------------------------------------------------------------
// Drawing helpers
// ---------------------------------------------------------------------------------------------

let private ink = "#222222"
let private muted = "#777777"
let private gridC = "#E4E4E4"
let private axisC = "#9A9A9A"
let blue = "#4472C4"
let orange = "#ED7D31"
let red = "#C0392B"
let private green = "#465A52"

let fmt (v: float) =
    let a = abs v
    if a < 1e-9 then "0"
    elif a >= 100.0 then sprintf "%.0f" v
    elif a >= 10.0 then sprintf "%.1f" v
    else sprintf "%.2f" v

/// A tidy tick step (1, 2, 2.5, 5 × 10^k) giving about n ticks over the range
let niceStep (range: float) (n: int) =
    if range <= 0.0 || Double.IsNaN range then 1.0
    else
        let raw = range / float (max 1 n)
        let p = 10.0 ** floor (log10 raw)
        [ 1.0; 2.0; 2.5; 5.0; 10.0 ] |> List.map (fun k -> k * p) |> List.find (fun s -> s >= raw * 0.999)

let private circlePts (cx: float) (cy: float) (r: float) (n: int) =
    [ for k in 0 .. n - 1 ->
          let a = 2.0 * Math.PI * float k / float n
          cx + r * cos a, cy + r * sin a ]

/// Plot frame: data bounds to a pixel box (optionally equal scale in x and y)
type private Frame =
    { X0: float; X1: float; Y0: float; Y1: float
      L: float; T: float; W: float; H: float }
    member f.Px(x: float) = f.L + (x - f.X0) / (f.X1 - f.X0) * f.W
    member f.Py(y: float) = f.T + f.H - (y - f.Y0) / (f.Y1 - f.Y0) * f.H

let private frame (xmin: float) (xmax: float) (ymin: float) (ymax: float) (l: float) (t: float) (w: float) (h: float) (equal: bool) =
    let pad (a: float) (b: float) = let d = max (b - a) 1e-9 in a - 0.06 * d, b + 0.06 * d
    let x0, x1 = pad xmin xmax
    let y0, y1 = pad ymin ymax
    if not equal then { X0 = x0; X1 = x1; Y0 = y0; Y1 = y1; L = l; T = t; W = w; H = h }
    else
        // same units per pixel both ways, centred in the box
        let k = max ((x1 - x0) / w) ((y1 - y0) / h)
        let cx, cy = (x0 + x1) / 2.0, (y0 + y1) / 2.0
        { X0 = cx - k * w / 2.0; X1 = cx + k * w / 2.0; Y0 = cy - k * h / 2.0; Y1 = cy + k * h / 2.0; L = l; T = t; W = w; H = h }

/// Gridlines, zero axes, tick labels and axis titles
let private axes (f: Frame) (xTitle: string) (yTitle: string) (equal: bool) =
    let sx = niceStep (f.X1 - f.X0) (max 4 (int (f.W / 90.0)))
    let sy = if equal then sx else niceStep (f.Y1 - f.Y0) (max 4 (int (f.H / 60.0)))
    [ for k in int (ceil (f.X0 / sx)) .. int (floor (f.X1 / sx)) do
          let x = float k * sx
          let px = f.Px x
          yield PLine(px, f.T, px, f.T + f.H, (if k = 0 then axisC else gridC), (if k = 0 then 1.2 else 1.0), false)
          yield PText(px, f.T + f.H + 11.0, fmt x, muted, 8.0, Centre, false)
      for k in int (ceil (f.Y0 / sy)) .. int (floor (f.Y1 / sy)) do
          let y = float k * sy
          let py = f.Py y
          yield PLine(f.L, py, f.L + f.W, py, (if k = 0 then axisC else gridC), (if k = 0 then 1.2 else 1.0), false)
          yield PText(f.L - 6.0, py, fmt y, muted, 8.0, Right, false)
      yield PPolygon([ f.L, f.T; f.L + f.W, f.T; f.L + f.W, f.T + f.H; f.L, f.T + f.H ], "#FFFFFF", 0.0, "#BDBDBD", 1.0)
      yield PText(f.L + f.W / 2.0, f.T + f.H + 28.0, xTitle, ink, 9.0, Centre, true)
      yield PText(f.L, f.T - 12.0, yTitle, ink, 9.0, Left, true) ]

let private title (w: float) (main: string) (sub: string) =
    [ PText(16.0, 16.0, main, ink, 12.0, Left, true)
      PText(16.0, 34.0, sub, muted, 8.5, Left, false)
      PLine(16.0, 46.0, w - 16.0, 46.0, green, 1.5, false) ]

/// Legend entries: (colour, line / dot, text) in a row under the title
let private legend (x: float) (y: float) (entries: (string * bool * string) list) =
    let mutable cx = x
    [ for colour, dot, t in entries do
          if dot then yield PPolygon(circlePts (cx + 8.0) y 4.0 16, colour, 1.0, "", 0.0)
          else yield PLine(cx, y, cx + 18.0, y, colour, 2.2, false)
          yield PText(cx + 24.0, y, t, ink, 8.5, Left, false)
          cx <- cx + 34.0 + float t.Length * 5.6 ]

let private barDia (area: float) = sqrt (4.0 * area / Math.PI)

/// "6 N24 + 1 TB36", from the bar names (blank names by size)
let barSummary (d: ColumnData) =
    let names =
        d.Section.Bars
        |> List.mapi (fun i b ->
            match List.tryItem i d.BarNames with
            | Some n when n.Trim() <> "" -> n.Trim()
            | _ -> sprintf "Ø%.0f" (barDia b.Area))
    names
    |> List.countBy id
    |> List.map (fun (n, k) -> sprintf "%d%s" k n)
    |> String.concat " + "

let private sub (d: ColumnData) =
    let a = d.Section.Bars |> List.sumBy (fun b -> b.Area)
    let ag = (gross d.Section).A
    let model =
        match d.Concrete.Model with
        | Rectangular -> "rectangular"
        | m -> (sprintf "%A" m).ToLowerInvariant()
    sprintf "f'c %.0f MPa  ·  %s, As = %.0f mm² (ρ = %.2f %%)  ·  φ0 = %.2f  ·  %s stress block" d.Concrete.Fc (barSummary d) a (100.0 * a / ag) d.Phi0 model

let private mnTheta (axis: char) = if axis = 'x' then 90.0, -90.0 else 0.0, 180.0

/// M of a case about the plot's axis (signed: + compression on the top face for x, on the +x face for y)
let private caseM (axis: char) (c: ColCase) = if axis = 'x' then c.Mx else c.My

// ---------------------------------------------------------------------------------------------
// Section
// ---------------------------------------------------------------------------------------------

let private sectionScene (d: ColumnData) (width: float) (height: float) =
    let g = centroid d.Section
    let outline =
        d.Section.Regions
        |> List.filter (fun r -> not r.IsVoid)
        |> List.map (fun r ->
            match r.Shape with
            | Polygon pts -> [ for p in pts -> p.X - g.X, p.Y - g.Y ]
            | Circle(c, r) -> circlePts (c.X - g.X) (c.Y - g.Y) r 96)
    let xs = outline |> List.concat |> List.map fst
    let ys = outline |> List.concat |> List.map snd
    let xmin, xmax, ymin, ymax = List.min xs, List.max xs, List.min ys, List.max ys
    let top, bottom = 84.0, 60.0
    let f = frame xmin xmax ymin ymax 90.0 top (width - 240.0) (height - top - bottom) true
    let pts (l: (float * float) list) = l |> List.map (fun (x, y) -> f.Px x, f.Py y)
    let k = f.W / (f.X1 - f.X0) // px per mm
    let isCircle = d.Section.Regions |> List.exists (fun r -> match r.Shape with Circle _ -> true | _ -> false)
    let ax = (xmax - xmin) / 2.0 + 24.0 / k
    let ay = (ymax - ymin) / 2.0 + 18.0 / k
    let names =
        d.Section.Bars |> List.mapi (fun i b ->
            match List.tryItem i d.BarNames with
            | Some n when n.Trim() <> "" -> n.Trim()
            | _ -> sprintf "Ø%.0f" (barDia b.Area))
    let kinds = names |> List.distinct
    let palette = [| "#2B2B2B"; "#C0392B"; "#4472C4"; "#ED7D31"; "#5B8C3A"; "#7B4F9E" |]
    let colourOf (n: string) = palette.[(List.findIndex ((=) n) kinds) % palette.Length]
    [ yield! title width "Section and reinforcement" (sub d)
      for o in outline do
          yield PPolygon(pts o, "#D9D9D9", 1.0, "#555555", 1.4)
      // axes through the centroid
      yield PLine(f.Px(-ax), f.Py 0.0, f.Px ax, f.Py 0.0, axisC, 1.0, true)
      yield PLine(f.Px 0.0, f.Py(-ay), f.Px 0.0, f.Py ay, axisC, 1.0, true)
      yield PText(f.Px ax + 6.0, f.Py 0.0, "x", ink, 10.0, Left, true)
      yield PText(f.Px 0.0 + 8.0, f.Py ay + 2.0, "y", ink, 10.0, Left, true)
      // bars to scale
      for b, n in List.zip d.Section.Bars names do
          let r = max 2.0 (barDia b.Area / 2.0 * k)
          yield PPolygon(circlePts (f.Px(b.X - g.X)) (f.Py(b.Y - g.Y)) r 28, colourOf n, 1.0, "#000000", 0.8)
      // dimensions
      if isCircle then
          yield PText(f.Px 0.0, f.Py ymin + 18.0, sprintf "Ø%.0f" (xmax - xmin), ink, 9.5, Centre, true)
      else
          let yb = f.Py ymin + 22.0
          yield PLine(f.Px xmin, yb, f.Px xmax, yb, muted, 1.0, false)
          yield PLine(f.Px xmin, yb - 5.0, f.Px xmin, yb + 5.0, muted, 1.0, false)
          yield PLine(f.Px xmax, yb - 5.0, f.Px xmax, yb + 5.0, muted, 1.0, false)
          yield PText(f.Px 0.0, yb + 11.0, sprintf "b = %.0f" (xmax - xmin), ink, 9.0, Centre, false)
          let xr = f.Px ax + 26.0
          yield PLine(xr, f.Py ymax, xr, f.Py ymin, muted, 1.0, false)
          yield PLine(xr - 5.0, f.Py ymax, xr + 5.0, f.Py ymax, muted, 1.0, false)
          yield PLine(xr - 5.0, f.Py ymin, xr + 5.0, f.Py ymin, muted, 1.0, false)
          yield PText(xr + 8.0, f.Py 0.0, sprintf "D = %.0f" (ymax - ymin), ink, 9.0, Left, false)
      // key on the right
      let kx = width - 104.0
      yield PText(kx, top + 4.0, "Bars", ink, 9.0, Left, true)
      for j, (n, cnt) in List.indexed (List.countBy id names) do
          let yy = top + 22.0 + 17.0 * float j
          yield PPolygon(circlePts (kx + 5.0) yy 5.0 16, colourOf n, 1.0, "#000000", 0.8)
          yield PText(kx + 16.0, yy, sprintf "%d %s" cnt n, ink, 9.0, Left, false)
      yield PText(16.0, height - 30.0, "+Mx: compression on the top face (+y)   ·   +My: compression on the +x face (right)   ·   bars drawn to scale (Ø from As)", muted, 8.0, Left, false) ]

// ---------------------------------------------------------------------------------------------
// M–N interaction (uniaxial, both directions)
// ---------------------------------------------------------------------------------------------

let private mnScene (d: ColumnData) (axis: char) (view: View) (width: float) (height: float) =
    let tp, tn = mnTheta axis
    let sec, c, duct = d.Section, d.Concrete, d.Duct
    let pos = Column.mnCurve sec c tp d.Phi0 duct
    let neg = Column.mnCurve sec c tn d.Phi0 duct
    // one line from the squash load round the negative side, through pure tension, up the positive side
    let line (scale: bool) =
        [ for n, m, phi in List.rev neg -> (if scale then -phi * m else -m), (if scale then phi * n else n)
          for n, m, phi in pos -> (if scale then phi * m else m), (if scale then phi * n else n) ]
    let unf, fac = line false, line true
    let loads = d.Cases |> List.map (fun cs -> cs, caseM axis cs)
    let allPts =
        (if view.Unfactored then unf else []) @ fac @ (if view.Loads then loads |> List.map (fun (cs, m) -> m, cs.N) else [])
    let xs, ys = allPts |> List.map fst, allPts |> List.map snd
    let mlim = max (abs (List.min xs)) (abs (List.max xs))
    let top, bottom = 92.0, 58.0
    let f = frame (-mlim) mlim (List.min ys) (List.max ys) 74.0 top (width - 100.0) (height - top - bottom) false
    let px (x, y) = f.Px x, f.Py y
    let mName = if axis = 'x' then "Mx" else "My"
    [ yield! title width (sprintf "M–N interaction – bending about %c (%s)" axis mName) (sub d)
      yield!
          legend 16.0 62.0
              [ if view.Unfactored then yield blue, false, "unfactored (Nu, Mu)"
                if view.Phi then yield orange, false, "design (φNu, φMu)"
                if view.Loads then yield red, true, "design actions (N*, M*d)" ]
      yield! axes f (sprintf "%s (kNm)   + compression on the %s" mName (if axis = 'x' then "top face" else "+x face")) "N (kN)  compression +" false
      if view.Unfactored then yield PPolyline(List.map px unf, blue, 1.6)
      if view.Phi then yield PPolyline(List.map px fac, orange, 2.2)
      if view.KeyPoints && view.Phi then
          for th, sgn in [ tp, 1.0; tn, -1.0 ] do
              let a = Column.axisProps sec c th duct
              let phiAt n = if n >= 0.0 then Design.phiCompression d.Phi0 a.PhiB n a.Nub else a.PhiT
              let pts =
                  [ "squash", 0.0, d.Phi0 * a.Nuo
                    "decompression", sgn * phiAt a.Nud * a.Mud, phiAt a.Nud * a.Nud
                    "balanced", sgn * d.Phi0 * a.Mub, d.Phi0 * a.Nub
                    "pure bending", sgn * a.PhiB * a.Muo, 0.0
                    "pure tension", 0.0, a.PhiT * a.Nuot ]
              for name, m, n in pts do
                  if not (sgn < 0.0 && (name = "squash" || name = "pure tension")) then
                      let x, y = px (m, n)
                      yield PPolygon([ x - 4.0, y; x, y - 4.0; x + 4.0, y; x, y + 4.0 ], ink, 1.0, "", 0.0)
                      if sgn > 0.0 then
                          yield PText(x + 8.0, y + 9.0, sprintf "%s  (%s kN, %s kNm)" name (fmt n) (fmt (abs m)), muted, 7.5, Left, false)
      if view.Loads then
          for cs, m in loads do
              let x, y = px (m, cs.N)
              yield PPolygon(circlePts x y 4.5 16, red, 1.0, "#FFFFFF", 0.8)
              yield PText(x + 7.0, y + 1.0, cs.Name, red, 8.0, Left, true)
      yield PText(16.0, height - 14.0, "Uniaxial curves (strain compatibility, OAR.CONCRETE.MNCOLUMN). Biaxial cases are checked on each case's moment contour.", muted, 8.0, Left, false) ]

// ---------------------------------------------------------------------------------------------
// Moment contour at N*
// ---------------------------------------------------------------------------------------------

let private contourScene (d: ColumnData) (i: int) (view: View) (width: float) (height: float) =
    let cs = d.Cases.[i]
    let sec, c, duct = d.Section, d.Concrete, d.Duct
    let ring = Column.contour sec c cs.N 72 d.Phi0 duct
    let unf = ring |> List.map (fun r -> r.Mx, r.My)
    let fac = ring |> List.map (fun r -> r.Phi * r.Mx, r.Phi * r.My)
    let cap =
        if cs.Mx = 0.0 && cs.My = 0.0 then None
        else
            try
                Some(Column.biaxial sec c cs.N cs.Mx cs.My d.Phi0 duct)
            with _ ->
                None
    let pts = (if view.Unfactored then unf else fac) @ fac @ [ cs.Mx, cs.My ]
    let lim = pts |> List.map (fun (x, y) -> max (abs x) (abs y)) |> List.max
    let top, bottom = 92.0, 58.0
    let f = frame (-lim) lim (-lim) lim 74.0 top (width - 300.0) (height - top - bottom) true
    let px (x, y) = f.Px x, f.Py y
    let mstar = sqrt (cs.Mx * cs.Mx + cs.My * cs.My)
    [ yield! title width (sprintf "Moment contour at N* = %.0f kN  –  %s%s" cs.N cs.Name (if cs.Description = "" then "" else "  ·  " + cs.Description)) (sub d)
      yield!
          legend 16.0 62.0
              [ if view.Unfactored then yield blue, false, "unfactored (Mux, Muy)"
                if view.Phi then yield orange, false, "design (φMux, φMuy)"
                if view.Loads then yield red, true, "design actions (M*x,d, M*y,d)" ]
      yield! axes f "Mx (kNm)" "My (kNm)" true
      if view.Unfactored then yield PPolyline(List.map px unf, blue, 1.6)
      if view.Phi then yield PPolyline(List.map px fac, orange, 2.2)
      if view.Loads then
          match cap with
          | Some r ->
              let ex, ey = px (r.Phi * r.Mx, r.Phi * r.My)
              let ox, oy = px (0.0, 0.0)
              yield PLine(ox, oy, ex, ey, red, 1.0, true)
              yield PPolygon([ ex - 4.0, ey; ex, ey - 4.0; ex + 4.0, ey; ex, ey + 4.0 ], orange, 1.0, ink, 0.8)
          | None -> ()
          let x, y = px (cs.Mx, cs.My)
          yield PPolygon(circlePts x y 4.5 16, red, 1.0, "#FFFFFF", 0.8)
          yield PText(x - 7.0, y + 9.0, cs.Name, red, 8.0, Right, true)
      // results on the right
      let kx = width - 210.0
      let row j (t: string) (bold: bool) = PText(kx, top + 4.0 + 17.0 * float j, t, ink, 9.0, Left, bold)
      yield row 0 cs.Name true
      yield row 1 (sprintf "N* = %.1f kN" cs.N) false
      yield row 2 (sprintf "M*x,d = %.1f   M*y,d = %.1f kNm" cs.Mx cs.My) false
      yield row 3 (sprintf "|M*d| = %.1f kNm" mstar) false
      match cap with
      | Some r ->
          yield row 4 (sprintf "φ = %.3f,  φMu = %.1f kNm" r.Phi (r.Phi * r.Mu)) false
          yield row 5 (sprintf "θ = %.1f°,  dn = %s mm" r.Theta (if Double.IsNaN r.Dn then "–" else sprintf "%.0f" r.Dn)) false
          let u = mstar / (r.Phi * r.Mu)
          yield PText(kx, top + 4.0 + 17.0 * 6.5, sprintf "|M*d| / φMu = %.0f %%" (100.0 * u), (if u <= 1.0 then green else red), 11.0, Left, true)
      | None ->
          if mstar > 0.0 then yield row 4 "no capacity at this N* (above φNuo?)" false
      yield PText(16.0, height - 14.0, "Strain compatibility (OAR.CONCRETE.MCONTOUR / COLUMN): capacity along the load direction at constant N*.", muted, 8.0, Left, false) ]

/// The picture for an item at a window size (logical pixels)
let scene (d: ColumnData) (item: Item) (view: View) (width: float) (height: float) : Scene =
    let w, h = max width 520.0, max height 380.0
    let prims =
        match item with
        | SectionItem -> sectionScene d w h
        | MNItem a -> mnScene d a view w h
        | ContourItem i -> contourScene d i view w h
    { Width = w; Height = h; Prims = prims }

// ---------------------------------------------------------------------------------------------
// SVG (tests and previews)
// ---------------------------------------------------------------------------------------------

let svg (s: Scene) =
    let esc (t: string) = t.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
    let n (v: float) = v.ToString("0.##", Globalization.CultureInfo.InvariantCulture)
    let pts (l: (float * float) list) = l |> List.map (fun (x, y) -> n x + "," + n y) |> String.concat " "
    let body =
        s.Prims
        |> List.map (fun p ->
            match p with
            | PLine(x1, y1, x2, y2, c, w, dashed) ->
                sprintf "<line x1='%s' y1='%s' x2='%s' y2='%s' stroke='%s' stroke-width='%s'%s/>" (n x1) (n y1) (n x2) (n y2) c (n w)
                    (if dashed then " stroke-dasharray='5,4'" else "")
            | PPolygon(p, fill, op, stroke, w) ->
                sprintf "<polygon points='%s' fill='%s' fill-opacity='%s' stroke='%s' stroke-width='%s'/>" (pts p) fill (n op)
                    (if stroke = "" then "none" else stroke) (n w)
            | PPolyline(p, c, w) -> sprintf "<polyline points='%s' fill='none' stroke='%s' stroke-width='%s'/>" (pts p) c (n w)
            | PText(x, y, t, c, size, al, bold) ->
                let anchor = match al with Left -> "start" | Centre -> "middle" | Right -> "end"
                sprintf "<text x='%s' y='%s' fill='%s' font-size='%s' text-anchor='%s' dominant-baseline='central' font-family='Segoe UI, Arial'%s>%s</text>"
                    (n x) (n y) c (n (size * 96.0 / 72.0)) anchor (if bold then " font-weight='bold'" else "") (esc t))
        |> String.concat "\n"
    sprintf "<svg xmlns='http://www.w3.org/2000/svg' width='%s' height='%s'><rect width='100%%' height='100%%' fill='white'/>\n%s\n</svg>" (n s.Width) (n s.Height) body
