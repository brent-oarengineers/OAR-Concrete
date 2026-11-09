/// Excel worksheet functions. This file only converts between Excel values and the engine;
/// all analysis is in OAR.Concrete.Core so it can be tested without Excel.
module OAR.Concrete.Excel.Functions

open System
open ExcelDna.Integration
open OAR.Concrete
open OAR.Concrete.Materials

[<Literal>]
let Category = "OAR Concrete"

let private version = "0.5.0"

// ---------- conversion helpers (same as OAR Timber) ----------

let private isMissing (o: obj) =
    match o with
    | null -> true
    | :? ExcelMissing
    | :? ExcelEmpty -> true
    | :? string as s -> s.Trim() = ""
    | _ -> false

let private tryNum (o: obj) =
    match o with
    | :? float as v -> Some v
    | :? string as s ->
        match Double.TryParse(s.Trim(), Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture) with
        | true, v -> Some v
        | _ -> None
    | :? bool as b -> Some(if b then 1.0 else 0.0)
    | _ -> None

let private num (name: string) (o: obj) =
    match o with
    | :? ExcelError -> failwithf "%s contains an Excel error" name
    | _ when isMissing o -> failwithf "%s is blank" name
    | _ ->
        match tryNum o with
        | Some v -> v
        | None -> failwithf "%s: '%O' is not a number" name o

let private numOpt (name: string) (o: obj) = if isMissing o then None else Some(num name o)

let private cellOpt (o: obj) =
    match o with
    | :? ExcelError -> None
    | _ when isMissing o -> None
    | _ -> tryNum o

let private text (o: obj) = if isMissing o then "" else (string o).Trim()

let private range (o: obj) : obj[,] =
    match o with
    | :? (obj[,]) as a -> a
    | v -> Array2D.create 1 1 v

/// Run a function and turn any failure into readable text in the cell
let private guard (f: unit -> obj) : obj =
    try
        f ()
    with ex ->
        box ("#OAR! " + ex.Message)

let private column (values: float list) =
    let out = Array2D.create values.Length 1 (box "")
    values |> List.iteri (fun i v -> out.[i, 0] <- (if Double.IsNaN v || Double.IsInfinity v then box "–" else box v))
    box out

// ---------- section from the sheet ----------

/// shape "Rect" with dims {b, D}, or "Circle" with dims {D}. Returns the builder and the default bar x.
let private shapeOf (shape: obj) (dims: obj) =
    let d = range dims |> Seq.cast<obj> |> Seq.choose cellOpt |> Array.ofSeq
    let s = (text shape).ToUpperInvariant()
    if s = "" || s.StartsWith "RECT" || s = "SLAB" || s = "BEAM" then
        if d.Length < 2 then failwith "dims: a rectangle needs 2 cells, b and D (mm)"
        (fun bars -> Section.rectangle d.[0] d.[1] bars), d.[0] / 2.0
    elif s.StartsWith "CIRC" then
        if d.Length < 1 then failwith "dims: a circle needs its diameter D (mm)"
        (fun bars -> Section.circular d.[0] bars), d.[0] / 2.0
    else failwithf "shape '%s' not recognised (Rect, Circle)" (text shape)

/// Bars table, one row per bar or layer: As (mm²) | depth from top face (mm) | fsy (MPa) | [x from left (mm)] | [Es (MPa)].
/// Rows with a blank or zero area are skipped.
/// Returns (sheet row index, bar) so a parallel range (e.g. bar classes) can be matched up.
let private barRowsOf (o: obj) (xDefault: float) =
    let a = range o
    if Array2D.length2 a < 3 then failwith "bars: select 3 columns As | depth | fsy (optionally x, Es)"
    [ for i in 0 .. Array2D.length1 a - 1 do
          match cellOpt a.[i, 0] with
          | Some area when area > 0.0 ->
              let row = sprintf "bars row %d" (i + 1)
              let depth = num (row + " depth") a.[i, 1]
              let fsy = num (row + " fsy") a.[i, 2]
              let x = if Array2D.length2 a > 3 then defaultArg (cellOpt a.[i, 3]) xDefault else xDefault
              let es = if Array2D.length2 a > 4 then defaultArg (cellOpt a.[i, 4]) EsDefault else EsDefault
              if fsy <= 0.0 then failwithf "%s: fsy must be positive" row
              yield i, Section.barAt area depth fsy x es
          | _ -> () ]

let private barsOf (o: obj) (xDefault: float) = barRowsOf o xDefault |> List.map snd

let private sectionOf shape dims bars =
    let build, xd = shapeOf shape dims
    let b = barsOf bars xd
    if b.IsEmpty then failwith "bars: no reinforcement"
    build b

/// "Sagging" (top in compression) = 90°, "Hogging" = −90°, or an angle in degrees
let private senseOf (o: obj) =
    match o with
    | :? float as v -> v
    | _ ->
        let t = (text o).ToUpperInvariant()
        match tryNum o with
        | Some v when not (o :? bool) -> v
        | _ ->
            if t.StartsWith "SAG" || t = "+" then 90.0
            elif t.StartsWith "HOG" || t = "-" then -90.0
            else failwithf "sense '%s' not recognised (Sagging, Hogging, or angle in degrees)" (text o)

// ---------- functions ----------

[<ExcelFunction(Name = "OAR.CONCRETE.VERSION", Description = "OAR Concrete add-in version", Category = Category)>]
let concreteVersion () : obj = box ("OAR Concrete " + version)

[<ExcelFunction(Name = "OAR.CONCRETE.ULT",
                Description = "Ultimate strength by strain compatibility (AS 3600 Cl 8.1). Spills 16 values down: dn mm | kuo | ku | do mm | d mm | εst | σst MPa | σsc MPa | Cc kN | Cs kN | T kN | z mm | Mu kNm | φ | kL | εL (Class L)",
                Category = Category,
                IsThreadSafe = true)>]
let concreteUlt
    ([<ExcelArgument(Name = "shape", Description = "Rect or Circle")>] shape: obj)
    ([<ExcelArgument(Name = "dims", Description = "Rect: 2 cells b, D (mm). Circle: D (mm)")>] dims: obj)
    ([<ExcelArgument(Name = "bars", Description = "Rows: As mm² | depth from top mm | fsy MPa | [x from left mm] | [Es MPa]")>] bars: obj)
    ([<ExcelArgument(Name = "fc", Description = "f'c, MPa")>] fc: obj)
    ([<ExcelArgument(Name = "model", Description = "Rectangular (Cl 8.1.3), Curvilinear (Cl 3.1.4) or Parabolic")>] model: obj)
    ([<ExcelArgument(Name = "sense", Description = "Sagging (top in compression), Hogging, or angle of the compression side in degrees (90 = top)")>] sense: obj)
    ([<ExcelArgument(Name = "[N]", Description = "Axial force kN, compression +. Blank = 0")>] nArg: obj)
    ([<ExcelArgument(Name = "[Ec]", Description = "Ec MPa for the curvilinear law. Blank = Table 3.1.2")>] ecArg: obj)
    ([<ExcelArgument(Name = "[classes]", Description = "Bar ductility class per bars row, N or L (blank = N). N = 0 only")>] classesArg: obj)
    ([<ExcelArgument(Name = "[classL]", Description = "Whole section (φ = 0.65 if Class L bars are in tension) or L steel only (Class L areas × 0.65/φ). Blank = whole section")>] modeArg: obj)
    : obj =
    guard (fun () ->
        let build, xd = shapeOf shape dims
        let rows = barRowsOf bars xd
        if rows.IsEmpty then failwith "bars: no reinforcement"
        let sec = build (rows |> List.map snd)
        let c = concrete (num "fc" fc) (numOpt "Ec" ecArg) (parseModel (text model))
        let n = defaultArg (numOpt "N" nArg) 0.0
        let theta = senseOf sense
        if n <> 0.0 then
            let r = Ultimate.analyse sec c theta n
            column [ r.Dn; r.Kuo; r.Ku; r.Do; r.D; r.EpsSt; r.SigSt; r.SigSc; r.Cc; r.Cs; r.T; r.Z; abs r.Mu; nan; nan; nan ]
        else
            let cls = range classesArg |> Seq.cast<obj> |> Array.ofSeq
            let flags = rows |> List.map (fun (i, _) -> i < cls.Length && (text cls.[i]).ToUpperInvariant() = "L")
            let mode = if (text modeArg).ToUpperInvariant().StartsWith "L" then Flexure.LSteelOnly else Flexure.WholeSection
            let f = Flexure.capacity sec c theta flags mode
            let r = f.Ult
            column [ r.Dn; r.Kuo; r.Ku; r.Do; r.D; r.EpsSt; r.SigSt; r.SigSc; r.Cc; r.Cs; r.T; r.Z; abs r.Mu; f.Phi; f.KL; f.EpsL ])

[<ExcelFunction(Name = "OAR.CONCRETE.CRACKED",
                Description = "Cracked elastic section (no concrete tension). Spills 5 values down: dn mm | Icr ×10⁶ mm⁴ | σst MPa (largest bar tension) | σsc MPa | σc MPa (extreme fibre)",
                Category = Category,
                IsThreadSafe = true)>]
let concreteCracked
    ([<ExcelArgument(Name = "shape", Description = "Rect or Circle")>] shape: obj)
    ([<ExcelArgument(Name = "dims", Description = "Rect: 2 cells b, D (mm). Circle: D (mm)")>] dims: obj)
    ([<ExcelArgument(Name = "bars", Description = "Rows: As mm² | depth from top mm | fsy MPa | [x from left mm] | [Es MPa]")>] bars: obj)
    ([<ExcelArgument(Name = "Ec", Description = "Concrete modulus for the modular ratio, MPa (Ec / (1 + φcc) for long term)")>] ecArg: obj)
    ([<ExcelArgument(Name = "M", Description = "Service moment, kNm (magnitude)")>] mArg: obj)
    ([<ExcelArgument(Name = "sense", Description = "Sagging (top in compression), Hogging, or angle in degrees")>] sense: obj)
    ([<ExcelArgument(Name = "[N]", Description = "Axial compression kN. Blank = 0")>] nArg: obj)
    : obj =
    guard (fun () ->
        let sec = sectionOf shape dims bars
        let m = defaultArg (numOpt "M" mArg) 0.0
        let r = Service.analyse sec (num "Ec" ecArg) (senseOf sense) m (defaultArg (numOpt "N" nArg) 0.0)
        column [ r.Dn; r.Icr / 1e6; r.SigSt; r.SigSc; r.SigC ])

[<ExcelFunction(Name = "OAR.CONCRETE.MNCURVE",
                Description = "M–N interaction curve for one bending direction (preview for the column sheet). Spills rows: N kN | M kNm | dn mm | φ | φN kN | φM kNm, from pure tension to squash",
                Category = Category,
                IsThreadSafe = true)>]
let concreteMnCurve
    ([<ExcelArgument(Name = "shape", Description = "Rect or Circle")>] shape: obj)
    ([<ExcelArgument(Name = "dims", Description = "Rect: 2 cells b, D (mm). Circle: D (mm)")>] dims: obj)
    ([<ExcelArgument(Name = "bars", Description = "Rows: As mm² | depth from top mm | fsy MPa | [x from left mm] | [Es MPa]")>] bars: obj)
    ([<ExcelArgument(Name = "fc", Description = "f'c, MPa")>] fc: obj)
    ([<ExcelArgument(Name = "model", Description = "Rectangular, Curvilinear or Parabolic")>] model: obj)
    ([<ExcelArgument(Name = "sense", Description = "Sagging / Hogging or angle in degrees")>] sense: obj)
    ([<ExcelArgument(Name = "[points]", Description = "Number of strain-compatibility points, default 20")>] pointsArg: obj)
    ([<ExcelArgument(Name = "[phi0]", Description = "φ for compression members (Table 2.2.2), default 0.65")>] phi0Arg: obj)
    ([<ExcelArgument(Name = "[class]", Description = "Reinforcement ductility N or L, default N")>] classArg: obj)
    ([<ExcelArgument(Name = "[Ec]", Description = "Ec MPa for the curvilinear law. Blank = Table 3.1.2")>] ecArg: obj)
    : obj =
    guard (fun () ->
        let sec = sectionOf shape dims bars
        let c = concrete (num "fc" fc) (numOpt "Ec" ecArg) (parseModel (text model))
        let n = int (defaultArg (numOpt "points" pointsArg) 20.0)
        let duct = if (text classArg).ToUpperInvariant() = "L" then Design.ClassL else Design.ClassN
        let pts = Interaction.curve sec c (senseOf sense) (max 2 n) (defaultArg (numOpt "phi0" phi0Arg) 0.65) duct
        let out = Array2D.create pts.Length 6 (box "")
        pts
        |> List.iteri (fun i p ->
            [ p.N; p.M; p.Dn; p.Phi; p.PhiN; p.PhiM ]
            |> List.iteri (fun j v -> out.[i, j] <- (if Double.IsNaN v then box "–" else box v)))
        box out)

// ---------- columns ----------

let private ductOf (classes: obj) =
    if range classes |> Seq.cast<obj> |> Seq.exists (fun o -> (text o).ToUpperInvariant() = "L") then Design.ClassL else Design.ClassN

/// "x" (or 90) = bending about the x axis, compression on the top face; "y" (or 0) = about the y axis
let private axisOf (o: obj) =
    match (text o).ToUpperInvariant() with
    | "X" | "MX" -> 90.0
    | "Y" | "MY" -> 0.0
    | _ -> senseOf o

let private colArgs shape dims bars fc model ecArg =
    sectionOf shape dims bars, concrete (num "fc" fc) (numOpt "Ec" ecArg) (parseModel (text model))

let private grid (rows: float list list) =
    let n = rows.Length
    let m = if n = 0 then 0 else rows.Head.Length
    let out = Array2D.create n m (box "")
    rows |> List.iteri (fun i r -> r |> List.iteri (fun j v -> out.[i, j] <- (if Double.IsNaN v then box "–" else box v)))
    box out

[<ExcelFunction(Name = "OAR.CONCRETE.BARLAYOUT",
                Description = "Typical column bar layout. Spills rows x | y (mm from the section centre), padded with blanks to [rows]. Rect: n1 bars along each b face, n2 along each D face (corners counted in both). Circle: n1 bars, first at the top",
                Category = Category, IsThreadSafe = true)>]
let concreteBarLayout
    ([<ExcelArgument(Name = "shape", Description = "Rect or Circle")>] shape: obj)
    ([<ExcelArgument(Name = "dims", Description = "Rect: b, D (mm). Circle: D (mm)")>] dims: obj)
    ([<ExcelArgument(Name = "edge", Description = "Face to bar centre: cover + ligature + db/2 (mm)")>] edgeArg: obj)
    ([<ExcelArgument(Name = "n1", Description = "Rect: bars per b face incl. corners. Circle: number of bars")>] n1Arg: obj)
    ([<ExcelArgument(Name = "[n2]", Description = "Rect: bars per D face incl. corners")>] n2Arg: obj)
    ([<ExcelArgument(Name = "[rows]", Description = "Rows to fill (blank rows after the last bar), default 40")>] rowsArg: obj)
    : obj =
    guard (fun () ->
        let d = range dims |> Seq.cast<obj> |> Seq.choose cellOpt |> Array.ofSeq
        let s = (text shape).ToUpperInvariant()
        let edge = num "edge" edgeArg
        let n1 = int (num "n1" n1Arg)
        let pts =
            if s.StartsWith "CIRC" then Column.circLayout d.[0] edge n1
            else Column.rectLayout d.[0] d.[1] edge n1 (int (num "n2" n2Arg))
        let rows = int (defaultArg (numOpt "rows" rowsArg) 40.0)
        if pts.Length > rows then failwithf "%d bars do not fit in the %d-row table" pts.Length rows
        let out = Array2D.create rows 2 (box "")
        pts |> List.iteri (fun i (x, y) -> out.[i, 0] <- box (Math.Round(x, 3)); out.[i, 1] <- box (Math.Round(y, 3)))
        box out)

[<ExcelFunction(Name = "OAR.CONCRETE.COLUMN",
                Description = "Column capacity at N* along the direction of (M*x, M*y), strain compatibility (biaxial). Spills 10 values down: φ | Mu kNm | φMu kNm | Mux kNm | Muy kNm | θ deg | dn mm | Nub kN | Nuo kN | εs outermost tension bar",
                Category = Category, IsThreadSafe = true)>]
let concreteColumn
    ([<ExcelArgument(Name = "shape", Description = "Rect or Circle")>] shape: obj)
    ([<ExcelArgument(Name = "dims", Description = "Rect: b, D (mm). Circle: D (mm)")>] dims: obj)
    ([<ExcelArgument(Name = "bars", Description = "Rows: As mm² | depth from top mm | fsy MPa | x from left mm | [Es MPa]")>] bars: obj)
    ([<ExcelArgument(Name = "fc", Description = "f'c, MPa")>] fc: obj)
    ([<ExcelArgument(Name = "model", Description = "Rectangular, Curvilinear or Parabolic")>] model: obj)
    ([<ExcelArgument(Name = "N", Description = "N*, kN (compression +)")>] nArg: obj)
    ([<ExcelArgument(Name = "Mx", Description = "M*x, kNm (about the x axis)")>] mxArg: obj)
    ([<ExcelArgument(Name = "My", Description = "M*y, kNm (about the y axis)")>] myArg: obj)
    ([<ExcelArgument(Name = "[phi0]", Description = "φ0 for compression (Table 2.2.2(d)), default 0.6")>] phi0Arg: obj)
    ([<ExcelArgument(Name = "[classes]", Description = "Bar classes (any L → Class L φ)")>] classesArg: obj)
    ([<ExcelArgument(Name = "[Ec]", Description = "Ec MPa (curvilinear / parabolic). Blank = Table 3.1.2")>] ecArg: obj)
    : obj =
    guard (fun () ->
        let sec, c = colArgs shape dims bars fc model ecArg
        let mx = defaultArg (numOpt "Mx" mxArg) 0.0
        let my = defaultArg (numOpt "My" myArg) 0.0
        let r = Column.biaxial sec c (num "N" nArg) mx my (defaultArg (numOpt "phi0" phi0Arg) 0.6) (ductOf classesArg)
        column [ r.Phi; r.Mu; r.Phi * r.Mu; r.Mx; r.My; r.Theta; r.Dn; r.Nub; r.Nuo; r.EpsS ])

[<ExcelFunction(Name = "OAR.CONCRETE.COLUMNAXIS",
                Description = "Column properties for bending about one axis. Spills 11 values down: Nuo kN | Nuot kN | Nub kN | Mub kNm | Muo kNm | kuo | φb | do mm | Nud kN | Mud kNm | φ tension",
                Category = Category, IsThreadSafe = true)>]
let concreteColumnAxis
    ([<ExcelArgument(Name = "shape", Description = "Rect or Circle")>] shape: obj)
    ([<ExcelArgument(Name = "dims", Description = "Rect: b, D (mm). Circle: D (mm)")>] dims: obj)
    ([<ExcelArgument(Name = "bars", Description = "Rows: As mm² | depth from top mm | fsy MPa | x from left mm | [Es MPa]")>] bars: obj)
    ([<ExcelArgument(Name = "fc", Description = "f'c, MPa")>] fc: obj)
    ([<ExcelArgument(Name = "model", Description = "Rectangular, Curvilinear or Parabolic")>] model: obj)
    ([<ExcelArgument(Name = "axis", Description = "x (Mx, top face in compression) or y (My, +x face in compression)")>] axisArg: obj)
    ([<ExcelArgument(Name = "[classes]", Description = "Bar classes (any L → Class L φ)")>] classesArg: obj)
    ([<ExcelArgument(Name = "[Ec]", Description = "Ec MPa. Blank = Table 3.1.2")>] ecArg: obj)
    : obj =
    guard (fun () ->
        let sec, c = colArgs shape dims bars fc model ecArg
        let a = Column.axisProps sec c (axisOf axisArg) (ductOf classesArg)
        column [ a.Nuo; a.Nuot; a.Nub; a.Mub; a.Muo; a.Kuo; a.PhiB; a.Do; a.Nud; a.Mud; a.PhiT ])

[<ExcelFunction(Name = "OAR.CONCRETE.MNCOLUMN",
                Description = "Column M–N curve for one axis, 40 rows from pure tension to the squash load: N kN | M kNm | φ | φN kN | φM kNm",
                Category = Category, IsThreadSafe = true)>]
let concreteMnColumn
    ([<ExcelArgument(Name = "shape", Description = "Rect or Circle")>] shape: obj)
    ([<ExcelArgument(Name = "dims", Description = "Rect: b, D (mm). Circle: D (mm)")>] dims: obj)
    ([<ExcelArgument(Name = "bars", Description = "Rows: As mm² | depth from top mm | fsy MPa | x from left mm | [Es MPa]")>] bars: obj)
    ([<ExcelArgument(Name = "fc", Description = "f'c, MPa")>] fc: obj)
    ([<ExcelArgument(Name = "model", Description = "Rectangular, Curvilinear or Parabolic")>] model: obj)
    ([<ExcelArgument(Name = "axis", Description = "x or y")>] axisArg: obj)
    ([<ExcelArgument(Name = "[phi0]", Description = "φ0 for compression, default 0.6")>] phi0Arg: obj)
    ([<ExcelArgument(Name = "[classes]", Description = "Bar classes (any L → Class L φ)")>] classesArg: obj)
    ([<ExcelArgument(Name = "[Ec]", Description = "Ec MPa. Blank = Table 3.1.2")>] ecArg: obj)
    : obj =
    guard (fun () ->
        let sec, c = colArgs shape dims bars fc model ecArg
        Column.mnCurve sec c (axisOf axisArg) (defaultArg (numOpt "phi0" phi0Arg) 0.6) (ductOf classesArg)
        |> List.map (fun (n, m, p) -> [ n; m; p; p * n; p * m ])
        |> grid)

[<ExcelFunction(Name = "OAR.CONCRETE.MCONTOUR",
                Description = "Biaxial moment capacity contour at axial force N (closed loop, [points] + 1 rows): Mux | Muy | φMux | φMuy (kNm)",
                Category = Category, IsThreadSafe = true)>]
let concreteMContour
    ([<ExcelArgument(Name = "shape", Description = "Rect or Circle")>] shape: obj)
    ([<ExcelArgument(Name = "dims", Description = "Rect: b, D (mm). Circle: D (mm)")>] dims: obj)
    ([<ExcelArgument(Name = "bars", Description = "Rows: As mm² | depth from top mm | fsy MPa | x from left mm | [Es MPa]")>] bars: obj)
    ([<ExcelArgument(Name = "fc", Description = "f'c, MPa")>] fc: obj)
    ([<ExcelArgument(Name = "model", Description = "Rectangular, Curvilinear or Parabolic")>] model: obj)
    ([<ExcelArgument(Name = "N", Description = "Axial force, kN (compression +)")>] nArg: obj)
    ([<ExcelArgument(Name = "[points]", Description = "Directions around the section, default 36")>] pointsArg: obj)
    ([<ExcelArgument(Name = "[phi0]", Description = "φ0 for compression, default 0.6")>] phi0Arg: obj)
    ([<ExcelArgument(Name = "[classes]", Description = "Bar classes (any L → Class L φ)")>] classesArg: obj)
    ([<ExcelArgument(Name = "[Ec]", Description = "Ec MPa. Blank = Table 3.1.2")>] ecArg: obj)
    : obj =
    guard (fun () ->
        let sec, c = colArgs shape dims bars fc model ecArg
        let n = int (defaultArg (numOpt "points" pointsArg) 36.0)
        Column.contour sec c (num "N" nArg) n (defaultArg (numOpt "phi0" phi0Arg) 0.6) (ductOf classesArg)
        |> List.map (fun r -> [ r.Mx; r.My; r.Phi * r.Mx; r.Phi * r.My ])
        |> grid)
