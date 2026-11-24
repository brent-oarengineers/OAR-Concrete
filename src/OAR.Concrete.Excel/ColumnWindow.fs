/// OAR > Column graphics: a window that draws the active column tab – the section and bars to scale, the M–N
/// interaction about each axis with the design actions, and the moment contour at the N* of each load case.
/// Same window as OAR Steel / OAR Timber's View graphics.
///
/// The pictures come from OAR.Concrete.ColumnView (plain data, tested without Excel). This file only reads the
/// tab and paints. It finds the tab's per-case OAR.CONCRETE.COLUMN formulas and evaluates their arguments, so the
/// window uses exactly the section, bars, materials and design actions the sheet checks.
namespace OAR.Concrete.Excel

open System
open System.Drawing
open System.Drawing.Drawing2D
open System.Windows.Forms
open ExcelDna.Integration
open OAR.Concrete
open OAR.Concrete.ColumnView

/// What goes in the header of a printed page, and where a PDF is saved by default
type PageInfo =
    { Job: string
      Project: string
      By: string
      Date: string
      Ref: string
      Folder: string }

module private ColumnSheet =

    type XlSheet = Microsoft.Office.Interop.Excel.Worksheet
    type XlRange = Microsoft.Office.Interop.Excel.Range

    /// A value read through the object model, in the shape the worksheet functions get from Excel-DNA
    let private norm (v: obj) : obj =
        let one (o: obj) : obj =
            match o with
            | null -> box ExcelEmpty.Value
            | :? int -> box ExcelError.ExcelErrorValue // error cells arrive as an Int32 code
            | :? decimal as d -> box (float d)
            | :? DateTime as d -> box (d.ToOADate())
            | other -> other
        match v with
        | :? (obj[,]) as a ->
            let b0, b1 = a.GetLowerBound 0, a.GetLowerBound 1
            box (Array2D.init (a.GetLength 0) (a.GetLength 1) (fun i j -> one (a.GetValue(i + b0, j + b1))))
        | other -> one other

    let private missing = box ExcelMissing.Value

    let text (o: obj) =
        match o with
        | :? string as s -> s
        | :? float as v -> v.ToString(Globalization.CultureInfo.InvariantCulture)
        | _ -> ""

    let private marker = "OAR.CONCRETE.COLUMN("

    /// The arguments of the COLUMN call in a formula, split at the top-level commas
    let arguments (formula: string) =
        let i = formula.IndexOf(marker, StringComparison.OrdinalIgnoreCase)
        let s = formula.Substring(i + marker.Length)
        let args = ResizeArray<string>()
        let cur = Text.StringBuilder()
        let mutable depth = 0
        let mutable quoted = false
        let mutable fin = false
        for ch in s do
            if not fin then
                match ch with
                | '"' ->
                    quoted <- not quoted
                    cur.Append ch |> ignore
                | '(' when not quoted ->
                    depth <- depth + 1
                    cur.Append ch |> ignore
                | ')' when not quoted && depth = 0 ->
                    args.Add(cur.ToString().Trim())
                    fin <- true
                | ')' when not quoted ->
                    depth <- depth - 1
                    cur.Append ch |> ignore
                | ',' when not quoted && depth = 0 ->
                    args.Add(cur.ToString().Trim())
                    cur.Clear() |> ignore
                | c -> cur.Append c |> ignore
        List.ofSeq args

    /// Value of one argument (a reference or a constant), evaluated on the sheet; ExcelMissing when blank
    let evalArg (ws: XlSheet) (args: string list) (i: int) : obj =
        if i >= args.Length || args.[i] = "" then missing
        else
            match ws.Evaluate(args.[i]) with
            | :? XlRange as r -> norm r.Value2
            | v -> norm v

    /// (row, column, formula) of each per-case COLUMN formula: `=IF($E$65="","",OAR.CONCRETE.COLUMN(...))` in the
    /// capacity table (the uniaxial ones inside INDEX(...) are not cases)
    let caseFormulas (ws: XlSheet) =
        let used = ws.UsedRange
        let r2 = min 400 (used.Row + used.Rows.Count)
        let f = ws.Range(ws.Cells.[1, 5], ws.Cells.[r2, 10]).Formula
        match f with
        | :? (obj[,]) as a ->
            let b0, b1 = a.GetLowerBound 0, a.GetLowerBound 1
            [ for i in 0 .. a.GetLength 0 - 1 do
                  for j in 0 .. a.GetLength 1 - 1 do
                      match a.GetValue(i + b0, j + b1) with
                      | :? string as s when s.IndexOf(marker, StringComparison.OrdinalIgnoreCase) >= 0
                                            && s.IndexOf("INDEX(" + marker, StringComparison.OrdinalIgnoreCase) < 0 ->
                          yield i + 1, j + 5, s
                      | _ -> () ]
            // a spilled (CSE) formula reports its text in every cell of its range: one case per column, top cell
            |> List.distinctBy (fun (_, col, _) -> col)
        | _ -> []

    let isColumnSheet (ws: XlSheet) =
        try
            not (caseFormulas ws).IsEmpty
        with _ ->
            false

    /// Header field: the cell to the right of the label in rows 1 to 4 (columns F to I)
    let private header (ws: XlSheet) (label: string) =
        match norm (ws.Range("F1:J4").Value2) with
        | :? (obj[,]) as h ->
            [ for r in 0 .. Array2D.length1 h - 1 do
                  for c in 0 .. Array2D.length2 h - 2 do
                      if String.Equals((text h.[r, c]).Trim().TrimEnd('.', ':'), label, StringComparison.OrdinalIgnoreCase) then
                          yield h.[r, c + 1] ]
            |> List.tryHead
        | _ -> None

    let headerText (ws: XlSheet) (label: string) =
        match header ws label with
        | Some v -> (text v).Trim()
        | None -> ""

    let pageInfo (ws: XlSheet) : PageInfo =
        let safe (f: unit -> string) =
            try
                f ()
            with _ ->
                ""
        let project = safe (fun () -> headerText ws "PROJECT")
        { Job = safe (fun () -> headerText ws "JOB No")
          Project = (if project.StartsWith("Project name", StringComparison.OrdinalIgnoreCase) then "" else project)
          By = safe (fun () -> headerText ws "BY")
          Date =
            safe (fun () ->
                match header ws "DATE" with
                | Some(:? float as d) when d > 1.0 -> DateTime.FromOADate(d).ToString("dd/MM/yyyy")
                | Some v -> (text v).Trim()
                | None -> "")
          Ref = safe (fun () -> headerText ws "COLUMN")
          Folder = safe (fun () -> (ws.Parent :?> Microsoft.Office.Interop.Excel.Workbook).Path) }

    /// Bar names from the Bar column (G) of the bar table rows the formula uses (bars range L:P → G is 5 to the left)
    let private barNames (ws: XlSheet) (barsArg: string) =
        try
            match ws.Evaluate(barsArg) with
            | :? XlRange as r ->
                let areas = norm r.Value2
                let names = norm (r.Offset(0, -5).Resize(r.Rows.Count, 1).Value2)
                match areas, names with
                | (:? (obj[,]) as a), (:? (obj[,]) as n) ->
                    [ for i in 0 .. Array2D.length1 a - 1 do
                          match a.[i, 0] with
                          | :? float as v when v > 0.0 -> yield (text n.[i, 0]).Trim()
                          | _ -> () ]
                | _ -> []
            | _ -> []
        with _ ->
            []

    /// The guard of a case formula, IF($E$65="",...): the load-table row of the case
    let private caseRow (formula: string) =
        let m = Text.RegularExpressions.Regex.Match(formula, @"IF\(\$?E\$?(\d+)\s*=\s*""""")
        if m.Success then Some(int m.Groups.[1].Value) else None

    let private cellText (ws: XlSheet) (r: int) (c: int) = (text (norm (ws.Cells.[r, c] :?> XlRange).Value2)).Trim()

    /// Read the tab
    let read (ws: XlSheet) : ColumnData =
        let cases = caseFormulas ws
        if cases.IsEmpty then failwith "this isn't an OAR column tab (no OAR.CONCRETE.COLUMN formulas found)"
        let _, _, f0 = cases.Head
        let a0 = arguments f0
        let arg = evalArg ws a0
        let sec, c, duct = Functions.columnModel (arg 0) (arg 1) (arg 2) (arg 3) (arg 4) (arg 9) (arg 10)
        let phi0 = Functions.numberOpt (arg 8) |> Option.defaultValue 0.6
        let used =
            [ for r, col, f in cases do
                  let a = arguments f
                  let v = evalArg ws a
                  match Functions.numberOpt (v 5) with
                  | Some n ->
                      let name, descr =
                          match caseRow f with
                          | Some lr -> cellText ws lr 2, cellText ws lr 4
                          | None -> cellText ws (r - 1) col, ""
                      let num i = Functions.numberOpt (v i) |> Option.defaultValue 0.0
                      yield
                          { Name = (if name = "" then sprintf "Case %d" (col - 4) else name)
                            Description = descr
                            N = n
                            Mx = num 6
                            My = num 7 }
                  | None -> () ]
        let colRef = headerText ws "COLUMN"
        let location = headerText ws "LOCATION"
        { Title = ([ colRef; location ] |> List.filter (fun t -> t <> "") |> String.concat "  ·  " |> fun t -> if t = "" then ws.Name else t)
          Section = sec
          Concrete = c
          Duct = duct
          Phi0 = phi0
          BarNames = (if a0.Length > 2 then barNames ws a0.[2] else [])
          Cases = used }

/// The drawing surface (double-buffered)
type private Canvas() as this =
    inherit Control()

    do
        this.SetStyle(
            ControlStyles.AllPaintingInWmPaint
            ||| ControlStyles.UserPaint
            ||| ControlStyles.OptimizedDoubleBuffer
            ||| ControlStyles.ResizeRedraw,
            true
        )

module private Painter =

    let private colour (hex: string) (opacity: float) =
        let c = ColorTranslator.FromHtml hex
        Color.FromArgb(int (255.0 * max 0.0 (min 1.0 opacity)), c)

    let private points (pts: (float * float) list) =
        pts |> List.map (fun (x, y) -> PointF(float32 x, float32 y)) |> List.toArray

    let draw (g: Graphics) (scene: Scene) =
        g.SmoothingMode <- SmoothingMode.AntiAlias
        g.TextRenderingHint <- System.Drawing.Text.TextRenderingHint.ClearTypeGridFit
        for p in scene.Prims do
            match p with
            | PLine(x1, y1, x2, y2, c, w, dashed) ->
                use pen = new Pen(colour c 1.0, float32 w)
                if dashed then pen.DashStyle <- DashStyle.Dash
                g.DrawLine(pen, float32 x1, float32 y1, float32 x2, float32 y2)
            | PPolygon(pts, fill, opacity, stroke, w) ->
                let arr = points pts
                if arr.Length >= 3 then
                    if opacity > 0.0 then
                        use brush = new SolidBrush(colour fill opacity)
                        g.FillPolygon(brush, arr)
                    if stroke <> "" then
                        use pen = new Pen(colour stroke 1.0, float32 w)
                        g.DrawPolygon(pen, arr)
            | PPolyline(pts, c, w) ->
                let arr = points pts
                if arr.Length >= 2 then
                    use pen = new Pen(colour c 1.0, float32 w)
                    pen.LineJoin <- LineJoin.Round
                    g.DrawLines(pen, arr)
            | PText(x, y, t, c, size, align, bold) ->
                let style = if bold then FontStyle.Bold else FontStyle.Regular
                use font = new Font("Segoe UI", float32 (size * 96.0 / 72.0), style, GraphicsUnit.Pixel)
                use brush = new SolidBrush(colour c 1.0)
                use fmt = new StringFormat(StringFormat.GenericTypographic)
                fmt.Alignment <-
                    (match align with
                     | Align.Left -> StringAlignment.Near
                     | Align.Centre -> StringAlignment.Center
                     | Align.Right -> StringAlignment.Far)
                fmt.LineAlignment <- StringAlignment.Center
                fmt.FormatFlags <- fmt.FormatFlags ||| StringFormatFlags.NoWrap ||| StringFormatFlags.NoClip
                g.DrawString(t, font, brush, PointF(float32 x, float32 y), fmt)

/// Printing the picture (to a printer, or straight to a PDF file through Windows' PDF printer)
module private Printer =

    let pdfPrinter = "Microsoft Print to PDF"

    let hasPdfPrinter () =
        try
            [ for p in Printing.PrinterSettings.InstalledPrinters -> string p ] |> List.exists (fun p -> p = pdfPrinter)
        with _ ->
            false

    /// A4 landscape or portrait, 12 mm margins
    let setPage (portrait: bool) (doc: Printing.PrintDocument) =
        let a4 =
            [ for o in doc.PrinterSettings.PaperSizes -> box o ]
            |> List.tryPick (fun o ->
                match o with
                | :? Printing.PaperSize as sz when sz.Kind = Printing.PaperKind.A4 -> Some sz
                | _ -> None)
        match a4 with
        | Some sz -> doc.DefaultPageSettings.PaperSize <- sz
        | None -> ()
        doc.DefaultPageSettings.Landscape <- not portrait
        doc.DefaultPageSettings.Margins <- Printing.Margins(47, 47, 47, 47)

    /// One page per picture: job header, the picture scaled to the page, footer
    let page (info: PageInfo) (pages: (float -> float -> Scene) list) (index: int ref) (e: Printing.PrintPageEventArgs) =
        let g = e.Graphics
        let m = e.MarginBounds
        let left, top, width, height = float32 m.Left, float32 m.Top, float32 m.Width, float32 m.Height
        use bold = new Font("Segoe UI", 10.0f, FontStyle.Bold)
        use small = new Font("Segoe UI", 8.0f)
        use dark = new SolidBrush(Color.FromArgb(34, 34, 34))
        use grey = new SolidBrush(Color.FromArgb(110, 110, 110))
        use rule = new Pen(Color.FromArgb(70, 90, 82), 1.5f)
        use far = new StringFormat(Alignment = StringAlignment.Far)
        g.DrawString("OAR Consulting Engineers Pty Ltd", bold, dark, left, top)
        let job = [ (if info.Job = "" then "" else "Job " + info.Job); info.Project; info.Ref ] |> List.filter (fun t -> t <> "")
        g.DrawString(String.Join("   ·   ", job), small, grey, left, top + 18.0f)
        g.DrawString("Column interaction", bold, dark, RectangleF(left, top, width, 34.0f), far)
        let stamp = [ info.Date; (if info.By = "" then "" else "By " + info.By) ] |> List.filter (fun t -> t <> "")
        g.DrawString(String.Join("   ·   ", stamp), small, grey, RectangleF(left, top + 18.0f, width, 16.0f), far)
        g.DrawLine(rule, left, top + 38.0f, left + width, top + 38.0f)
        let version = Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString(3)
        g.DrawString(sprintf "OAR Concrete %s  ·  same engine and inputs as the calculation sheet" version, small, grey, left, top + height - 12.0f)
        g.DrawString(sprintf "%s   ·   page %d of %d" (DateTime.Now.ToString("dd/MM/yyyy HH:mm")) (index.Value + 1) pages.Length, small, grey,
                     RectangleF(left, top + height - 12.0f, width, 14.0f), far)
        let areaTop, areaHeight = top + 44.0f, height - 44.0f - 18.0f
        let logicalWidth = float width * 0.97
        let k = float width / logicalWidth
        let scene = pages.[index.Value] logicalWidth (float areaHeight / k)
        let state = g.Save()
        g.TranslateTransform(left, areaTop)
        g.ScaleTransform(float32 k, float32 k)
        Painter.draw g scene
        g.Restore state
        index.Value <- index.Value + 1
        e.HasMorePages <- index.Value < pages.Length

    let document (title: string) (info: PageInfo) (pages: (float -> float -> Scene) list) =
        let doc = new Printing.PrintDocument()
        doc.DocumentName <- title
        doc.PrintController <- new Printing.StandardPrintController()
        let index = ref 0
        doc.BeginPrint.Add(fun _ -> index.Value <- 0)
        doc.PrintPage.Add(fun e -> page info pages index e)
        doc

type private ColumnForm(load: unit -> ColumnData, pageInfo: unit -> PageInfo) as this =
    inherit Form()

    let tree =
        new TreeView(Dock = DockStyle.Left, Width = 230, HideSelection = false, ShowLines = false, ShowRootLines = false,
                     ShowPlusMinus = false, FullRowSelect = true, ShowNodeToolTips = true, BorderStyle = BorderStyle.None)
    let top = new FlowLayoutPanel(Dock = DockStyle.Top, Height = 36, WrapContents = false, Padding = Padding(10, 7, 8, 0))
    let check (text: string) = new CheckBox(Text = text, Checked = true, AutoSize = true, Margin = Padding(0, 3, 16, 0))
    let cbUnf = check "Unfactored"
    let cbPhi = check "φ capacity"
    let cbLoads = check "Design actions"
    let cbKey = check "Key points"
    let refresh = new Button(Text = "Refresh from sheet", AutoSize = true, Margin = Padding(12, 0, 0, 0))
    let savePdf = new Button(Text = "Save PDF", AutoSize = true, Margin = Padding(6, 0, 0, 0))
    let print = new Button(Text = "Print…", AutoSize = true, Margin = Padding(6, 0, 0, 0))
    let cbAll = new CheckBox(Text = "All pages", Checked = false, AutoSize = true, Margin = Padding(10, 3, 0, 0))
    let cbPortrait = new CheckBox(Text = "Portrait", Checked = false, AutoSize = true, Margin = Padding(10, 3, 0, 0))
    let status =
        new Label(Dock = DockStyle.Bottom, Height = 26, TextAlign = ContentAlignment.MiddleLeft, Padding = Padding(10, 0, 0, 0),
                  BackColor = Color.FromArgb(244, 244, 244))
    let canvas = new Canvas(Dock = DockStyle.Fill, BackColor = Color.White)
    let hint = "Pick the section, an M–N curve or a load case's moment contour on the left.  The window follows edits on the sheet."

    let mutable data: ColumnData option = None
    let mutable item: Item option = None
    let mutable scene: Scene option = None
    let mutable problem = ""

    let scale () =
        use g = canvas.CreateGraphics()
        float g.DpiX / 96.0

    let view () : View =
        { Unfactored = cbUnf.Checked
          Phi = cbPhi.Checked
          Loads = cbLoads.Checked
          KeyPoints = cbKey.Checked }

    let redraw () =
        scene <- None
        match data, item with
        | Some d, Some it ->
            try
                let s = scale ()
                scene <- Some(ColumnView.scene d it (view ()) (float canvas.ClientSize.Width / s) (float canvas.ClientSize.Height / s))
                problem <- ""
            with e ->
                problem <- e.Message
        | _ -> ()
        canvas.Invalidate()

    let label (d: ColumnData) (it: Item) =
        ColumnView.items d |> List.tryFind (fun (_, _, _, i) -> i = it) |> Option.map (fun (_, l, _, _) -> l) |> Option.defaultValue ""

    /// Print the picture showing (or every picture, "All pages") on A4: to a PDF file when `toPdf`, else the print dialog
    let output (toPdf: bool) =
        match data, item with
        | Some d, Some it ->
            try
                let info = pageInfo ()
                let v = view ()
                let chosen = if cbAll.Checked then ColumnView.items d |> List.map (fun (_, _, _, i) -> i) else [ it ]
                let pages = chosen |> List.map (fun i -> fun w h -> ColumnView.scene d i v w h)
                let what = if cbAll.Checked then "interaction" else label d it
                let title = sprintf "%s - %s" (if info.Ref = "" then "Column" else info.Ref) what
                use doc = Printer.document title info pages
                if toPdf && Printer.hasPdfPrinter () then
                    let clean =
                        String(title |> Seq.map (fun c -> if Array.contains c (IO.Path.GetInvalidFileNameChars()) then '-' else c) |> Seq.toArray)
                    use dialog =
                        new SaveFileDialog(Title = "Save column graphics as PDF", Filter = "PDF (*.pdf)|*.pdf", FileName = clean + ".pdf",
                                           OverwritePrompt = true)
                    if info.Folder <> "" && IO.Directory.Exists info.Folder then dialog.InitialDirectory <- info.Folder
                    if dialog.ShowDialog(this) = DialogResult.OK then
                        doc.PrinterSettings.PrinterName <- Printer.pdfPrinter
                        doc.PrinterSettings.PrintToFile <- true
                        doc.PrinterSettings.PrintFileName <- dialog.FileName
                        Printer.setPage cbPortrait.Checked doc
                        doc.Print()
                        status.Text <- "Saved " + dialog.FileName
                        try
                            Diagnostics.Process.Start(dialog.FileName) |> ignore
                        with _ ->
                            ()
                else
                    if toPdf then
                        MessageBox.Show(this, "The 'Microsoft Print to PDF' printer isn't installed on this computer, so pick a printer (or another PDF printer) instead.",
                                        "OAR Concrete – Column graphics")
                        |> ignore
                    use dialog = new PrintDialog(Document = doc, UseEXDialog = true, AllowSomePages = false)
                    Printer.setPage cbPortrait.Checked doc
                    if dialog.ShowDialog(this) = DialogResult.OK then
                        Printer.setPage cbPortrait.Checked doc
                        doc.Print()
                        status.Text <- "Sent to " + doc.PrinterSettings.PrinterName
            with e ->
                MessageBox.Show(this, "Couldn't print: " + e.Message, "OAR Concrete – Column graphics", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                |> ignore
        | _ -> ()

    /// Read the sheet again and fill the list, keeping the selected picture if it is still there
    let reload () =
        let keep = item
        try
            let d = load ()
            data <- Some d
            problem <- ""
            this.Text <- "OAR Concrete – Column graphics – " + d.Title
            tree.BeginUpdate()
            tree.Nodes.Clear()
            let mutable first: TreeNode = null
            let mutable kept: TreeNode = null
            let groups = Collections.Generic.Dictionary<string, TreeNode>()
            for g, l, tip, it in ColumnView.items d do
                let parent =
                    match groups.TryGetValue g with
                    | true, p -> p
                    | _ ->
                        let p = tree.Nodes.Add g
                        p.ForeColor <- Color.FromArgb(110, 110, 110)
                        groups.[g] <- p
                        p
                let node = parent.Nodes.Add l
                node.Tag <- box it
                node.ToolTipText <- tip
                if isNull first then first <- node
                if Some it = keep then kept <- node
            tree.ExpandAll()
            tree.EndUpdate()
            let pick = if isNull kept then first else kept
            tree.SelectedNode <- null
            tree.SelectedNode <- pick
        with e ->
            data <- None
            item <- None
            tree.Nodes.Clear()
            problem <- e.Message
            redraw ()

    do
        this.Text <- "OAR Concrete – Column graphics"
        this.Font <- new Font("Segoe UI", 9.0f)
        this.ClientSize <- Size(1220, 860)
        this.MinimumSize <- Size(820, 560)
        this.StartPosition <- FormStartPosition.CenterScreen
        this.ShowIcon <- false
        this.BackColor <- Color.White
        status.Text <- hint
        top.Controls.AddRange(
            [| cbUnf :> Control
               cbPhi :> Control
               cbLoads :> Control
               cbKey :> Control
               refresh :> Control
               savePdf :> Control
               print :> Control
               cbAll :> Control
               cbPortrait :> Control |]
        )
        this.Controls.Add canvas
        this.Controls.Add top
        this.Controls.Add status
        this.Controls.Add tree
        canvas.BringToFront()

        tree.AfterSelect.Add(fun e ->
            match (if isNull e.Node then null else e.Node.Tag) with
            | :? Item as it ->
                item <- Some it
                redraw ()
            | _ -> ())
        for cb in [ cbUnf; cbPhi; cbLoads; cbKey ] do
            cb.CheckedChanged.Add(fun _ -> redraw ())
        refresh.Click.Add(fun _ -> reload ())
        savePdf.Click.Add(fun _ -> output true)
        print.Click.Add(fun _ -> output false)
        canvas.Resize.Add(fun _ -> redraw ())
        canvas.Paint.Add(fun e ->
            let g = e.Graphics
            g.Clear Color.White
            let s = float32 (float g.DpiX / 96.0)
            g.ScaleTransform(s, s)
            match scene with
            | Some sc -> Painter.draw g sc
            | None -> ()
            if problem <> "" then
                use font = new Font("Segoe UI", 10.0f)
                use brush = new SolidBrush(Color.FromArgb(192, 57, 43))
                let area = RectangleF(20.0f, 60.0f, float32 canvas.ClientSize.Width / s - 40.0f, 200.0f)
                g.DrawString("Can't draw this: " + problem, font, brush, area))
        this.Shown.Add(fun _ -> reload ())

    member _.Reload() = reload ()

module ColumnWindow =

    let mutable private current: ColumnForm option = None
    let mutable private sheet: Microsoft.Office.Interop.Excel.Worksheet = null
    let mutable private pending = false

    let private openForm () =
        match current with
        | Some f when not f.IsDisposed -> Some f
        | _ -> None

    /// Queue one refresh for when Excel is idle (after it has recalculated)
    let private queueReload () =
        if (openForm ()).IsSome && not pending then
            pending <- true
            try
                ExcelAsyncUtil.QueueAsMacro(
                    ExcelAction(fun () ->
                        pending <- false
                        match openForm () with
                        | Some f -> f.Reload()
                        | None -> ())
                )
            with _ ->
                pending <- false

    /// After any cell is edited: an open window re-reads its tab and redraws
    let sheetChanged () = queueReload ()

    /// Another sheet brought forward: an open window switches to it if it is a column tab
    let sheetActivated (sh: obj) =
        match sh with
        | :? Microsoft.Office.Interop.Excel.Worksheet as ws when (openForm ()).IsSome && ColumnSheet.isColumnSheet ws ->
            sheet <- ws
            queueReload ()
        | _ -> ()

    let canShow (ws: Microsoft.Office.Interop.Excel.Worksheet) = ColumnSheet.isColumnSheet ws

    /// Open (or replace) the window for a tab; it stays open beside Excel and follows edits and the column tab in front
    let show (excelHwnd: int) (ws: Microsoft.Office.Interop.Excel.Worksheet) =
        match openForm () with
        | Some f -> f.Close()
        | None -> ()
        sheet <- ws
        let form = new ColumnForm((fun () -> ColumnSheet.read sheet), (fun () -> ColumnSheet.pageInfo sheet))
        current <- Some form
        let owner =
            { new IWin32Window with
                member _.Handle = IntPtr excelHwnd }
        form.Show owner
