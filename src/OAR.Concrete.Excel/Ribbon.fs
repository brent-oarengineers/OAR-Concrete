/// "OAR" ribbon tab: firm-wide workbook tools (not concrete-specific).
/// Export calcs to PDF: every visible tab with the OAR calc tab colour (465A52), in tab order,
/// using each tab's own print area and page setup, into one PDF.
module OAR.Concrete.Excel.Ribbon

open System
open System.IO
open System.Reflection
open System.Runtime.InteropServices
open System.Windows.Forms
open ExcelDna.Integration
open ExcelDna.Integration.CustomUI

/// OAR calc tab colour 465A52 as Excel stores it (BGR: R + 256 G + 65536 B)
let calcTabColour = 0x46 + 0x5A * 256 + 0x52 * 65536

// ---------- late-bound COM helpers (no interop assembly needed) ----------
let private get (o: obj) (name: string) (args: obj[]) =
    o.GetType().InvokeMember(name, BindingFlags.GetProperty, null, o, args)
let private call (o: obj) (name: string) (args: obj[]) =
    o.GetType().InvokeMember(name, BindingFlags.InvokeMethod, null, o, args)
let private setp (o: obj) (name: string) (value: obj) =
    o.GetType().InvokeMember(name, BindingFlags.SetProperty, null, o, [| value |]) |> ignore

let private isCalcTab (ws: obj) =
    let visible = get ws "Visible" [||]
    let isVisible =
        match visible with
        | :? int as v -> v = -1                       // xlSheetVisible
        | :? bool as b -> b
        | _ -> false
    let colour = get (get ws "Tab" [||]) "Color" [||]  // False when no tab colour
    let isCalc =
        match colour with
        | :? int as c -> c = calcTabColour
        | :? float as c -> int c = calcTabColour
        | _ -> false
    isVisible && isCalc

let private say (msg: string) (icon: MessageBoxIcon) =
    MessageBox.Show(msg, "OAR – Export calcs to PDF", MessageBoxButtons.OK, icon) |> ignore

let private exportTo (app: obj) (wb: obj) (sheets: obj) (calcTabs: string list) (file: string) =
    let original = get app "ActiveSheet" [||]
    let screen = get app "ScreenUpdating" [||]
    try
        try
            setp app "ScreenUpdating" (box false)
            call app "Calculate" [||] |> ignore
            // select the calc tabs as one group, then export the group as a single PDF
            let group = get sheets "Item" [| box (calcTabs |> List.map box |> Array.ofList) |]
            call group "Select" [||] |> ignore
            let active = get app "ActiveSheet" [||]
            // ExportAsFixedFormat(Type = xlTypePDF, Filename, Quality = standard, IncludeDocProperties,
            //                     IgnorePrintAreas, From, To, OpenAfterPublish)
            call active "ExportAsFixedFormat"
                [| box 0; box file; box 0; box true; box false; Type.Missing; Type.Missing; box true |] |> ignore
        with ex ->
            let inner =
                match ex with
                | :? TargetInvocationException as t when not (isNull t.InnerException) -> t.InnerException
                | e -> e
            say ("Export failed: " + inner.Message) MessageBoxIcon.Warning
    finally
        (try call original "Select" [||] |> ignore with _ -> ())   // ungroup the tabs
        setp app "ScreenUpdating" screen

let exportCalcs () =
    let app = ExcelDnaUtil.Application
    let wb = get app "ActiveWorkbook" [||]
    if isNull wb then
        say "Open a calc workbook first." MessageBoxIcon.Information
    else
        let sheets = get wb "Worksheets" [||]
        let n = get sheets "Count" [||] :?> int
        let calcTabs =
            [ for i in 1 .. n do
                let ws = get sheets "Item" [| box i |]
                if isCalcTab ws then yield (get ws "Name" [||] :?> string) ]
        if calcTabs.IsEmpty then
            say "No calc tabs found. Calc tabs are the visible tabs coloured OAR green (465A52)." MessageBoxIcon.Information
        else
            let fullName = get wb "FullName" [||] :?> string
            let folder =
                match get wb "Path" [||] :?> string with
                | "" -> Environment.GetFolderPath Environment.SpecialFolder.MyDocuments
                | p -> p
            let initial = Path.Combine(folder, Path.GetFileNameWithoutExtension fullName + " - Calcs.pdf")
            let chosen = call app "GetSaveAsFilename" [| box initial; box "PDF Files (*.pdf), *.pdf"; box 1; box "Export calcs to PDF" |]
            match chosen with
            | :? string as file when file <> "" -> exportTo app wb sheets calcTabs file
            | _ -> ()   // cancelled (Excel returns False)

// ---------- OAR > Column graphics ----------

let private excel () = ExcelDnaUtil.Application :?> Microsoft.Office.Interop.Excel.Application

/// Run an event handler without ever letting an error reach Excel
let private safely (f: unit -> unit) =
    try
        f ()
    with _ ->
        ()

let private tell (text: string) =
    safely (fun () ->
        let a = excel ()
        let owner = { new IWin32Window with member _.Handle = nativeint a.Hwnd }
        MessageBox.Show(owner, text, "OAR Concrete – Column graphics", MessageBoxButtons.OK, MessageBoxIcon.Information) |> ignore)

/// The section, M–N curves and moment contours of the active column tab, in their own window
let columnGraphics () =
    try
        let a = excel ()
        match a.ActiveSheet with
        | :? Microsoft.Office.Interop.Excel.Worksheet as ws when ColumnWindow.canShow ws ->
            ColumnWindow.show a.Hwnd ws
            safely (fun () -> a.StatusBar <- "Column graphics – window opened (it follows edits and the column tab in front)")
        | :? Microsoft.Office.Interop.Excel.Worksheet ->
            tell "The active tab is not an OAR column tab: it has no OAR.CONCRETE.COLUMN formulas in its capacity table. Open a tab made from OAR - RC Column Design."
        | _ -> tell "The active sheet is not a worksheet."
    with e ->
        tell ("The window couldn't open: " + e.Message)

[<ExcelCommand(Name = "OAR.CONCRETE.COLUMNGRAPHICS", Description = "Open the column graphics window for the active column tab")>]
let columnGraphicsCommand () = columnGraphics ()

// keep the Application object alive so its event connections are not collected
let mutable private hooked: Microsoft.Office.Interop.Excel.Application = null

let hook () =
    if isNull hooked then
        let a = excel ()
        a.add_SheetChange(Microsoft.Office.Interop.Excel.AppEvents_SheetChangeEventHandler(fun _ _ -> safely ColumnWindow.sheetChanged))
        a.add_SheetActivate(Microsoft.Office.Interop.Excel.AppEvents_SheetActivateEventHandler(fun sh -> safely (fun () -> ColumnWindow.sheetActivated sh)))
        a.add_WorkbookActivate(Microsoft.Office.Interop.Excel.AppEvents_WorkbookActivateEventHandler(fun wb -> safely (fun () -> ColumnWindow.sheetActivated wb.ActiveSheet)))
        hooked <- a

type AddIn() =
    interface IExcelAddIn with
        member _.AutoOpen() = ExcelAsyncUtil.QueueAsMacro(ExcelAction(fun () -> safely hook))
        member _.AutoClose() = ()

/// Ribbon. The tab uses a shared idQ so OAR Steel / OAR Timber can add their own groups to the same "OAR" tab.
[<ComVisible(true)>]
type OarRibbon() =
    inherit ExcelRibbon()
    override _.GetCustomUI(_ribbonId) =
        """<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" xmlns:oar="OAR.Engineers">
  <ribbon>
    <tabs>
      <tab idQ="oar:OarTab" label="OAR">
        <group id="OarConcreteCalcs" label="Calcs">
          <button id="OarExportCalcs" label="Export calcs to PDF" size="large" imageMso="FileSaveAsPdfOrXps"
                  screentip="Export calcs to PDF"
                  supertip="Exports every visible calc tab (OAR green tab colour) in tab order into one PDF, using each tab's print area. Data tabs are skipped."
                  onAction="OnExportCalcs" />
        </group>
        <group id="OarConcreteColumn" label="Concrete column">
          <button id="OarColumnGraphics" label="Column graphics" size="large" imageMso="ChartTypeXYScatterInsertGallery"
                  screentip="Column graphics"
                  supertip="Draw the active column tab: the section and bars to scale, the M–N interaction about x and y with the design actions, and the moment contour at each load case's N*. Save PDF or print from the window."
                  onAction="OnColumnGraphics" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>"""
    member _.OnExportCalcs(_control: IRibbonControl) =
        ExcelAsyncUtil.QueueAsMacro(fun () -> exportCalcs ())
    member _.OnColumnGraphics(_control: IRibbonControl) = columnGraphics ()
