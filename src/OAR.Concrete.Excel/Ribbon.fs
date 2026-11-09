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
      </tab>
    </tabs>
  </ribbon>
</customUI>"""
    member _.OnExportCalcs(_control: IRibbonControl) =
        ExcelAsyncUtil.QueueAsMacro(fun () -> exportCalcs ())
