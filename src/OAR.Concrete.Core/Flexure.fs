/// Bending capacity with φ, including the two ways of treating Class L reinforcement.
///
/// Whole section: Table 2.2.2 as written – if Class L bars are in tension, φ = 0.65 for the section.
/// L steel only (engineering judgement, not a Table 2.2.2 provision): the Class L bar areas are
/// multiplied by k = 0.65 / φ and φ comes from the Class N formula on the reduced section, so the
/// Class L steel works at 0.65 and the Class N steel at φ. k depends on φ, so the two are iterated
/// (start φ = 0.85; usually settles in one or two passes).
module OAR.Concrete.Flexure

open OAR.Concrete.Materials
open OAR.Concrete.Geometry
open OAR.Concrete.Section

type ClassLMode =
    | WholeSection
    | LSteelOnly

type Result =
    { /// analysis of the section actually used (Class L areas reduced in LSteelOnly mode)
      Ult: Ultimate.Result
      Phi: float
      /// factor on the Class L bar areas (1 in WholeSection mode or with no Class L bars)
      KL: float
      /// largest tensile strain in a Class L bar at ultimate (0 if none in tension)
      EpsL: float }

/// Largest tensile strain among the flagged bars, for the analysis result r in direction θ
let private maxTensileStrain (sec: Section) (thetaDeg: float) (r: Ultimate.Result) (flags: bool list) =
    let d = dirOf thetaDeg
    let _, sTop = Section.extent sec d
    List.zip sec.Bars flags
    |> List.filter snd
    |> List.map (fun (b, _) ->
        let depth = sTop - sOf d { X = b.X; Y = b.Y }
        EpsCu * (depth - r.Dn) / r.Dn)
    |> List.fold max 0.0

/// Muo and φ for bending only (N = 0). classL: one flag per bar, in the order of sec.Bars.
let capacity (sec: Section) (c: Concrete) (thetaDeg: float) (classL: bool list) (mode: ClassLMode) =
    if classL.Length <> sec.Bars.Length then failwith "class flags: one per bar"
    let anyL = classL |> List.exists id
    match mode with
    | LSteelOnly when anyL ->
        let reduced k =
            { sec with Bars = List.map2 (fun (b: Bar) isL -> if isL then { b with Area = b.Area * k } else b) sec.Bars classL }
        let mutable phi = 0.85
        let mutable r = Ultimate.muo (reduced (0.65 / phi)) c thetaDeg
        let mutable i = 0
        let mutable settled = false
        while not settled && i < 30 do
            let phiNew = Design.phiBending Design.ClassN r.Kuo
            if abs (phiNew - phi) < 1e-9 then settled <- true
            else
                phi <- phiNew
                r <- Ultimate.muo (reduced (0.65 / phi)) c thetaDeg
            i <- i + 1
        let k = 0.65 / phi
        { Ult = r; Phi = phi; KL = k; EpsL = maxTensileStrain (reduced k) thetaDeg r classL }
    | _ ->
        let r = Ultimate.muo sec c thetaDeg
        let epsL = maxTensileStrain sec thetaDeg r classL
        // Table 2.2.2: Class L in tension → 0.65 for the whole section
        let phi = if epsL > 0.0 then 0.65 else Design.phiBending Design.ClassN r.Kuo
        { Ult = r; Phi = phi; KL = 1.0; EpsL = epsL }
