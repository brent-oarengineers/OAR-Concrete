/// AS 3600-2018 material properties: concrete stress-strain laws and reinforcement.
/// Sign convention used everywhere in the engine: compression is POSITIVE (strain, stress, force).
module OAR.Concrete.Materials

open System

/// Elastic modulus of reinforcement, MPa (Cl 3.2.2)
[<Literal>]
let EsDefault = 200000.0

/// Concrete strain at the extreme compression fibre at ultimate (Cl 8.1.2 / 8.1.3)
[<Literal>]
let EpsCu = 0.003

// ---------------------------------------------------------------------------------------------
// Concrete properties
// ---------------------------------------------------------------------------------------------

let private fcTable = [| 20.0; 25.0; 32.0; 40.0; 50.0; 65.0; 80.0; 100.0 |]
let private ecTable = [| 24000.0; 26700.0; 30100.0; 32800.0; 34800.0; 37400.0; 39600.0; 42200.0 |]

/// f'c must be inside the range AS 3600 covers
let checkFc (fc: float) =
    if Double.IsNaN fc || fc < 20.0 || fc > 100.0 then
        failwithf "f'c = %g MPa is outside AS 3600 (20 to 100 MPa)" fc

/// Ec, MPa: Table 3.1.2, linear between the tabulated grades (density 2400 kg/m³)
let ecTable312 (fc: float) =
    checkFc fc
    let i = fcTable |> Array.findIndexBack (fun x -> x <= fc) |> min (fcTable.Length - 2)
    let t = (fc - fcTable.[i]) / (fcTable.[i + 1] - fcTable.[i])
    ecTable.[i] + t * (ecTable.[i + 1] - ecTable.[i])

/// α2 of the rectangular stress block, Cl 8.1.3:  0.85 − 0.0015 f'c, not less than 0.67
let alpha2 (fc: float) = max 0.67 (0.85 - 0.0015 * fc)

/// γ of the rectangular stress block, Cl 8.1.3:  0.97 − 0.0025 f'c, not less than 0.67
let gamma (fc: float) = max 0.67 (0.97 - 0.0025 * fc)

/// α1 for the squash load, Cl 10.6.2.2:  1.0 − 0.003 f'c, within 0.72 … 0.85
let alpha1 (fc: float) = min 0.85 (max 0.72 (1.0 - 0.003 * fc))

/// Characteristic flexural tensile strength, Cl 3.1.1.3:  0.6 √f'c
let fctf (fc: float) = 0.6 * sqrt fc

/// Which compressive stress-strain law is used at ultimate
type ConcreteModel =
    /// Rectangular stress block α2 f'c over γ dn (Cl 8.1.3)
    | Rectangular
    /// Curvilinear curve of Cl 3.1.4 (commentary form, as Brent's Circular_Pile_AS3600.py):
    /// σ = σcp n η / (n − 1 + η^(nk)), σcp = 0.9 f'c, η = ε/εcp
    | Curvilinear
    /// Parabola–rectangle (Hognestad form, as RAM Concept): σ = 0.9 f'c [2η − η²], η = ε/ε0, up to
    /// ε0 = 2 × 0.9 f'c / Ec (initial slope = Ec), then 0.9 f'c to εcu
    | Parabolic

let parseModel (text: string) =
    let t = if isNull text then "" else text.Trim().ToUpperInvariant()
    if t = "" || t.StartsWith "RECT" then Rectangular
    elif t.StartsWith "CURV" then Curvilinear
    elif t.StartsWith "PARA" then Parabolic
    else failwithf "stress block '%s' not recognised (Rectangular, Curvilinear, Parabolic)" text

type Concrete =
    { Fc: float
      /// Elastic modulus, MPa (used by the curvilinear law and the cracked analysis)
      Ec: float
      Model: ConcreteModel }

/// Concrete with Ec from Table 3.1.2 unless given
let concrete (fc: float) (ec: float option) (model: ConcreteModel) =
    checkFc fc
    let e = defaultArg ec (ecTable312 fc)
    if e <= 0.0 then failwith "Ec must be positive"
    { Fc = fc; Ec = e; Model = model }

/// Parameters of the curvilinear law (Cl 3.1.4): σcp, εcp, n, k (descending branch)
let curvilinearParameters (c: Concrete) =
    let scp = 0.9 * c.Fc
    let ecp = 4.11 * Math.Pow(scp, 0.75) / c.Ec
    let ecpMod = scp / ecp
    let n = c.Ec / (c.Ec - ecpMod)
    let k = max (0.67 + scp / 62.0) 1.0
    scp, ecp, n, k

/// Strain at the peak of the parabolic law: ε0 = 2 σcp / Ec, σcp = 0.9 f'c (initial tangent = Ec)
let parabolaEps0 (c: Concrete) = 2.0 * 0.9 * c.Fc / c.Ec

/// Concrete stress for a strain (compression +). Tension carries no stress.
let stress (c: Concrete) (eps: float) =
    if eps <= 0.0 then 0.0
    else
        match c.Model with
        | Rectangular ->
            // α2 f'c acts where the strain is at least εcu (1 − γ), i.e. within γ dn of the extreme fibre
            if eps >= EpsCu * (1.0 - gamma c.Fc) * (1.0 - 1e-9) then alpha2 c.Fc * c.Fc else 0.0
        | Parabolic ->
            let fm = 0.9 * c.Fc
            let e0 = parabolaEps0 c
            if eps < e0 then
                let r = eps / e0
                fm * (2.0 * r - r * r)
            else fm
        | Curvilinear ->
            let scp, ecp, n, kDesc = curvilinearParameters c
            let eta = eps / ecp
            let k = if eps <= ecp then 1.0 else kDesc
            scp * n * eta / (n - 1.0 + Math.Pow(eta, n * k))

/// Strains at which the law has a step or kink; the integrator puts a strip boundary there
let breakStrains (c: Concrete) =
    match c.Model with
    | Rectangular -> [ EpsCu * (1.0 - gamma c.Fc) ]
    | Parabolic -> [ parabolaEps0 c ]
    | Curvilinear ->
        let _, ecp, _, _ = curvilinearParameters c
        [ ecp ]

// ---------------------------------------------------------------------------------------------
// Reinforcement
// ---------------------------------------------------------------------------------------------

/// Elastic – perfectly plastic reinforcement (Cl 3.2)
type Steel = { Fsy: float; Es: float }

let steel (fsy: float) = { Fsy = fsy; Es = EsDefault }

/// Steel stress for a strain (compression +), elastic – perfectly plastic
let steelStress (s: Steel) (eps: float) = max (-s.Fsy) (min s.Fsy (s.Es * eps))
