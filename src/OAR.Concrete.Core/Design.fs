/// AS 3600-2018 design factors that depend on the section analysis (Table 2.2.2).
/// The flexural calc sheet shows the same φ formula as a visible worksheet formula; these are here
/// for the M–N curves of the column sheet.
module OAR.Concrete.Design

open OAR.Concrete.Materials

type Ductility =
    | ClassN
    | ClassL

/// φ for bending without axial force, Table 2.2.2: Class N  1.24 − 13 kuo / 12 within 0.65 … 0.85;
/// Class L  0.65
let phiBending (ductility: Ductility) (kuo: float) =
    match ductility with
    | ClassL -> 0.65
    | ClassN -> min 0.85 (max 0.65 (1.24 - 13.0 * kuo / 12.0))

/// φ for combined bending and compression, Table 2.2.2:
/// N ≥ Nub: φ0;  N < Nub: φ0 + (φb − φ0)(1 − N / Nub). φ0 is the compression-member value
/// (argument, confirm against Table 2.2.2 before the column sheet is issued).
let phiCompression (phi0: float) (phiB: float) (n: float) (nub: float) =
    if n >= nub then phi0 else phi0 + (phiB - phi0) * (1.0 - n / nub)

/// φ for combined bending and tension: Class N interpolates from φb (N = 0) to 0.85 (N = Nuot)
let phiTension (ductility: Ductility) (phiB: float) (nTension: float) (nuot: float) =
    match ductility with
    | ClassL -> 0.65
    | ClassN -> phiB + (0.85 - phiB) * min 1.0 (abs nTension / abs nuot)

/// Squash load Nuo, Cl 10.6.2.2: α1 f'c Ac + fsy As (N); Ac = gross concrete less bars.
/// Callers pass each bar's stress at squash: min(fsy, Es εcu), so high-strength bars (e.g. threadbar) are not counted at full yield.
let squash (fc: float) (grossArea: float) (bars: (float * float) list) =
    let asTotal = bars |> List.sumBy fst
    alpha1 fc * fc * (grossArea - asTotal) + (bars |> List.sumBy (fun (a, fsy) -> a * fsy))

/// Neutral axis depth at the balanced condition: extreme tension bar at yield, dnb = do εcu / (εcu + εsy)
let dnBalanced (doo: float) (fsy: float) (es: float) = doo * EpsCu / (EpsCu + fsy / es)
