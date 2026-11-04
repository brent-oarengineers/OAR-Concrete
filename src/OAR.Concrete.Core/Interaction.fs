/// M–N interaction curves for one bending direction θ (groundwork for the column sheet).
///
/// Points run from pure tension (all bars yield in tension) through pure bending to the squash load.
/// Between, each point is a neutral axis depth dn with εcu at the extreme fibre (Ultimate.evaluate).
/// The last point is the squash load Nuo of Cl 10.6.2.2 (α1 f'c), which is reached in a straight line
/// from the strain-compatibility point with dn = depth / γ (rectangular block over the whole depth).
module OAR.Concrete.Interaction

open System
open OAR.Concrete.Materials
open OAR.Concrete.Geometry
open OAR.Concrete.Section

type Point =
    { /// kN, compression +
      N: float
      /// kN·m about the neutral-axis direction
      M: float
      /// mm; NaN for the pure tension and squash end points
      Dn: float
      Phi: float
      PhiN: float
      PhiM: float }

/// Curve with nPoints strain-compatibility points (plus the two ends). phi0 = φ for compression members.
let curve (sec: Section) (c: Concrete) (thetaDeg: float) (nPoints: int) (phi0: float) (ductility: Design.Ductility) =
    let d = dirOf thetaDeg
    let h = Ultimate.depthIn sec thetaDeg
    let cen = centroid sec
    let sRef = sOf d cen
    // pure bending: kuo and φb for this direction
    let bend = Ultimate.muo sec c thetaDeg
    let phiB = Design.phiBending ductility bend.Kuo
    // balanced point (outermost tension bar just yielding)
    let outer = sec.Bars |> List.minBy (fun b -> sOf d { X = b.X; Y = b.Y })
    let dnb = Design.dnBalanced bend.Do outer.Steel.Fsy outer.Steel.Es
    let nub = (Ultimate.evaluate sec c thetaDeg dnb DefaultStrips).N / 1000.0
    // ends
    let nuot = -(sec.Bars |> List.sumBy (fun b -> b.Area * b.Steel.Fsy)) / 1000.0
    let mTension = -(sec.Bars |> List.sumBy (fun b -> b.Area * b.Steel.Fsy * (sOf d { X = b.X; Y = b.Y } - sRef))) / 1e6
    let nuo = Design.squash c.Fc (gross sec).A (sec.Bars |> List.map (fun b -> b.Area, b.Steel.Fsy)) / 1000.0
    let phiFor n =
        if n >= 0.0 then Design.phiCompression phi0 phiB n nub
        else Design.phiTension ductility phiB n nuot
    let point n m dn =
        let p = phiFor n
        { N = n; M = m; Dn = dn; Phi = p; PhiN = p * n; PhiM = p * m }
    let dnMax = h / gamma c.Fc
    let dnMin = 0.02 * h
    // geometric spacing gives more points at small dn, where the curve turns fastest
    let mids =
        [ for i in 0 .. nPoints - 1 do
              let t = float i / float (max 1 (nPoints - 1))
              let dn = dnMin * Math.Pow(dnMax / dnMin, t)
              let st = Ultimate.evaluate sec c thetaDeg dn DefaultStrips
              yield point (st.N / 1000.0) (st.Mu / 1e6) dn ]
    [ yield point nuot mTension nan
      yield point 0.0 bend.Mu bend.Dn
      yield! mids |> List.filter (fun p -> p.N > 0.0)
      yield point nuo 0.0 nan ]
    |> List.sortBy (fun p -> p.N)
