# OAR Concrete

Excel add-in and calc templates for reinforced concrete design to AS 3600-2018, in the OAR house calc style.
Same build, release and install arrangement as OAR Steel and OAR Timber.

v0.1 has the **RC flexural member** calc (slab strip or beam, sagging and hogging, ULS and cracked SLS
stresses). The engine is general enough for the next sheet, **columns** (rectangular and circular, M–N curves,
biaxial bending): see *Groundwork for the column sheet* below.

## What is where

| Folder | Contents |
|---|---|
| `src/OAR.Concrete.Core` | Section engine (F#, netstandard2.0). No Excel code. |
| `src/OAR.Concrete.Excel` | Excel-DNA add-in (net48): converts sheet ranges to the engine and back. One .xll per bitness. |
| `tests/OAR.Concrete.Tests` | Benchmark tests (plain console runner): closed-form cases, `tools/reference_check.py`, Brent's pile script. |
| `tools/` | `reference_check.py` – independent scipy re-implementation used to produce benchmark values; `Circular_Pile_AS3600.py` – Brent's original pile script (curvilinear benchmark). |
| `template/` | `RC Flexural Design - Base.xlsx` and the script that builds it (`python template/build_template.py`). |
| `deploy/` | `Install-OARConcrete.cmd/.ps1` – per-user install / update at login. |

## How the engine works (in plain terms)

`Core` is built in layers, each file using only the ones above it:

1. **Materials.fs** – f'c checks, Ec (Table 3.1.2, interpolated), α2 and γ (Cl 8.1.3), α1 (Cl 10.6.2.2), f'ct.f, and three
   concrete stress–strain laws, all with **compression positive**:
   - *Rectangular*: α2 f'c wherever the strain is at least εcu(1 − γ) – that is exactly "α2 f'c over γ dn from the
     extreme fibre", written as a stress–strain law so the same integrator handles any shape.
   - *Curvilinear*: the Cl 3.1.4 curve from your `Circular_Pile_AS3600.py` (σcp = 0.9 f'c, εcp = 4.11 σcp^0.75 / Ec,
     n = Ec / (Ec − Ecp), k = 1 rising, max(0.67 + σcp/62, 1) falling).
   - *Parabolic*: parabola to 0.9 f'c at ε0 = 0.002, flat to εcu = 0.003.
   - Steel: elastic – perfectly plastic, Es 200 GPa.
2. **Geometry.fs** – concrete regions: polygons, circles, and voids (for hollow sections later). The analysis cuts
   the compression zone into thin bands parallel to the neutral axis. A polygon band is found by clipping the polygon
   with two lines (Sutherland–Hodgman) and the shoelace formulas; a circle band uses the exact circular-segment
   formulas (your hand derivation: chord 2√(2Rx − x²) integrated).
3. **Section.fs** – a section = regions + individual bars (x, y, area, fsy, Es). Builders for a b × D rectangle and a
   circle. `strips` puts a band boundary wherever the stress law has a step or kink (the edge of the rectangular
   block, εcp, ε0), so the rectangular block is integrated exactly.
4. **Ultimate.fs** – strain compatibility. For a direction θ (90° = top in compression, −90° = bottom) and a
   neutral-axis depth dn, the extreme fibre is at 0.003; concrete stresses are summed band by band, bars get
   σs − σc (displaced concrete) when in compression. `solveDn` bisects on dn until the internal force equals the
   applied N (0 for a slab or beam). Reports dn, kuo (to the outermost tension bar), ku (to the tension resultant),
   strains and stresses, forces, lever arm and Mu.
5. **Service.fs** – cracked elastic section: linear concrete, no tension, (Es − Ec) at compression bars (= the usual
   (n − 1)Asc). With N = 0 the neutral axis comes from "first moment = 0", then curvature from M. Gives dn, Icr, bar
   stresses and the extreme-fibre concrete stress. Pass Ec / (1 + φcc) for a long-term modular ratio.
6. **Design.fs** – φ (Table 2.2.2) for bending, bending + compression and bending + tension, squash load, balanced dn.
7. **Interaction.fs** – M–N curve for one direction (pure tension → pure bending → … → squash), with φ.

## How the calc sheet uses the add-in

Code checks are visible worksheet formulas: bar areas and depths (the *Reinforcement Layers* table), α2, γ, Ec, n,
φ, minimum strength, crack-control limits, bar spacing and shrinkage steel. Only the section analysis comes from
the add-in:

```
=OAR.CONCRETE.ULT(shape, dims, bars, fc, model, sense, [N], [Ec])      spills 13 values down
=OAR.CONCRETE.CRACKED(shape, dims, bars, Ec, M, sense, [N])            spills 5 values down
```

| Argument | |
|---|---|
| `shape` | `Rect` (dims = b, D) or `Circle` (dims = D) |
| `bars` | rows of **As (mm²) \| depth from top face (mm) \| fsy (MPa) \| [x from left, mm] \| [Es]**; zero / blank rows skipped. On the sheet this is the Reinforcement Layers table (one row per layer). |
| `model` | `Rectangular`, `Curvilinear` or `Parabolic` |
| `sense` | `Sagging` (top in compression), `Hogging`, or an angle in degrees for the column sheet |
| `N` | axial force, kN, compression + (blank = 0) |

`ULT` spills: dn, kuo, ku, do, d, εst, σst, σsc, Cc, Cs, T, z, Muo (mm, MPa, kN, kNm).
`CRACKED` spills: dn, Icr (×10⁶ mm⁴), σst (largest bar tension), σsc, σc.
Also `OAR.CONCRETE.MNCURVE(...)` (rows N, M, dn, φ, φN, φM – preview for the column sheet) and
`OAR.CONCRETE.VERSION()`. Errors show as `#OAR! …` text, and the summary status then reads INPUT ERROR.

### The flexural calc tab

- **Member type** Slab / Beam. Slab: b = strip width (1000 per metre), mats as *bar @ spacing*, cover to the bar.
  Beam: *number × bar*, cover to the ligature plus the ligature size.
- Each face: main bars plus extras (number within b, bar, layer 1 alongside or layer 2 inside with a clear gap).
  Mesh wires L6–L12 are on the Bars tab (Class L; SL92 = L9 @ 200).
- ULS both directions with the selected stress block; the other block (rectangular ↔ curvilinear) is shown as an
  info row for comparison.
- φ from kuo, Class N formula; 0.65 if any bar on the tension face is Class L.
- Minimum strength: Muo ≥ 1.2 Z f'ct.f, or deemed-to-comply Ast ≥ αb (D/d)² f'ct.f/fsy · bd (OK* when only the
  deemed route passes). αb 0.20 beams; slabs 0.19 on beams/walls, 0.24 on columns.
- SLS: σscr under Ms* ≤ σs.max and σscr.1 under Ms1* ≤ 0.8 fsy, both faces; bar spacing ≤ 300 (beam) or
  min(2D, 300) (slab); slab shrinkage & temperature steel as the previous OAR slab sheet.

## Checking the template in Excel

With the add-in loaded the example tabs should read:

| Tab | Member | Max utilisation | Governs |
|---|---|---|---|
| S1 | 175 slab, f'c 32, SL92 T&B + 3 N16 top (previous slab sheet example) | 102 % FAIL | hogging: φMuo 36.7 vs M* 37.3 kNm (φ = 0.65, Class L mesh in the tension face) |
| B1 | 1200 × 600 band beam, f'c 40, 7 + 4 N24 top (2 layers), 6 N24 bottom | 128 % FAIL | hogging σscr 332 MPa vs 260 MPa (legacy table B, s = 175) |

Key values (engine = `tools/reference_check.py` to 4 figures): S1 Muo 32.73 / 56.43 kNm, dn 25.25 mm both ways,
σscr 158 / 279 MPa; B1 Muo 751.2 / 1176.9 kNm, σscr 75.5 / 331.5 MPa.

## What is checked, and against what

`dotnet run --project tests/OAR.Concrete.Tests -c Release` (48 checks):

| Area | Benchmark | Tolerance |
|---|---|---|
| Rectangular block, singly reinforced | dn = Ast fsy / (α2 f'c γ b), Muo = T(d − γdn/2) | 1e-7 |
| Rect / parabolic / curvilinear, doubly reinforced beam, sag and hog | `reference_check.py` (scipy quad + brentq) | 2e-5 |
| Cracked section, singly reinforced | kd, Icr = bx³/3 + nAst(d − x)², σs, σc closed form | 1e-5 |
| Cracked section with (n − 1)Asc, both directions | `reference_check.py` | 1e-5 |
| Circular pile, curvilinear | Brent's `Circular_Pile_AS3600.py`: Mu 106.345 kNm, dn 85.81, kuo 0.2524 | 5e-5 |
| Circle vs 720-gon of equal area; 8-bar pile rotated 45° | same Mu | 2e-4 / 1e-6 |
| M–N curve | squash α1 f'c Ac + fsy As, pure tension, pure-bending point = Muo, solve for N = 1500 kN | 1e-9 |

## Open items to confirm before relying on them

- **Crack-control stress limits.** The sheet uses the *legacy* limits from the previous OAR slab sheet (fits to the
  AS 3600-2009 tables). AS 3600-2018 Amdt 2 replaced these with crack-width tables (Tbl 8.6.2.2 / 9.5.2.1) that are
  not built in – enter σs.max on the sheet from your copy, or ask for those tables to be added.
- **Class L rule.** φ = 0.65 if *any* tension-face bar is Class L (conservative where N bars carry most of the
  steel, e.g. S1).
- **Shrinkage steel** uses the previous sheet's interpolation on b D (the old sheet used b d).
- **φ for compression members** (MNCURVE default φ0 = 0.65): confirm against Table 2.2.2 before the column sheet.
- Not covered: shear, torsion, deflection, detailing.

## Groundwork for the column sheet

Already in the engine and tested: any bending direction θ (biaxial: run θ and read Mx, My), any axial force
(`Ultimate.analyse … N`), circular sections (exact) and polygons (any outline, voids for hollow sections),
individual bars at x, y, the squash load, Nub and φ interpolation, and `Interaction.curve`. Still to do for that
sheet: the bar-layout builder for rectangular columns (bars per face), solving θ so the capacity vector lines up
with (Mx*, My*), slenderness moment magnification (Cl 10.4/10.5), and the M–N chart.

## Building and testing

```
dotnet run --project tests/OAR.Concrete.Tests -c Release
dotnet build src/OAR.Concrete.Excel -c Release
python template/build_template.py
```

The packed add-ins are `src/OAR.Concrete.Excel/bin/Release/net48/publish/OAR.Concrete.Excel-AddIn64-packed.xll`
(and the 32-bit one). GitHub Actions does the same on every push and publishes a release from `main`
(see `DEPLOYMENT.md`).
