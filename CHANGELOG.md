# OAR Concrete – changelog

Each version that reaches `main` is built, tested and published automatically as a GitHub Release.
The calc sheet shows the version it was calculated with (under the green title bar).

## v0.5.0
- New **OAR** ribbon tab with **Export calcs to PDF**: exports every visible calc tab (OAR green tab colour 465A52)
  in tab order into one PDF, using each tab's print area and page setup. Data tabs (other colours) and hidden tabs are
  skipped. Defaults to "<workbook name> - Calcs.pdf" next to the workbook and opens the PDF when done.
  Works in any OAR calc workbook, not just the concrete templates. The tab is shared (idQ `oar:OarTab`) so OAR Steel /
  OAR Timber can add groups to the same tab later.

## v0.4.5
- Column sheet: new check **εs ≤ εsu** – largest strain in the outermost tension bar at capacity (over the load cases)
  against the fracture strain, Class N 5 % / Class L 1.5 % (AS/NZS 4671 uniform elongation; 1.5 % if any bar is Class L).
  The capacity table shows εs for each case. `OAR.CONCRETE.COLUMN` spills a 10th value, εs.

## v0.4.4
- Column sheet: custom bars can have their own yield strength and Es. The column Bars tab has an **Es** column and
  high-strength threadbar rows **TB26.5 – TB47** (950 MPa, Es 205 GPa – confirm with the supplier); add rows for any
  other bar. The Bar Schedule shows Es and passes it to the add-in.
- Every bar's stress comes from its own strain (elastic – perfectly plastic, own fsy and Es), so a central threadbar only
  reaches 950 MPa where the strain allows. At the squash load each bar is limited to min(fsy, Es εcu)
  (615 MPa for threadbar; no change for 500 MPa bars).

## v0.4.3
- Column Bar Schedule: the 10 custom bar rows (C1–C10, input cells) now sit at the top, above the calculated typical
  layout, so typing into the add-in array ("You can't change part of an array") no longer happens by accident.

## v0.4.2
- Column sheet: new **Interaction Curve – Key Points** table on page 1 (φN and φM about each axis at squash, decompression,
  balanced, pure bending and pure tension, with φ at each point). Load cases now start page 2.
- `OAR.CONCRETE.COLUMNAXIS` spills 11 values: adds the decompression point Nud, Mud (zero strain at the outermost
  tension bar, dn = do) and φ at pure tension.

## v0.4.1
- Column sheet charts cleaned up: no chart titles or legends (caption above each chart, colour key in the note),
  thin blue / orange lines, light major and minor gridlines, axes at the left and bottom, square border.
- The flexural and column templates are attached to each release as separate downloads.

## v0.4.0
- **New template: OAR - RC Column Design.xlsx** – rectangular and circular columns to AS 3600-2018: typical perimeter
  bar layout plus custom bars, braced slenderness (Cl 10.3 – 10.5), biaxial capacity by strain compatibility along the
  load direction and the Cl 10.6.4 interaction, M–N charts for both axes, moment contour and section plot.
- New add-in functions: `OAR.CONCRETE.BARLAYOUT`, `COLUMN`, `COLUMNAXIS`, `MNCOLUMN`, `MCONTOUR`.
- House-style helpers shared between templates (`template/oar_style.py`).

## v0.3.2
- Single mat position: *Reinforcement layout* now offers **Single mat – central**, **– at top cover** and
  **– at bottom cover**. The mat (Top column) sits at D/2 or at the entered top / bottom cover; extras still go on top of it
  (Top column) or under it (Bottom column).

## v0.3.1
- Class S reinforcement: S10–S32 (fsy 250 MPa, deformed, ductility Class N, same areas as N bars) on the Bars tab.

## v0.3.0
- **Single layer of reinforcement**: bar "None" for a face with no mat (e.g. slab on ground with one mat at cover).
  A direction whose tension face has no steel uses the other face's layers for φ class, minimum strength and crack
  control; the deemed-to-comply Ast.min route is then not used.
- **Reinforcement layout: Single central mat** – one mat at D/2 for pool shells and slabs on ground, with extras on top
  of the mat (Top column) and/or under it (Bottom column). The mat serves both directions.
- New *Tension Reinforcement by Direction* table (which bars count as tension steel for each direction, and why).
- Example S2: 150 pool floor, SL92 central + 5 N12 under the mat.
- Template renamed **OAR - RC Flexural Design.xlsx** (delete the old *RC Flexural Design - Base.xlsx*).
- Add-in version check: each calc tab records the add-in version it was built for; an older or missing add-in shows a
  red message under the title bar and UPDATE ADD-IN in the status box (on screen and printed).

## v0.2.0
- Crack control to AS 3600-2018 Amdt 2: σs.max from Tbl 8.6.2.2(A)/(B) (beams) and 9.5.2.1(A)/(B) (slabs, Ds ≤ / > 300)
  for w'max 0.2 / 0.3 / 0.4 mm (new input, default 0.3), new *Crack Control* data tab. Spacing ignores bars under half
  the largest diameter. *Legacy* keeps the previous OAR sheet limits. S1 now 108 %, B1 141 % (crack control governs).
- Benchmarked against RAM Concept (AS 3600-2018 Amdt 2 Final Design Check, two slab sections, both directions):
  Muo, dn and φMu agree to 0.1 %, cracked bar stresses to 0.6 % (tests added).
- **Parabolic** stress block now peaks at ε0 = 2 × 0.9 f'c / Ec (Hognestad form, as RAM Concept; was a fixed
  0.002). Parabolic capacities change by less than 0.1 %; rectangular and curvilinear are unchanged.

## v0.1.0
- New section engine (OAR.Concrete.Core): polygon / circle / void regions with individual bars; strain
  compatibility at any bending direction and axial force with rectangular (Cl 8.1.3), curvilinear (Cl 3.1.4) or
  parabolic concrete; cracked elastic section; φ (Table 2.2.2); M–N curves.
- Worksheet functions `OAR.CONCRETE.ULT`, `OAR.CONCRETE.CRACKED`, `OAR.CONCRETE.MNCURVE`, `OAR.CONCRETE.VERSION`.
- Calc template *RC Flexural Design - Base.xlsx*: slab or beam, both directions, ULS, minimum strength, cracked
  bar stresses (Ms*, Ms1*), bar spacing and slab shrinkage steel. Examples S1 (previous slab sheet) and B1.
- Class L reinforcement: *Whole section 0.65* (default) or *L steel only* (Class L areas × 0.65/φ, φ iterated),
  with a Class L strain check (≤ 1.5 %) in L steel only mode.
- 60 benchmark tests incl. Brent's circular pile script (106.345 kNm).
