# OAR Concrete – changelog

Each version that reaches `main` is built, tested and published automatically as a GitHub Release.
The calc sheet shows the version it was calculated with (under the green title bar).

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
