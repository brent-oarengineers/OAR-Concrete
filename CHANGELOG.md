# OAR Concrete – changelog

Each version that reaches `main` is built, tested and published automatically as a GitHub Release.
The calc sheet shows the version it was calculated with (under the green title bar).

## v0.2.0
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
