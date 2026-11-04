# OAR Concrete – changelog

Each version that reaches `main` is built, tested and published automatically as a GitHub Release.
The calc sheet shows the version it was calculated with (under the green title bar).

## v0.1.0
- New section engine (OAR.Concrete.Core): polygon / circle / void regions with individual bars; strain
  compatibility at any bending direction and axial force with rectangular (Cl 8.1.3), curvilinear (Cl 3.1.4) or
  parabolic concrete; cracked elastic section; φ (Table 2.2.2); M–N curves.
- Worksheet functions `OAR.CONCRETE.ULT`, `OAR.CONCRETE.CRACKED`, `OAR.CONCRETE.MNCURVE`, `OAR.CONCRETE.VERSION`.
- Calc template *RC Flexural Design - Base.xlsx*: slab or beam, both directions, ULS, minimum strength, cracked
  bar stresses (Ms*, Ms1*), bar spacing and slab shrinkage steel. Examples S1 (previous slab sheet) and B1.
- 48 benchmark tests incl. Brent's circular pile script (106.345 kNm).
