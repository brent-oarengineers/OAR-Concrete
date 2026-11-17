# OAR Concrete – releasing and installing

## First-time setup (once)
1. On GitHub create an empty repository `brent-oarengineers/OAR-Concrete` (private or public, no README).
2. Either let Claude push this folder to it (attach the repo to a Claude session with push access), or upload the
   folder contents through the GitHub web page ("uploading an existing file", keep the folder structure,
   including the hidden `.github` folder).
3. The first push to `main` builds, tests and publishes the current version (e.g. *OAR Concrete v0.1.0*). Then follow the steps below.


## How a release happens (Brent)
1. Claude pushes changes to a branch and opens a pull request on GitHub.
2. GitHub builds the add-in on Windows and runs every benchmark test. The pull request shows a green tick
   (passed) or red cross (failed). The built add-in can be downloaded from the run's "Artifacts" for checking.
3. You review and click **Merge**. That is the sign-off.
4. GitHub publishes a Release named after the version (e.g. *OAR Concrete v0.1.0*) with one file,
   `OAR-Concrete-v0.1.0.zip`. Releases page: https://github.com/brent-oarengineers/OAR-Concrete/releases
   (ignore the "Source code" downloads GitHub adds to every release).
5. Download the zip and extract it into `X:\Programs\Spreadsheets\`, choosing **Replace** when asked.
   Its folders match the server and follow the OAR Loader's `ADDIN-CONTRACT.md`
   (repo brent-oarengineers/OAR-Excel-Loader):
   - `Add-Ins\OAR Concrete\VERSION.txt` – the approved version (written by the same build as the add-in files)
   - `Add-Ins\OAR Concrete\OAR.Concrete-AddIn64.xll` / `OAR.Concrete-AddIn32.xll` – packed add-ins
   - `Add-Ins\OAR Concrete\Releases\v<version>\` – archive copy of this version's add-ins
   - `Add-Ins\OAR Concrete\` also holds the changelog, release notes and the old installer (now a forwarder, below)
   - `Calc Templates\OAR - RC Flexural Design.xlsx`, `OAR - RC Column Design.xlsx`, `OAR - Tilt Panel Design.xlsx` – the calc templates

Staff pick up the new version at their next Excel start.

**Rolling back:** write an older number in `Add-Ins\OAR Concrete\VERSION.txt` (its files must be in
`Releases\v<version>\`). Every PC loads that version at its next Excel start.

## Installing on a PC
OAR Concrete is loaded by the shared **OAR Loader**, the only add-in registered with Excel on OAR PCs. Install
the loader once per user (see `X:\Programs\Spreadsheets\Add-Ins\OAR Loader\`); there is no separate OAR Concrete
install. Check with `=OAR.CONCRETE.VERSION()` in Excel.

`Install-OARConcrete.cmd` (and `.ps1`) on the share now only run `..\OAR Loader\Install-OARAddins.cmd`, so an old
Startup shortcut to it installs the loader instead of registering OAR Concrete directly. If a PC still has OAR
Concrete registered directly (the old install), the loader skips it so it is never loaded twice; remove it in
File > Options > Add-ins > Excel Add-ins > Go… to hand it over to the loader.

## Calc records
The workbook recalculates with whatever add-in version is installed when it is opened. The issued PDF is the
record of a calc; the title block shows the add-in version it was calculated with.
