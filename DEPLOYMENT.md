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
   Its folders match the server:
   - `Add-Ins\OAR Concrete\` – add-ins (64/32-bit), install script, changelog, release notes
   - `Add-Ins\OAR Concrete\Releases\v0.1.0\` – archive copy of this version's add-ins
   - `Calc Templates\OAR - RC Flexural Design.xlsx` – the calc template

Staff pick up the new add-in at their next login.

Release folder: `X:\Programs\Spreadsheets\Add-Ins\OAR Concrete\` (the script also finds it as
`\\LOA-TrueNAS\Data\Programs\Spreadsheets\Add-Ins\OAR Concrete\`, so a different drive letter on a PC doesn't matter).

## Installing on a PC (once per user)
1. Press Win+R, type `shell:startup`, Enter. Create a shortcut there to
   `X:\Programs\Spreadsheets\Add-Ins\OAR Concrete\Install-OARConcrete.cmd`.
2. Double-click the shortcut once (Excel closed). It copies the add-in to `%LOCALAPPDATA%\OAR\` and registers it
   with Excel, so it loads automatically from then on.
3. Open Excel and check `=OAR.CONCRETE.VERSION()`.

At every login the script copies a newer add-in from the NAS if there is one. It never loads the add-in from the NAS
directly, because Excel locks a loaded add-in and the NAS copy could then not be updated. A log is kept in
`%LOCALAPPDATA%\OAR\install.log`.

Brent's development PC: the script sees the existing OAR registration and leaves it alone.

## Calc records
The workbook recalculates with whatever add-in version is installed when it is opened. The issued PDF is the
record of a calc; the title block shows the add-in version it was calculated with.
