# OAR Concrete is now loaded by the shared OAR Loader (see ..\OAR Loader\). This old installer only
# forwards to the loader's installer; it no longer registers OAR Concrete with Excel.
# Shipped so the copy already on the share is overwritten with this forwarder.
$loader = Join-Path $PSScriptRoot '..\OAR Loader\Install-OARAddins.cmd'
if (Test-Path $loader) { & cmd.exe /c "`"$loader`"" }
