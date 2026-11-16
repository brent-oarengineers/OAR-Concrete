@echo off
rem OAR Concrete is now loaded by the shared OAR Loader (see ..\OAR Loader\). This old installer only
rem forwards to the loader's installer, so existing Startup shortcuts keep working. It no longer
rem registers OAR Concrete with Excel.
call "%~dp0..\OAR Loader\Install-OARAddins.cmd" %*
