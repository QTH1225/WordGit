@echo off
setlocal
rem A module path inherited from PowerShell 7 breaks Windows PowerShell 5.1's built-in modules.
set "PSModulePath="
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
set "WordGitExitCode=%errorlevel%"
echo.
pause
exit /b %WordGitExitCode%
