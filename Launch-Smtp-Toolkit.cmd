@echo off
REM ===========================================================================
REM  SMTP Toolkit - launcher
REM
REM  -STA is REQUIRED. WinForms will not run reliably in a multi-threaded
REM  apartment, and the failure is intermittent rather than obvious.
REM
REM  -WindowStyle Hidden is deliberately NOT used here: it hides the form as
REM  well as the console. The script hides its own console window in-process
REM  instead. Pass -NoHideConsole if you want to see the console for debugging.
REM
REM  -ExecutionPolicy Bypass is kept only for this launcher, so the tool still
REM  runs on an unknown machine and on a file carrying a mark-of-the-web.
REM ===========================================================================
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0Smtp-Toolkit.ps1" %*
if errorlevel 1 (
  echo.
  echo The toolkit exited with an error. Re-run with -NoHideConsole to see it:
  echo   powershell.exe -NoProfile -STA -File "%~dp0Smtp-Toolkit.ps1" -NoHideConsole
  echo.
  pause
)
