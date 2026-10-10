@echo off
chcp 936 >nul 2>&1
setlocal enabledelayedexpansion

rem ============================================================
rem   Right-Click New Menu Template Fix Tool v2.9
rem   Step 1: Thorough cleanup (MS Office + WPS) + user verify
rem   Step 2: Detect Office installations + version display
rem   Step 3: Fix (MS Office NullFile / WPS FileName)
rem   Step 4: Verify + restart Explorer
rem   Both installed: user selects 1=MS Office / 2=WPS / C=Cancel
rem   Only WPS: auto WPS fix + restart Explorer
rem   Only MS Office: auto MS Office fix + restart Explorer
rem   Extensions: .xlsx / .pptx / .docx
rem   Also removes .ppt/.doc/.xls (97-2003) from New menu
rem ============================================================

set SILENT=0
set RESTART_EXPLORER=0
set USER_CHOICE=

:parse_args
if "%~1"=="" goto args_done
if /i "%~1"=="/silent" set SILENT=1& shift& goto parse_args
if /i "%~1"=="/quiet" set SILENT=1& shift& goto parse_args
if /i "%~1"=="/restart" set RESTART_EXPLORER=1& shift& goto parse_args
if /i "%~1"=="/mode" set USER_CHOICE=%~2& shift& shift& goto parse_args
shift
goto parse_args
:args_done

net session >nul 2>&1
if errorlevel 1 goto no_admin
goto admin_ok

:no_admin
echo [ERROR] Please run this script as Administrator
echo.
if "%SILENT%"=="0" pause
exit /b 1

:admin_ok

echo ============================================================
echo   Right-Click New Menu Template Fix Tool v2.9
echo ============================================================
echo.

rem ============================================================
rem Step 1: Thorough cleanup
rem ============================================================
echo [Step 1/Clean] Cleaning all right-click New menu entries...
set CLEAN_COUNT=0

rem Clean MS Office modern formats (extension level + ProgID level)
call :clean_ext ".xlsx" "Excel.Sheet.12"
call :clean_ext ".pptx" "PowerPoint.Show.12"
call :clean_ext ".docx" "Word.Document.12"

rem Clean MS Office 97-2003 formats
call :clean_ext ".xls" "Excel.Sheet.8"
call :clean_ext ".ppt" "PowerPoint.Show.8"
call :clean_ext ".doc" "Word.Document.8"

rem Clean WPS-specific format classes (et/wps/dps)
for %%e in (et wps dps) do (
  reg delete "HKCR\%%e" /f >nul 2>&1
  if not errorlevel 1 (
    echo   - HKCR\%%e removed
    set /a CLEAN_COUNT+=1
  )
)

rem Clean WPS-specific extension classes (.et/.wps/.dps)
for %%e in (.et .wps .dps) do (
  reg delete "HKCR\%%e" /f >nul 2>&1
  if not errorlevel 1 (
    echo   - HKCR\%%e removed
    set /a CLEAN_COUNT+=1
  )
)

echo   Total cleaned: %CLEAN_COUNT% items
echo.

rem ============================================================
rem Step 1b: Refresh cache + restart Explorer after cleanup
rem ============================================================
echo   Refreshing ShellNew cache...
call :cache_clean
echo   Cache refreshed (all Office extensions removed from cache).
echo   Restarting Explorer to apply cleanup...
taskkill /f /im explorer.exe >nul 2>&1
timeout /t 1 /nobreak >nul 2>&1
start explorer.exe
echo   Explorer restarted.
echo.

rem ============================================================
rem Step 1c: User verification loop
rem ============================================================
if "%SILENT%"=="1" goto skip_clean_verify

:clean_verify_loop
echo [Step 1/Verify] Please right-click and check New menu now.
echo   These should NOT appear in the New menu:
echo     .xlsx / .pptx / .docx / .xls / .ppt / .doc
echo.
choice /c YN /n /m "  All removed? [Y=proceed, N=re-clean]: "
if errorlevel 2 goto recheck_clean
goto clean_verify_done

:recheck_clean
echo.
echo   Re-checking registry for remaining ShellNew entries...
set RECHECK_FOUND=0
call :recheck_ext ".xlsx" "Excel.Sheet.12"
call :recheck_ext ".pptx" "PowerPoint.Show.12"
call :recheck_ext ".docx" "Word.Document.12"
call :recheck_ext ".xls" "Excel.Sheet.8"
call :recheck_ext ".ppt" "PowerPoint.Show.8"
call :recheck_ext ".doc" "Word.Document.8"

if "%RECHECK_FOUND%"=="0" (
  echo   No remaining ShellNew entries in registry.
  echo   Proceeding to fix step...
  goto clean_verify_done
) else (
  echo   Found and cleaned %RECHECK_FOUND% more entries.
  echo   Refreshing cache and restarting Explorer...
  call :cache_clean
  taskkill /f /im explorer.exe >nul 2>&1
  timeout /t 1 /nobreak >nul 2>&1
  start explorer.exe
  echo   Explorer restarted.
  echo.
  goto clean_verify_loop
)

:skip_clean_verify
echo [Step 1/Verify] Silent mode: skip user verification.
echo.

:clean_verify_done
echo.

rem ============================================================
rem Step 2: Detect Office installations
rem ============================================================
echo [Step 2/Detect] Detecting Office installation...

set HAS_MSOFFICE=0
set HAS_WPS=0
set "MSOFFICE_PATH="
set "MSOFFICE_VERSION="
set "WPS_INSTALL_PATH="
set "WPS_VERSION="
set "WPS_TEMPLATE_DIR="

rem Detect MS Office
for /f "skip=2 tokens=2,*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE" /ve 2^>nul') do set "MSO_WORD=%%b"
if not defined MSO_WORD for /f "skip=2 tokens=2,*" %%a in ('reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE" /ve 2^>nul') do set "MSO_WORD=%%b"
if defined MSO_WORD if exist "%MSO_WORD%" (
  set HAS_MSOFFICE=1
  for %%d in ("%MSO_WORD%") do set "MSOFFICE_PATH=%%~dpd"
  for /f "skip=2 tokens=2,*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v "VersionToReport" 2^>nul') do set "MSOFFICE_VERSION=%%b"
  if not defined MSOFFICE_VERSION for /f "skip=2 tokens=2,*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot" /v "Path" 2^>nul') do set "MSOFFICE_VERSION=16.0"
  echo   - MS Office detected (version %MSOFFICE_VERSION%)
)

rem Detect WPS
call :detect_wps
if "%HAS_WPS%"=="1" echo   - WPS Office detected (version %WPS_VERSION%)

rem Exit if nothing detected
if "%HAS_MSOFFICE%"=="0" if "%HAS_WPS%"=="0" (
  echo   [WARNING] No Office software detected, exiting.
  echo.
  if "%SILENT%"=="0" pause
  exit /b 0
)
echo.

rem ============================================================
rem Step 2b: Determine fix mode + user choice if both installed
rem ============================================================
set FIX_MODE=

if "%HAS_MSOFFICE%"=="1" if "%HAS_WPS%"=="0" (
  set FIX_MODE=msoffice
  echo   Only MS Office installed.
)

if "%HAS_WPS%"=="1" if "%HAS_MSOFFICE%"=="0" (
  set FIX_MODE=wps
  echo   Only WPS Office installed. Auto WPS fix mode.
)

if "%HAS_MSOFFICE%"=="1" if "%HAS_WPS%"=="1" (
  echo   Both MS Office and WPS Office detected.
  echo.
  echo   Detected versions:
  echo     MS Office: %MSOFFICE_VERSION%
  echo     WPS Office: %WPS_VERSION%
  echo.
  if "%SILENT%"=="1" (
    if "%USER_CHOICE%"=="1" (
      set FIX_MODE=msoffice
      echo   Silent mode: MS Office fix.
    ) else if "%USER_CHOICE%"=="2" (
      set FIX_MODE=wps
      echo   Silent mode: WPS fix.
    ) else (
      set FIX_MODE=wps
      echo   Silent mode: default WPS fix.
    )
  ) else (
    call :choose_mode
    if errorlevel 3 goto cancel_exit
  )
)
goto fix_start

:choose_mode
echo   Please choose fix mode:
echo     [1] Microsoft Office fix new (NullFile)
echo     [2] WPS fix new (FileName with template)
echo     [C] Cancel and exit
echo.
choice /c 12C /n /m "  Enter choice [1/2/C]: "
if errorlevel 3 exit /b 3
if errorlevel 2 (
  set FIX_MODE=wps
  echo   WPS fix mode selected.
  exit /b 2
)
set FIX_MODE=msoffice
echo   MS Office fix mode selected.
exit /b 1

:cancel_exit
echo.
echo   Cancelled by user.
if "%SILENT%"=="0" pause
exit /b 0

:fix_start
echo.

rem ============================================================
rem Step 3: Fix
rem ============================================================
echo [Step 3/Fix] Fixing right-click New menu...
set MODIFY_COUNT=0

if "%FIX_MODE%"=="msoffice" echo   Mode: MS Office (NullFile)
if "%FIX_MODE%"=="wps" echo   Mode: WPS (FileName)

if not "%FIX_MODE%"=="msoffice" goto skip_msfix

rem === MS Office mode: NullFile under ProgID ===
call :fix_msoffice ".xlsx" "Excel.Sheet.12"
if errorlevel 1 set /a MODIFY_COUNT+=1
call :fix_msoffice ".pptx" "PowerPoint.Show.12"
if errorlevel 1 set /a MODIFY_COUNT+=1
call :fix_msoffice ".docx" "Word.Document.12"
if errorlevel 1 set /a MODIFY_COUNT+=1

rem Set display names + remove FriendlyTypeName
echo   Setting display names...
reg add "HKCR\Excel.Sheet.12" /ve /d "Microsoft Excel" /f >nul 2>&1
reg add "HKCR\PowerPoint.Show.12" /ve /d "Microsoft PowerPoint" /f >nul 2>&1
reg add "HKCR\Word.Document.12" /ve /d "Microsoft Word" /f >nul 2>&1
reg delete "HKCR\Excel.Sheet.12" /v "FriendlyTypeName" /f >nul 2>&1
reg delete "HKCR\PowerPoint.Show.12" /v "FriendlyTypeName" /f >nul 2>&1
reg delete "HKCR\Word.Document.12" /v "FriendlyTypeName" /f >nul 2>&1
echo   Display names set.
goto fix_done
:skip_msfix

if not "%FIX_MODE%"=="wps" goto fix_done
call :prepare_wps_templates

rem === WPS mode: FileName with template path ===
call :fix_wps ".xlsx" "Excel.Sheet.12" "%XLSX_TEMPLATE%"
if errorlevel 1 set /a MODIFY_COUNT+=1
call :fix_wps ".pptx" "PowerPoint.Show.12" "%PPTX_TEMPLATE%"
if errorlevel 1 set /a MODIFY_COUNT+=1
call :fix_wps ".docx" "Word.Document.12" "%DOCX_TEMPLATE%"
if errorlevel 1 set /a MODIFY_COUNT+=1

rem Set display names
echo   Setting display names...
reg add "HKCR\Excel.Sheet.12" /ve /d "Microsoft Excel" /f >nul 2>&1
reg add "HKCR\PowerPoint.Show.12" /ve /d "Microsoft PowerPoint" /f >nul 2>&1
reg add "HKCR\Word.Document.12" /ve /d "Microsoft Word" /f >nul 2>&1
reg delete "HKCR\Excel.Sheet.12" /v "FriendlyTypeName" /f >nul 2>&1
reg delete "HKCR\PowerPoint.Show.12" /v "FriendlyTypeName" /f >nul 2>&1
reg delete "HKCR\Word.Document.12" /v "FriendlyTypeName" /f >nul 2>&1
echo   Display names set.

:fix_done
echo.
echo   Updating ShellNew cache...
call :cache_update
echo   Sending SHChangeNotify...
call :send_notify
echo.

rem ============================================================
rem Step 4: Verify + restart Explorer
rem ============================================================
echo [Step 4/Verify] Verifying fix results...
set VERIFY_OK=0

if "%FIX_MODE%"=="msoffice" (
  call :verify_msoffice ".xlsx" "Excel.Sheet.12"
  call :verify_msoffice ".pptx" "PowerPoint.Show.12"
  call :verify_msoffice ".docx" "Word.Document.12"
)
if "%FIX_MODE%"=="wps" (
  call :verify_wps ".xlsx"
  call :verify_wps ".pptx"
  call :verify_wps ".docx"
)

echo   Verified: %VERIFY_OK% / 3
echo.

if "%MODIFY_COUNT%"=="0" goto no_changes

echo ============================================================
echo   Fix complete! Fixed %MODIFY_COUNT% right-click New items
echo   Mode: %FIX_MODE%
echo ============================================================
echo.
if "%SILENT%"=="0" msg * Right-click New template fix complete!

if "%FIX_MODE%"=="wps" (
  echo   WPS version: %WPS_VERSION%
  echo   Template dir: %FIXED_TEMPLATE_DIR%
)
if "%FIX_MODE%"=="msoffice" echo   MS Office mode: NullFile + display names
echo.

rem Restart Explorer after fix
echo   Restarting Explorer to apply changes...
taskkill /f /im explorer.exe >nul 2>&1
timeout /t 1 /nobreak >nul 2>&1
start explorer.exe
echo   Explorer restarted.
echo.

goto end_script

:no_changes
echo   No registry entries modified.
echo.

:end_script
echo.
if "%SILENT%"=="0" pause
endlocal
ver >nul
exit /b 0

rem ============================================================
rem Subroutines
rem ============================================================

:clean_ext
set "_CEXT=%~1"
set "_CPROGID=%~2"
reg query "HKCR\%_CEXT%\ShellNew" >nul 2>&1
if not errorlevel 1 (
  echo   - %_CEXT% remove ShellNew (extension level)
  reg delete "HKCR\%_CEXT%\ShellNew" /f >nul 2>&1
  set /a CLEAN_COUNT+=1
)
reg query "HKCR\%_CEXT%\%_CPROGID%\ShellNew" >nul 2>&1
if not errorlevel 1 (
  echo   - %_CEXT% remove ShellNew (ProgID level)
  reg delete "HKCR\%_CEXT%\%_CPROGID%\ShellNew" /f >nul 2>&1
  set /a CLEAN_COUNT+=1
)
ver >nul
exit /b 0

:recheck_ext
set "_REXT=%~1"
set "_RPROGID=%~2"
set "_FOUND=0"
reg query "HKCR\%_REXT%\ShellNew" >nul 2>&1
if not errorlevel 1 (
  echo   - %_REXT% still has ShellNew (extension level), cleaning
  reg delete "HKCR\%_REXT%\ShellNew" /f >nul 2>&1
  set _FOUND=1
)
reg query "HKCR\%_REXT%\%_RPROGID%\ShellNew" >nul 2>&1
if not errorlevel 1 (
  echo   - %_REXT% still has ShellNew (ProgID level), cleaning
  reg delete "HKCR\%_REXT%\%_RPROGID%\ShellNew" /f >nul 2>&1
  set _FOUND=1
)
if "%_FOUND%"=="1" set /a RECHECK_FOUND+=1
ver >nul
exit /b 0

:fix_msoffice
set "_EXT=%~1"
set "_PROGID=%~2"
reg add "HKCR\%_EXT%" /ve /d "%_PROGID%" /f >nul 2>&1
reg add "HKCR\%_EXT%\%_PROGID%" /f >nul 2>&1
reg add "HKCR\%_EXT%\%_PROGID%\ShellNew" /f >nul 2>&1
reg add "HKCR\%_EXT%\%_PROGID%\ShellNew" /v "NullFile" /t REG_SZ /d "" /f >nul 2>&1
reg delete "HKCR\%_EXT%\ShellNew" /f >nul 2>&1
ver >nul
echo   [%_EXT%] fix success (NullFile)
exit /b 1

:fix_wps
set "_EXT=%~1"
set "_PROGID=%~2"
set "_TEMPLATE=%~3"
if "%_TEMPLATE%"=="" echo   [%_EXT%] no template, skip & exit /b 0
reg add "HKCR\%_EXT%" /ve /d "%_PROGID%" /f >nul 2>&1
reg add "HKCR\%_EXT%\ShellNew" /f >nul 2>&1
reg delete "HKCR\%_EXT%\ShellNew" /v "NullFile" /f >nul 2>&1
reg add "HKCR\%_EXT%\ShellNew" /v "FileName" /t REG_SZ /d "%_TEMPLATE%" /f >nul 2>&1
reg add "HKCR\%_EXT%\%_PROGID%" /f >nul 2>&1
reg add "HKCR\%_EXT%\%_PROGID%\ShellNew" /f >nul 2>&1
reg delete "HKCR\%_EXT%\%_PROGID%\ShellNew" /v "NullFile" /f >nul 2>&1
reg add "HKCR\%_EXT%\%_PROGID%\ShellNew" /v "FileName" /t REG_SZ /d "%_TEMPLATE%" /f >nul 2>&1
ver >nul
echo   [%_EXT%] fix success (FileName)
exit /b 1

:verify_msoffice
set "_VEXT=%~1"
set "_VPROGID=%~2"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ext='%_VEXT%';$progId='%_VPROGID%'; $sn=Get-ItemProperty ('Registry::HKEY_CLASSES_ROOT\'+$ext+'\'+$progId+'\ShellNew') -EA SilentlyContinue; $dn=(Get-ItemProperty ('Registry::HKEY_CLASSES_ROOT\'+$progId) -EA SilentlyContinue).'(default)'; if($sn -and $sn.NullFile -ne $null -and $dn){Write-Host('  - ['+$ext+'] OK'); exit 0}else{if(-not $sn -or $sn.NullFile -eq $null){Write-Host('  - ['+$ext+'] ShellNew NullFile missing')}; if(-not $dn){Write-Host('  - ['+$ext+'] display name missing')}; exit 1}"
if errorlevel 1 exit /b 0
set /a VERIFY_OK+=1
exit /b 0

:verify_wps
set "_VEXT=%~1"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ext='%_VEXT%'; $fn=(Get-ItemProperty ('Registry::HKEY_CLASSES_ROOT\'+$ext+'\ShellNew') -EA SilentlyContinue).FileName; if($fn){Write-Host('  - ['+$ext+'] OK'); exit 0}else{Write-Host('  - ['+$ext+'] ShellNew FileName missing'); exit 1}"
if errorlevel 1 exit /b 0
set /a VERIFY_OK+=1
exit /b 0

:cache_clean
powershell -NoProfile -ExecutionPolicy Bypass -Command "&{ $p='HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Discardable\PostSetup\ShellNew'; $c=(Get-ItemProperty -Path $p -Name 'Classes' -EA SilentlyContinue).Classes; if(-not $c){$c=@()}; $l=[System.Collections.ArrayList]@(); foreach($i in $c){ if($i -notmatch '^\.(xlsx|pptx|docx|xls|ppt|doc|et|wps|dps)$'){[void]$l.Add($i)} }; $rk=[Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('Software\Microsoft\Windows\CurrentVersion\Explorer\Discardable\PostSetup\ShellNew',$true); $rk.SetValue('Classes',$l.ToArray([string]),[Microsoft.Win32.RegistryValueKind]::MultiString); $rk.Close(); Write-Host '  done' }"
ver >nul
exit /b 0

:cache_update
powershell -NoProfile -ExecutionPolicy Bypass -Command "&{ $p='HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Discardable\PostSetup\ShellNew'; $c=(Get-ItemProperty -Path $p -Name 'Classes' -EA SilentlyContinue).Classes; if(-not $c){$c=@()}; $l=[System.Collections.ArrayList]@(); foreach($i in $c){ if($i -notmatch '^\.(ppt|doc|xls|et|wps|dps)$'){[void]$l.Add($i)} }; foreach($e in @('.xlsx','.pptx','.docx')){ if($l -notcontains $e){[void]$l.Add($e); Write-Host('  add '+$e)} }; $rk=[Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('Software\Microsoft\Windows\CurrentVersion\Explorer\Discardable\PostSetup\ShellNew',$true); $rk.SetValue('Classes',$l.ToArray([string]),[Microsoft.Win32.RegistryValueKind]::MultiString); $rk.Close(); Write-Host '  done' }"
if errorlevel 1 (
  echo   [WARNING] cache update failed
) else (
  echo   cache update success
)
exit /b 0

:send_notify
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -Namespace W -Name S -MemberDefinition '[System.Runtime.InteropServices.DllImport(\"shell32.dll\")] public static extern void SHChangeNotify(int w, int f, IntPtr d1, IntPtr d2);'; [W.S]::SHChangeNotify(0x08000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)"
ver >nul
exit /b 0

:detect_wps
for /f "skip=2 tokens=2,*" %%a in ('reg query "HKLM\SOFTWARE\Kingsoft\Office" /v "InstallRoot" 2^>nul') do set "WPS_INSTALL_PATH=%%b"
if not defined WPS_INSTALL_PATH if exist "C:\Program Files\Kingsoft\WPS Office" set "WPS_INSTALL_PATH=C:\Program Files\Kingsoft\WPS Office"
if not defined WPS_INSTALL_PATH if exist "C:\Program Files (x86)\Kingsoft\WPS Office" set "WPS_INSTALL_PATH=C:\Program Files (x86)\Kingsoft\WPS Office"
if defined WPS_INSTALL_PATH if exist "%WPS_INSTALL_PATH%" (
  for /f "delims=" %%d in ('dir /b /ad "%WPS_INSTALL_PATH%" 2^>nul ^| findstr /r "^[0-9][0-9]*\.[0-9]"') do (
    if exist "%WPS_INSTALL_PATH%\%%d\office6\et.exe" (
      set "WPS_VERSION=%%d"
      set HAS_WPS=1
    )
  )
)
if "%HAS_WPS%"=="1" set "WPS_TEMPLATE_DIR=%WPS_INSTALL_PATH%\%WPS_VERSION%\office6\mui\zh_CN\templates"
exit /b 0

:prepare_wps_templates
set "FIXED_TEMPLATE_DIR=C:\ProgramData\WPS Templates"
echo   Fixed template dir: %FIXED_TEMPLATE_DIR%
if not exist "%FIXED_TEMPLATE_DIR%" mkdir "%FIXED_TEMPLATE_DIR%" >nul 2>&1
set "WPS_VERSION_FILE=%FIXED_TEMPLATE_DIR%\.wps_version"
set WPS_NEED_UPDATE=0
if not exist "%WPS_VERSION_FILE%" set WPS_NEED_UPDATE=1
if "%WPS_NEED_UPDATE%"=="1" echo   First run, copying templates...
if "%WPS_NEED_UPDATE%"=="1" goto do_copy
for /f "usebackq delims=" %%v in ("%WPS_VERSION_FILE%") do set SAVED_WPS_VERSION=%%v
if not "%SAVED_WPS_VERSION%"=="%WPS_VERSION%" (
  set WPS_NEED_UPDATE=1
  echo   WPS updated, re-copying templates...
  goto do_copy
)
if not exist "%FIXED_TEMPLATE_DIR%\newfile.xlsx" set WPS_NEED_UPDATE=1
if not exist "%FIXED_TEMPLATE_DIR%\newfile.pptx" set WPS_NEED_UPDATE=1
if not exist "%FIXED_TEMPLATE_DIR%\Normal.dotm" set WPS_NEED_UPDATE=1
if "%WPS_NEED_UPDATE%"=="1" echo   Template files incomplete, re-copying...
if "%WPS_NEED_UPDATE%"=="1" goto do_copy
echo   WPS version unchanged, skip copy
goto skip_copy
:do_copy
if exist "%WPS_TEMPLATE_DIR%\newfile.xlsx" (
  copy /y "%WPS_TEMPLATE_DIR%\newfile.xlsx" "%FIXED_TEMPLATE_DIR%\newfile.xlsx" >nul 2>&1
  echo   [xlsx] copied
)
if exist "%WPS_TEMPLATE_DIR%\newfile.pptx" (
  copy /y "%WPS_TEMPLATE_DIR%\newfile.pptx" "%FIXED_TEMPLATE_DIR%\newfile.pptx" >nul 2>&1
  echo   [pptx] copied
)
if exist "%WPS_TEMPLATE_DIR%\Normal.dotm" (
  copy /y "%WPS_TEMPLATE_DIR%\Normal.dotm" "%FIXED_TEMPLATE_DIR%\Normal.dotm" >nul 2>&1
  echo   [docx] copied
)
> "%WPS_VERSION_FILE%" echo %WPS_VERSION%
echo   Version recorded: %WPS_VERSION%
:skip_copy
if exist "%FIXED_TEMPLATE_DIR%\newfile.xlsx" set "XLSX_TEMPLATE=%FIXED_TEMPLATE_DIR%\newfile.xlsx"
if exist "%FIXED_TEMPLATE_DIR%\newfile.pptx" set "PPTX_TEMPLATE=%FIXED_TEMPLATE_DIR%\newfile.pptx"
if exist "%FIXED_TEMPLATE_DIR%\Normal.dotm" set "DOCX_TEMPLATE=%FIXED_TEMPLATE_DIR%\Normal.dotm"
exit /b 0
