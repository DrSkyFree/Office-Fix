@echo off
chcp 936 >nul 2>&1
setlocal

rem ============================================================
rem   Right-Click New Menu Template Fix Tool v3.0 (streamlined)
rem   Step 1: Clean old ShellNew entries (MS Office + WPS) + user verify
rem   Step 2: Detect Office installations
rem   Step 3: Fix - MS Office (NullFile) / WPS (FileName + template)
rem   Step 4: Verify + refresh ShellNew cache + restart Explorer
rem   Args: /silent or /quiet : no prompts
rem         /mode 1|2         : preselect fix mode when both installed
rem         /restart         : accepted for compatibility (now default)
rem ============================================================

set SILENT=0
set USER_CHOICE=

:parse_args
if "%~1"=="" goto args_done
if /i "%~1"=="/silent" set SILENT=1
if /i "%~1"=="/quiet" set SILENT=1
if /i "%~1"=="/mode" (set "USER_CHOICE=%~2" & shift)
shift
goto parse_args
:args_done

net session >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Please run this script as Administrator.
    if "%SILENT%"=="0" pause
    exit /b 1
)

echo ============================================================
echo   Right-Click New Menu Template Fix Tool v3.0
echo ============================================================
echo.

echo [Step 1/Clean] Cleaning all right-click New menu entries...
call :clean_all
echo   Total cleaned: %CLEAN_COUNT% items.
call :cache_sync "^\.(xlsx|pptx|docx|xls|ppt|doc|et|wps|dps)$" "@()"
call :restart_explorer

if "%SILENT%"=="1" goto detect

:verify_loop
echo Please right-click on the desktop or in a folder and check the New menu:
echo   .xlsx / .pptx / .docx / .xls / .ppt / .doc should all be gone.
choice /c YN /n /m "  All removed? [Y=proceed, N=re-clean]: "
if errorlevel 2 goto recheck
goto detect

:recheck
call :clean_all
if "%CLEAN_COUNT%"=="0" (
    echo   No remaining entries in registry.
    goto detect
)
echo   Cleaned %CLEAN_COUNT% more entries, refreshing again...
call :cache_sync "^\.(xlsx|pptx|docx|xls|ppt|doc|et|wps|dps)$" "@()"
call :restart_explorer
goto verify_loop

:detect
echo.
echo [Step 2/Detect] Detecting Office installation...
set HAS_MSOFFICE=0
set HAS_WPS=0
set "MSOFFICE_VERSION="
set "WPS_VERSION="
set "WPS_TEMPLATE_DIR="

set "MSO_WORD="
for /f "skip=2 tokens=2,*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE" /ve 2^>nul') do set "MSO_WORD=%%b"
if not defined MSO_WORD for /f "skip=2 tokens=2,*" %%a in ('reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE" /ve 2^>nul') do set "MSO_WORD=%%b"
if defined MSO_WORD if exist "%MSO_WORD%" call :detect_msoffice
call :detect_wps
if "%HAS_WPS%"=="1" echo   - WPS Office detected (version %WPS_VERSION%)

if "%HAS_MSOFFICE%"=="0" if "%HAS_WPS%"=="0" (
    echo   [WARNING] No Office software detected, exiting.
    if "%SILENT%"=="0" pause
    exit /b 0
)

set FIX_MODE=
if "%HAS_MSOFFICE%"=="1" if "%HAS_WPS%"=="0" set FIX_MODE=msoffice
if "%HAS_WPS%"=="1" if "%HAS_MSOFFICE%"=="0" set FIX_MODE=wps
if "%HAS_MSOFFICE%"=="1" if "%HAS_WPS%"=="1" (
    echo   Both MS Office %MSOFFICE_VERSION% and WPS %WPS_VERSION% detected.
    if "%SILENT%"=="1" (
        set FIX_MODE=wps
        if "%USER_CHOICE%"=="1" set FIX_MODE=msoffice
    ) else (
        call :choose_mode
        if errorlevel 3 goto cancelled
    )
)

echo.
echo [Step 3/Fix] Fixing right-click New menu (mode: %FIX_MODE%)...
set MODIFY_COUNT=0
if not "%FIX_MODE%"=="msoffice" goto fix_wps_mode

call :fix_msoffice ".xlsx" "Excel.Sheet.12" && set /a MODIFY_COUNT+=1
call :fix_msoffice ".pptx" "PowerPoint.Show.12" && set /a MODIFY_COUNT+=1
call :fix_msoffice ".docx" "Word.Document.12" && set /a MODIFY_COUNT+=1
goto display_names

:fix_wps_mode
call :prepare_wps_templates
call :fix_wps ".xlsx" "Excel.Sheet.12" "%XLSX_TEMPLATE%" && set /a MODIFY_COUNT+=1
call :fix_wps ".pptx" "PowerPoint.Show.12" "%PPTX_TEMPLATE%" && set /a MODIFY_COUNT+=1
call :fix_wps ".docx" "Word.Document.12" "%DOCX_TEMPLATE%" && set /a MODIFY_COUNT+=1

:display_names
echo   Setting display names...
reg add "HKCR\Excel.Sheet.12" /ve /d "Microsoft Excel" /f >nul 2>&1
reg add "HKCR\PowerPoint.Show.12" /ve /d "Microsoft PowerPoint" /f >nul 2>&1
reg add "HKCR\Word.Document.12" /ve /d "Microsoft Word" /f >nul 2>&1
reg delete "HKCR\Excel.Sheet.12" /v "FriendlyTypeName" /f >nul 2>&1
reg delete "HKCR\PowerPoint.Show.12" /v "FriendlyTypeName" /f >nul 2>&1
reg delete "HKCR\Word.Document.12" /v "FriendlyTypeName" /f >nul 2>&1

call :cache_sync "^\.(ppt|doc|xls|et|wps|dps)$" "@('.xlsx','.pptx','.docx')"
call :send_notify

echo.
echo [Step 4/Verify] Verifying fix results...
set VERIFY_OK=0
if "%FIX_MODE%"=="msoffice" (
    call :verify_ms ".xlsx" "Excel.Sheet.12"
    call :verify_ms ".pptx" "PowerPoint.Show.12"
    call :verify_ms ".docx" "Word.Document.12"
) else (
    call :verify_wps ".xlsx"
    call :verify_wps ".pptx"
    call :verify_wps ".docx"
)
echo   Verified: %VERIFY_OK% / 3
echo.

if "%MODIFY_COUNT%"=="0" (
    echo   No registry entries modified.
    goto end
)

echo ============================================================
echo   Fix complete! Fixed %MODIFY_COUNT% New items (mode: %FIX_MODE%)
echo ============================================================
if "%FIX_MODE%"=="wps" echo   Template dir: C:\ProgramData\WPS Templates
if "%SILENT%"=="0" msg * Right-click New template fix complete!
call :restart_explorer
goto end

:cancelled
echo   Cancelled by user.
if "%SILENT%"=="0" pause
exit /b 0

:end
echo.
if "%SILENT%"=="0" pause
exit /b 0

rem ============================================================
rem  Subroutines
rem ============================================================

:clean_all
rem Clean all extensions (ext level + ProgID level) and WPS root keys
set CLEAN_COUNT=0
call :clean_ext .xlsx Excel.Sheet.12
call :clean_ext .pptx PowerPoint.Show.12
call :clean_ext .docx Word.Document.12
call :clean_ext .xls Excel.Sheet.8
call :clean_ext .ppt PowerPoint.Show.8
call :clean_ext .doc Word.Document.8
for %%e in (et wps dps .et .wps .dps) do (
    reg delete "HKCR\%%e" /f >nul 2>&1
    if not errorlevel 1 (echo   - HKCR\%%e removed & set /a CLEAN_COUNT+=1)
)
exit /b 0

:clean_ext
reg delete "HKCR\%~1\ShellNew" /f >nul 2>&1
if not errorlevel 1 (echo   - %~1 ShellNew removed [ext level] & set /a CLEAN_COUNT+=1)
reg delete "HKCR\%~1\%~2\ShellNew" /f >nul 2>&1
if not errorlevel 1 (echo   - %~1 ShellNew removed [ProgID level] & set /a CLEAN_COUNT+=1)
exit /b 0

:detect_msoffice
set HAS_MSOFFICE=1
for /f "skip=2 tokens=2,*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v "VersionToReport" 2^>nul') do set "MSOFFICE_VERSION=%%b"
if not defined MSOFFICE_VERSION set "MSOFFICE_VERSION=16.0"
echo   - MS Office detected (version %MSOFFICE_VERSION%)
exit /b 0

:detect_wps
set "WPS_INSTALL_PATH="
for /f "skip=2 tokens=2,*" %%a in ('reg query "HKLM\SOFTWARE\Kingsoft\Office" /v "InstallRoot" 2^>nul') do set "WPS_INSTALL_PATH=%%b"
if not defined WPS_INSTALL_PATH for %%d in ("C:\Program Files\Kingsoft\WPS Office" "C:\Program Files (x86)\Kingsoft\WPS Office") do if not defined WPS_INSTALL_PATH if exist "%%~d" set "WPS_INSTALL_PATH=%%~d"
if not defined WPS_INSTALL_PATH exit /b 0
for /f "delims=" %%d in ('dir /b /ad "%WPS_INSTALL_PATH%" 2^>nul ^| findstr /r "^[0-9][0-9]*\.[0-9]"') do (
    if exist "%WPS_INSTALL_PATH%\%%d\office6\et.exe" (set "WPS_VERSION=%%d" & set HAS_WPS=1)
)
if "%HAS_WPS%"=="1" set "WPS_TEMPLATE_DIR=%WPS_INSTALL_PATH%\%WPS_VERSION%\office6\mui\zh_CN\templates"
exit /b 0

:choose_mode
echo   Choose fix mode:
echo     [1] Microsoft Office (NullFile)
echo     [2] WPS (FileName template)
echo     [C] Cancel
choice /c 12C /n /m "  Enter choice [1/2/C]: "
if errorlevel 3 exit /b 3
if errorlevel 2 (set "FIX_MODE=wps" & echo   WPS mode selected.) else (set "FIX_MODE=msoffice" & echo   MS Office mode selected.)
exit /b 0

:fix_msoffice
reg add "HKCR\%~1" /ve /d "%~2" /f >nul 2>&1
reg delete "HKCR\%~1\ShellNew" /f >nul 2>&1
reg add "HKCR\%~1\%~2\ShellNew" /v "NullFile" /t REG_SZ /d "" /f >nul 2>&1
echo   [%~1] fixed (NullFile)
exit /b 1

:fix_wps
if "%~3"=="" (echo   [%~1] no template, skip & exit /b 0)
reg add "HKCR\%~1" /ve /d "%~2" /f >nul 2>&1
for %%k in ("%~1\ShellNew" "%~1\%~2\ShellNew") do (
    reg add "HKCR\%%~k" /v "FileName" /t REG_SZ /d "%~3" /f >nul 2>&1
    reg delete "HKCR\%%~k" /v "NullFile" /f >nul 2>&1
)
echo   [%~1] fixed (FileName)
exit /b 1

:prepare_wps_templates
set "FIXED_TEMPLATE_DIR=C:\ProgramData\WPS Templates"
if not exist "%FIXED_TEMPLATE_DIR%" mkdir "%FIXED_TEMPLATE_DIR%" >nul 2>&1
set "WPS_VERSION_FILE=%FIXED_TEMPLATE_DIR%\.wps_version"
set "SAVED_WPS_VERSION="
if exist "%WPS_VERSION_FILE%" for /f "usebackq delims=" %%v in ("%WPS_VERSION_FILE%") do set "SAVED_WPS_VERSION=%%v"
if "%SAVED_WPS_VERSION%"=="%WPS_VERSION%" if exist "%FIXED_TEMPLATE_DIR%\newfile.xlsx" if exist "%FIXED_TEMPLATE_DIR%\newfile.pptx" if exist "%FIXED_TEMPLATE_DIR%\Normal.dotm" goto templates_ok
echo   Syncing WPS templates to %FIXED_TEMPLATE_DIR% ...
for %%f in (newfile.xlsx newfile.pptx Normal.dotm) do if exist "%WPS_TEMPLATE_DIR%\%%f" copy /y "%WPS_TEMPLATE_DIR%\%%f" "%FIXED_TEMPLATE_DIR%\" >nul 2>&1
> "%WPS_VERSION_FILE%" echo %WPS_VERSION%
:templates_ok
if exist "%FIXED_TEMPLATE_DIR%\newfile.xlsx" set "XLSX_TEMPLATE=%FIXED_TEMPLATE_DIR%\newfile.xlsx"
if exist "%FIXED_TEMPLATE_DIR%\newfile.pptx" set "PPTX_TEMPLATE=%FIXED_TEMPLATE_DIR%\newfile.pptx"
if exist "%FIXED_TEMPLATE_DIR%\Normal.dotm" set "DOCX_TEMPLATE=%FIXED_TEMPLATE_DIR%\Normal.dotm"
exit /b 0

:verify_ms
powershell -NoProfile -ExecutionPolicy Bypass -Command "$e='%~1';$p='%~2';if((Get-ItemProperty ('Registry::HKEY_CLASSES_ROOT\'+$e+'\'+$p+'\ShellNew') -EA 0).NullFile -ne $null -and (Get-ItemProperty ('Registry::HKEY_CLASSES_ROOT\'+$p) -EA 0).'(default)'){Write-Host ('  - ['+$e+'] OK');exit 0};Write-Host ('  - ['+$e+'] FAIL');exit 1"
if not errorlevel 1 set /a VERIFY_OK+=1
exit /b 0

:verify_wps
powershell -NoProfile -ExecutionPolicy Bypass -Command "$e='%~1';$f=(Get-ItemProperty ('Registry::HKEY_CLASSES_ROOT\'+$e+'\ShellNew') -EA 0).FileName;if($f -and (Test-Path $f)){Write-Host ('  - ['+$e+'] OK');exit 0};Write-Host ('  - ['+$e+'] FAIL');exit 1"
if not errorlevel 1 set /a VERIFY_OK+=1
exit /b 0

:cache_sync
rem %1 = regex of extensions to remove from cache, %2 = PS array of extensions to ensure
powershell -NoProfile -ExecutionPolicy Bypass -Command "$p='HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Discardable\PostSetup\ShellNew';$c=@((Get-ItemProperty $p -EA 0).Classes);if(-not $c){$c=@()};$l=New-Object System.Collections.ArrayList;foreach($i in $c){if($i -and $i -notmatch '%~1'){[void]$l.Add($i)}};foreach($e in %~2){if(-not $l.Contains($e)){[void]$l.Add($e)}};$k=[Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($p);$k.SetValue('Classes',[string[]]$l,[Microsoft.Win32.RegistryValueKind]::MultiString);$k.Close()"
exit /b 0

:send_notify
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -Namespace W -Name S -MemberDefinition '[System.Runtime.InteropServices.DllImport(\"shell32.dll\")] public static extern void SHChangeNotify(int w, int f, IntPtr d1, IntPtr d2);'; [W.S]::SHChangeNotify(0x08000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)"
exit /b 0

:restart_explorer
echo   Restarting Explorer to apply changes...
taskkill /f /im explorer.exe >nul 2>&1
timeout /t 1 /nobreak >nul 2>&1
start explorer.exe
echo   Explorer restarted.
exit /b 0
