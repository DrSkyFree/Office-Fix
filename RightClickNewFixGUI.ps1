# ============================================================
#  Right-Click New Menu Template Fix Tool (GUI Edition) v3.0
#
#  Ported from: 右键新建模板修复.bat v2.9
#  Step 1: Clean old ShellNew entries (MS Office + WPS) + user verify loop
#  Step 2: Detect Office installations and show versions
#  Step 3: Fix (MS Office NullFile / WPS FileName with template)
#  Step 4: Verify registry writes + restart Explorer
#
#  Requires : Windows PowerShell 5.1 (built into Win10/Win11)
#  Usage    : double-click "右键新建模板修复GUI.bat", or run:
#             powershell -NoProfile -ExecutionPolicy Bypass -STA -File RightClickNewFixGUI.ps1
#             -DetectOnly : read-only detection, no GUI, no registry changes
# ============================================================

param(
    [switch]$DetectOnly
)

$script:txtLog = $null

# ------------------------------------------------------------
# Extension / ProgID mapping (same as the original .bat)
# ------------------------------------------------------------
$script:OfficeExtMap = @(
    @{ Ext = '.xlsx'; ProgID = 'Excel.Sheet.12';      DisplayName = 'Microsoft Excel 工作簿 (.xlsx)' },
    @{ Ext = '.pptx'; ProgID = 'PowerPoint.Show.12';   DisplayName = 'Microsoft PowerPoint 演示文稿 (.pptx)' },
    @{ Ext = '.docx'; ProgID = 'Word.Document.12';    DisplayName = 'Microsoft Word 文档 (.docx)' },
    @{ Ext = '.xls';  ProgID = 'Excel.Sheet.8';        DisplayName = 'Microsoft Excel 97-2003 工作簿 (.xls)' },
    @{ Ext = '.ppt';  ProgID = 'PowerPoint.Show.8';   DisplayName = 'Microsoft PowerPoint 97-2003 演示文稿 (.ppt)' },
    @{ Ext = '.doc';  ProgID = 'Word.Document.8';     DisplayName = 'Microsoft Word 97-2003 文档 (.doc)' }
)

# Modern formats used by the fix step
$script:ModernExtMap = @(
    @{ Ext = '.xlsx'; ProgID = 'Excel.Sheet.12';    ProgName = 'Microsoft Excel';       TemplateFile = 'newfile.xlsx' },
    @{ Ext = '.pptx'; ProgID = 'PowerPoint.Show.12'; ProgName = 'Microsoft PowerPoint'; TemplateFile = 'newfile.pptx' },
    @{ Ext = '.docx'; ProgID = 'Word.Document.12';  ProgName = 'Microsoft Word';       TemplateFile = 'Normal.dotm' }
)

# WPS-specific root keys to purge during cleanup
$script:WpsRootKeys = @('et', 'wps', 'dps', '.et', '.wps', '.dps')

$script:FixedTemplateDir = 'C:\ProgramData\WPS Templates'
$script:WpsVersionFile = Join-Path $script:FixedTemplateDir '.wps_version'

# ------------------------------------------------------------
# Logging (GUI textbox if available, otherwise console)
# ------------------------------------------------------------
function Log {
    param([string]$Message)
    $line = ('[{0}] {1}' -f (Get-Date -Format 'HH:mm:ss'), $Message)
    if ($null -ne $script:txtLog) {
        $script:txtLog.AppendText("$line`r`n")
        [System.Windows.Forms.Application]::DoEvents()
    } else {
        Write-Host $line
    }
}

# ------------------------------------------------------------
# Registry helpers (HKCR via .NET, avoids parsing reg.exe output)
# ------------------------------------------------------------
function Test-RegKeyExists {
    param([string]$Path)
    $k = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey($Path)
    if ($null -ne $k) { $k.Close(); return $true }
    return $false
}

function Get-RegString {
    param([string]$Path, [string]$Name)
    $k = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey($Path)
    if ($null -eq $k) { return $null }
    $v = $k.GetValue($Name)
    $k.Close()
    return $v
}

function Set-RegString {
    param([string]$Path, [string]$Name, [string]$Value)
    $k = [Microsoft.Win32.Registry]::ClassesRoot.CreateSubKey($Path)
    $k.SetValue($Name, $Value, [Microsoft.Win32.RegistryValueKind]::String)
    $k.Close()
}

function Remove-RegKeyIfExists {
    param([string]$Path)
    if (-not (Test-RegKeyExists $Path)) { return $false }
    try {
        [Microsoft.Win32.Registry]::ClassesRoot.DeleteSubKeyTree($Path)
        return $true
    } catch {
        return $false
    }
}

function Remove-RegValueIfExists {
    param([string]$Path, [string]$Name)
    $k = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey($Path, $true)
    if ($null -eq $k) { return $false }
    if ($null -ne $k.GetValue($Name)) {
        $k.DeleteValue($Name)
        $k.Close()
        return $true
    }
    $k.Close()
    return $false
}

# ------------------------------------------------------------
# Explorer ShellNew cache (HKCU ... Discardable PostSetup ShellNew)
# ------------------------------------------------------------
function Update-ShellNewCache {
    param([string[]]$RemoveExtensions, [string[]]$EnsureExtensions)
    $cachePath = 'Software\Microsoft\Windows\CurrentVersion\Explorer\Discardable\PostSetup\ShellNew'
    $key = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($cachePath)
    if ($null -eq $key) { return $false }
    $classes = @($key.GetValue('Classes'))
    $list = New-Object System.Collections.ArrayList
    foreach ($item in $classes) {
        if (-not $item) { continue }
        if ($RemoveExtensions -notcontains $item.ToString().ToLower()) {
            [void]$list.Add($item.ToString())
        }
    }
    foreach ($e in $EnsureExtensions) {
        if (-not $list.Contains($e)) { [void]$list.Add($e) }
    }
    $key.SetValue('Classes', [string[]]@($list), [Microsoft.Win32.RegistryValueKind]::MultiString)
    $key.Close()
    return $true
}

function Clear-OfficeShellNewCache {
    # Remove ALL office-related extensions from the cache
    $all = @('.xlsx', '.pptx', '.docx', '.xls', '.ppt', '.doc', '.et', '.wps', '.dps')
    return (Update-ShellNewCache -RemoveExtensions $all -EnsureExtensions @())
}

function Refresh-OfficeShellNewCache {
    # Remove 97-2003 + WPS formats, ensure modern formats present
    $legacy = @('.ppt', '.doc', '.xls', '.et', '.wps', '.dps')
    $modern = @('.xlsx', '.pptx', '.docx')
    return (Update-ShellNewCache -RemoveExtensions $legacy -EnsureExtensions $modern)
}

# ------------------------------------------------------------
# Shell notification (SHChangeNotify)
# ------------------------------------------------------------
function Send-ShellNotify {
    try {
        if (-not ('W.ShellNotify' -as [type])) {
            Add-Type -Namespace 'W' -Name 'ShellNotify' -MemberDefinition `
                '[System.Runtime.InteropServices.DllImport("shell32.dll")] public static extern void SHChangeNotify(int wEventId, int uFlags, System.IntPtr dwItem1, System.IntPtr dwItem2);'
        }
        [W.ShellNotify]::SHChangeNotify(0x08000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)
        return $true
    } catch {
        return $false
    }
}

# ------------------------------------------------------------
# Explorer restart
# ------------------------------------------------------------
function Restart-Explorer {
    Log '正在重启资源管理器以应用更改...'
    try { Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue } catch { }
    Start-Sleep -Seconds 1
    if (-not (Get-Process -Name explorer -ErrorAction SilentlyContinue)) {
        Start-Process 'explorer.exe'
    }
    Log '资源管理器已重启。'
}

# ------------------------------------------------------------
# Step 1: Cleanup
# ------------------------------------------------------------
function Clean-ShellNewEntries {
    $count = 0
    foreach ($m in $script:OfficeExtMap) {
        if (Remove-RegKeyIfExists "$($m.Ext)\ShellNew") {
            Log ("  - 已清理 {0}（扩展名级 ShellNew）" -f $m.Ext)
            $count++
        }
        if (Remove-RegKeyIfExists "$($m.Ext)\$($m.ProgID)\ShellNew") {
            Log ("  - 已清理 {0}（ProgID 级 ShellNew）" -f $m.Ext)
            $count++
        }
    }
    foreach ($k in $script:WpsRootKeys) {
        if (Remove-RegKeyIfExists $k) {
            Log ("  - 已清理 WPS 相关键 HKCR\{0}" -f $k)
            $count++
        }
    }
    return $count
}

function Recheck-AndCleanShellNew {
    # Re-check registry for remaining ShellNew entries; clean if found.
    # Returns the number of extensions whose entries were (re)cleaned.
    $found = 0
    foreach ($m in $script:OfficeExtMap) {
        $had = $false
        if (Test-RegKeyExists "$($m.Ext)\ShellNew") {
            Log ("  - 复查发现 {0} 仍存在扩展名级 ShellNew，已清理" -f $m.Ext)
            [void](Remove-RegKeyIfExists "$($m.Ext)\ShellNew")
            $had = $true
        }
        if (Test-RegKeyExists "$($m.Ext)\$($m.ProgID)\ShellNew") {
            Log ("  - 复查发现 {0} 仍存在 ProgID 级 ShellNew，已清理" -f $m.Ext)
            [void](Remove-RegKeyIfExists "$($m.Ext)\$($m.ProgID)\ShellNew")
            $had = $true
        }
        if ($had) { $found++ }
    }
    return $found
}

# ------------------------------------------------------------
# Step 2: Detection
# ------------------------------------------------------------
function Detect-MsOffice {
    $appPathKeys = @(
        'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE',
        'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE'
    )
    $wordPath = $null
    foreach ($p in $appPathKeys) {
        $k = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($p)
        if ($null -ne $k) {
            $v = [string]$k.GetValue('')
            $k.Close()
            if ($v -and (Test-Path $v)) { $wordPath = $v; break }
        }
    }
    if (-not $wordPath) {
        return @{ Found = $false; Version = '' }
    }
    $version = $null
    $c2r = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey('SOFTWARE\Microsoft\Office\ClickToRun\Configuration')
    if ($null -ne $c2r) {
        $version = [string]$c2r.GetValue('VersionToReport')
        $c2r.Close()
    }
    if (-not $version) {
        $root = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey('SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot')
        if ($null -ne $root) { $version = '16.0 (MSI)'; $root.Close() }
    }
    if (-not $version) { $version = '未知版本' }
    return @{ Found = $true; Version = $version }
}

function Detect-Wps {
    $installPath = $null
    $k = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey('SOFTWARE\Kingsoft\Office')
    if ($null -ne $k) {
        $installPath = [string]$k.GetValue('InstallRoot')
        $k.Close()
    }
    if (-not $installPath -or -not (Test-Path $installPath)) {
        foreach ($d in @('C:\Program Files\Kingsoft\WPS Office', 'C:\Program Files (x86)\Kingsoft\WPS Office')) {
            if (Test-Path $d) { $installPath = $d; break }
        }
    }
    if (-not $installPath -or -not (Test-Path $installPath)) {
        return @{ Found = $false; Version = ''; InstallPath = ''; TemplateDir = '' }
    }
    $version = $null
    $dirs = Get-ChildItem -Path $installPath -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^[0-9]+\.[0-9]' } |
        Sort-Object Name -Descending
    foreach ($d in $dirs) {
        if (Test-Path (Join-Path $d.FullName 'office6\et.exe')) { $version = $d.Name; break }
    }
    if (-not $version) {
        return @{ Found = $false; Version = ''; InstallPath = $installPath; TemplateDir = '' }
    }
    $templateDir = Join-Path $installPath ("{0}\office6\mui\zh_CN\templates" -f $version)
    return @{ Found = $true; Version = $version; InstallPath = $installPath; TemplateDir = $templateDir }
}

# ------------------------------------------------------------
# Step 3: Fix - MS Office mode (NullFile)
# ------------------------------------------------------------
function Fix-MsOfficeExtension {
    param($Map)
    # Ensure extension default value points to the ProgID
    Set-RegString -Path $Map.Ext -Name '' -Value $Map.ProgID
    # Extension-level ShellNew must be removed
    [void](Remove-RegKeyIfExists "$($Map.Ext)\ShellNew")
    # ProgID-level ShellNew with NullFile
    $snPath = "$($Map.Ext)\$($Map.ProgID)\ShellNew"
    Set-RegString -Path $snPath -Name 'NullFile' -Value ''
    # Display name + remove FriendlyTypeName
    Set-RegString -Path $Map.ProgID -Name '' -Value $Map.ProgName
    [void](Remove-RegValueIfExists -Path $Map.ProgID -Name 'FriendlyTypeName')
    Log ("  - [{0}] 已按 MS Office 方案修复（NullFile）" -f $Map.Ext)
    return $true
}

# ------------------------------------------------------------
# Step 3: Fix - WPS mode (FileName + template file)
# ------------------------------------------------------------
function Prepare-WpsTemplates {
    param($WpsInfo)
    # Returns a hashtable: Ext -> template full path (only for available templates)
    Log ("  固定模板目录: {0}" -f $script:FixedTemplateDir)
    if (-not (Test-Path $script:FixedTemplateDir)) {
        [void](New-Item -ItemType Directory -Path $script:FixedTemplateDir -Force)
    }

    $savedVersion = ''
    if (Test-Path $script:WpsVersionFile) {
        $savedVersion = [string](Get-Content -Path $script:WpsVersionFile -ErrorAction SilentlyContinue | Select-Object -First 1)
    }

    $needCopy = $false
    if ($savedVersion -ne $WpsInfo.Version) {
        $needCopy = $true
        if ($savedVersion) { Log '  检测到 WPS 版本变化，重新复制模板...' }
        else { Log '  首次运行，复制 WPS 模板到固定目录...' }
    }
    foreach ($m in $script:ModernExtMap) {
        $target = Join-Path $script:FixedTemplateDir $m.TemplateFile
        if (-not (Test-Path $target)) { $needCopy = $true }
    }
    if ($needCopy) {
        foreach ($m in $script:ModernExtMap) {
            $src = Join-Path $WpsInfo.TemplateDir $m.TemplateFile
            $dst = Join-Path $script:FixedTemplateDir $m.TemplateFile
            if (Test-Path $src) {
                Copy-Item -Path $src -Destination $dst -Force
                Log ("  - [{0}] 模板已复制: {1}" -f $m.Ext, $m.TemplateFile)
            } else {
                Log ("  - [{0}] 源模板不存在，跳过复制: {1}" -f $m.Ext, $src)
            }
        }
        Set-Content -Path $script:WpsVersionFile -Value $WpsInfo.Version -Encoding ASCII
        Log ("  已记录 WPS 版本: {0}" -f $WpsInfo.Version)
    } else {
        Log '  WPS 版本未变化且模板完整，跳过复制。'
    }

    $result = @{}
    foreach ($m in $script:ModernExtMap) {
        $target = Join-Path $script:FixedTemplateDir $m.TemplateFile
        if (Test-Path $target) { $result[$m.Ext] = $target }
    }
    return $result
}

function Fix-WpsExtension {
    param($Map, [string]$TemplatePath)
    if (-not $TemplatePath) {
        Log ("  - [{0}] 无可用模板，跳过" -f $Map.Ext)
        return $false
    }
    # Ensure extension default value points to the ProgID
    Set-RegString -Path $Map.Ext -Name '' -Value $Map.ProgID
    # Extension-level ShellNew: FileName = template
    $snExt = "$($Map.Ext)\ShellNew"
    Set-RegString -Path $snExt -Name 'FileName' -Value $TemplatePath
    [void](Remove-RegValueIfExists -Path $snExt -Name 'NullFile')
    # ProgID-level ShellNew: FileName = template
    $snProg = "$($Map.Ext)\$($Map.ProgID)\ShellNew"
    Set-RegString -Path $snProg -Name 'FileName' -Value $TemplatePath
    [void](Remove-RegValueIfExists -Path $snProg -Name 'NullFile')
    # Display name + remove FriendlyTypeName
    Set-RegString -Path $Map.ProgID -Name '' -Value $Map.ProgName
    [void](Remove-RegValueIfExists -Path $Map.ProgID -Name 'FriendlyTypeName')
    Log ("  - [{0}] 已按 WPS 方案修复（FileName: {1}）" -f $Map.Ext, $Map.ProgName)
    return $true
}

# ------------------------------------------------------------
# Step 4: Verification
# ------------------------------------------------------------
function Test-MsOfficeFix {
    param($Map)
    $nullFile = Get-RegString -Path "$($Map.Ext)\$($Map.ProgID)\ShellNew" -Name 'NullFile'
    $displayName = Get-RegString -Path $Map.ProgID -Name ''
    return (($null -ne $nullFile) -and $displayName)
}

function Test-WpsFix {
    param($Map)
    $fn = Get-RegString -Path "$($Map.Ext)\ShellNew" -Name 'FileName'
    return ($fn -and (Test-Path $fn))
}

# ============================================================
# -DetectOnly mode: read-only detection, then exit
# ============================================================
if ($DetectOnly) {
    Log '==== 只读检测模式（不修改任何注册表） ===='
    $ms = Detect-MsOffice
    if ($ms.Found) { Log ("Microsoft Office: 已安装, 版本 {0}" -f $ms.Version) }
    else { Log 'Microsoft Office: 未检测到' }
    $wps = Detect-Wps
    if ($wps.Found) {
        Log ("WPS Office    : 已安装, 版本 {0}" -f $wps.Version)
        Log ("  安装目录: {0}" -f $wps.InstallPath)
        Log ("  模板目录: {0}" -f $wps.TemplateDir)
    } else { Log 'WPS Office    : 未检测到' }
    exit 0
}

# ------------------------------------------------------------
# Hide the PowerShell console window (GUI mode only;
# -DetectOnly keeps its console for readable output)
# ------------------------------------------------------------
if (-not ('W.ConsoleHider' -as [type])) {
    Add-Type -Namespace 'W' -Name 'ConsoleHider' -MemberDefinition `
        '[System.Runtime.InteropServices.DllImport("kernel32.dll")] public static extern System.IntPtr GetConsoleWindow();
         [System.Runtime.InteropServices.DllImport("user32.dll")] public static extern bool ShowWindow(System.IntPtr hWnd, int nCmdShow);'
}
$consoleHwnd = [W.ConsoleHider]::GetConsoleWindow()
if ($consoleHwnd -ne [IntPtr]::Zero) {
    [void][W.ConsoleHider]::ShowWindow($consoleHwnd, 0)  # SW_HIDE
}

# ============================================================
# Elevation check (GUI mode requires Administrator)
# ============================================================
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identity)
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Log '需要管理员权限，正在请求提权并重新启动...'
    try {
        Start-Process -FilePath 'powershell.exe' -Verb RunAs -WindowStyle Hidden -ArgumentList @(
            '-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-WindowStyle', 'Hidden', '-File', ('"{0}"' -f $PSCommandPath)
        )
    } catch {
        Write-Host '用户取消了提权，无法运行修复工具。' -ForegroundColor Red
    }
    exit 0
}

# ============================================================
# GUI
# ============================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$fontUi = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
$fontTitle = New-Object System.Drawing.Font('Microsoft YaHei UI', 14, [System.Drawing.FontStyle]::Bold)
$fontMono = New-Object System.Drawing.Font('Consolas', 9)

$form = New-Object System.Windows.Forms.Form
$form.Text = '右键新建菜单模板修复工具 v3.0'
$form.ClientSize = New-Object System.Drawing.Size(770, 608)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedSingle'
$form.MaximizeBox = $false
$form.Font = $fontUi

# ---- Title ----
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = '右键新建菜单模板修复工具'
$lblTitle.Font = $fontTitle
$lblTitle.Location = New-Object System.Drawing.Point(20, 12)
$lblTitle.Size = New-Object System.Drawing.Size(400, 28)
$form.Controls.Add($lblTitle)

$lblSub = New-Object System.Windows.Forms.Label
$lblSub.Text = '清理并修复 .xlsx / .pptx / .docx 右键新建菜单（支持 Microsoft Office 与 WPS 双方案）'
$lblSub.ForeColor = [System.Drawing.Color]::Gray
$lblSub.Location = New-Object System.Drawing.Point(22, 42)
$lblSub.Size = New-Object System.Drawing.Size(720, 18)
$form.Controls.Add($lblSub)

# ---- Environment detection group ----
$grpEnv = New-Object System.Windows.Forms.GroupBox
$grpEnv.Text = '环境检测（第 2 步结果）'
$grpEnv.Location = New-Object System.Drawing.Point(20, 68)
$grpEnv.Size = New-Object System.Drawing.Size(730, 96)
$form.Controls.Add($grpEnv)

$lblMsCap = New-Object System.Windows.Forms.Label
$lblMsCap.Text = 'Microsoft Office:'
$lblMsCap.Location = New-Object System.Drawing.Point(15, 30)
$lblMsCap.Size = New-Object System.Drawing.Size(130, 18)
$grpEnv.Controls.Add($lblMsCap)

$lblMs = New-Object System.Windows.Forms.Label
$lblMs.Text = '尚未检测'
$lblMs.ForeColor = [System.Drawing.Color]::Gray
$lblMs.Location = New-Object System.Drawing.Point(150, 30)
$lblMs.Size = New-Object System.Drawing.Size(560, 18)
$grpEnv.Controls.Add($lblMs)

$lblWpsCap = New-Object System.Windows.Forms.Label
$lblWpsCap.Text = 'WPS Office:'
$lblWpsCap.Location = New-Object System.Drawing.Point(15, 58)
$lblWpsCap.Size = New-Object System.Drawing.Size(130, 18)
$grpEnv.Controls.Add($lblWpsCap)

$lblWps = New-Object System.Windows.Forms.Label
$lblWps.Text = '尚未检测'
$lblWps.ForeColor = [System.Drawing.Color]::Gray
$lblWps.Location = New-Object System.Drawing.Point(150, 58)
$lblWps.Size = New-Object System.Drawing.Size(560, 18)
$grpEnv.Controls.Add($lblWps)

# ---- Fix mode selection group ----
$grpMode = New-Object System.Windows.Forms.GroupBox
$grpMode.Text = '修复方案选择（第 3 步）'
$grpMode.Location = New-Object System.Drawing.Point(20, 172)
$grpMode.Size = New-Object System.Drawing.Size(730, 92)
$form.Controls.Add($grpMode)

$rbMs = New-Object System.Windows.Forms.RadioButton
$rbMs.Text = '1. Microsoft Office 修复（NullFile 方案）'
$rbMs.Location = New-Object System.Drawing.Point(15, 26)
$rbMs.Size = New-Object System.Drawing.Size(320, 22)
$rbMs.Enabled = $false
$grpMode.Controls.Add($rbMs)

$rbWps = New-Object System.Windows.Forms.RadioButton
$rbWps.Text = '2. WPS 修复（模板 FileName 方案）'
$rbWps.Location = New-Object System.Drawing.Point(15, 54)
$rbWps.Size = New-Object System.Drawing.Size(320, 22)
$rbWps.Enabled = $false
$grpMode.Controls.Add($rbWps)

$lblModeHint = New-Object System.Windows.Forms.Label
$lblModeHint.Text = '请先点击「① 检测并清理旧右键菜单」开始'
$lblModeHint.ForeColor = [System.Drawing.Color]::Gray
$lblModeHint.Location = New-Object System.Drawing.Point(345, 36)
$lblModeHint.Size = New-Object System.Drawing.Size(370, 40)
$grpMode.Controls.Add($lblModeHint)

# ---- Log group ----
$grpLog = New-Object System.Windows.Forms.GroupBox
$grpLog.Text = '操作日志'
$grpLog.Location = New-Object System.Drawing.Point(20, 272)
$grpLog.Size = New-Object System.Drawing.Size(730, 274)
$form.Controls.Add($grpLog)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ReadOnly = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.Font = $fontMono
$txtLog.BackColor = [System.Drawing.Color]::White
$txtLog.Location = New-Object System.Drawing.Point(10, 22)
$txtLog.Size = New-Object System.Drawing.Size(708, 240)
$grpLog.Controls.Add($txtLog)
$script:txtLog = $txtLog

# ---- Buttons ----
$btnStep1 = New-Object System.Windows.Forms.Button
$btnStep1.Text = '① 检测并清理旧右键菜单'
$btnStep1.Location = New-Object System.Drawing.Point(20, 558)
$btnStep1.Size = New-Object System.Drawing.Size(230, 38)
$form.Controls.Add($btnStep1)

$btnFix = New-Object System.Windows.Forms.Button
$btnFix.Text = '② 执行修复'
$btnFix.Location = New-Object System.Drawing.Point(270, 558)
$btnFix.Size = New-Object System.Drawing.Size(230, 38)
$btnFix.Enabled = $false
$form.Controls.Add($btnFix)

$btnExit = New-Object System.Windows.Forms.Button
$btnExit.Text = '退出'
$btnExit.Location = New-Object System.Drawing.Point(620, 558)
$btnExit.Size = New-Object System.Drawing.Size(130, 38)
$form.Controls.Add($btnExit)

# ---- State ----
$script:MsInfo = $null
$script:WpsInfo = $null
$script:FixMode = ''   # 'msoffice' | 'wps'

# ------------------------------------------------------------
# Step 1 + 2 handler
# ------------------------------------------------------------
$btnStep1.Add_Click({
    $btnStep1.Enabled = $false
    $btnFix.Enabled = $false
    $rbMs.Enabled = $false
    $rbWps.Enabled = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        Log '================ 第 1 步：清理旧右键新建菜单 ================'
        Log '正在检测注册表中的 Office / WPS 右键新建项...'
        $cleaned = Clean-ShellNewEntries
        Log ("本次共清理 {0} 项。" -f $cleaned)

        Log '正在刷新 ShellNew 缓存...'
        [void](Clear-OfficeShellNewCache)
        Log '缓存已刷新。'

        Restart-Explorer

        # ---- User verification loop ----
        $verifyMsg = @'
请现在在桌面或任意文件夹空白处点击右键，查看「新建」子菜单。

以下项目应已全部消失：
  Microsoft Excel 工作簿 (.xlsx)
  Microsoft PowerPoint 演示文稿 (.pptx)
  Microsoft Word 文档 (.docx)
  Excel 97-2003 工作簿 (.xls)
  PowerPoint 97-2003 演示文稿 (.ppt)
  Word 97-2003 文档 (.doc)

是否已确认清理完毕？

[是] 已确认，进入后续步骤
[否] 重新检查注册表并再次清理
'@
        while ($true) {
            $dr = [System.Windows.Forms.MessageBox]::Show($form, $verifyMsg,
                '用户确认 - 第 1/4 步', [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($dr -eq [System.Windows.Forms.DialogResult]::Yes) { break }

            Log '正在复查注册表残留...'
            $remaining = Recheck-AndCleanShellNew
            if ($remaining -gt 0) {
                Log ("复查发现并清理了 {0} 个扩展名的残留项。" -f $remaining)
                [void](Clear-OfficeShellNewCache)
                Restart-Explorer
            } else {
                Log '注册表已确认无残留项，直接继续后续步骤。'
                break
            }
        }

        # ---- Step 2: Detection ----
        Log ''
        Log '================ 第 2 步：检测 Office 安装情况 ================'
        $script:MsInfo = Detect-MsOffice
        $script:WpsInfo = Detect-Wps

        if ($script:MsInfo.Found) {
            $lblMs.Text = ('已安装，版本 {0}' -f $script:MsInfo.Version)
            $lblMs.ForeColor = [System.Drawing.Color]::Green
            Log ("  - 检测到 Microsoft Office，版本: {0}" -f $script:MsInfo.Version)
        } else {
            $lblMs.Text = '未检测到'
            $lblMs.ForeColor = [System.Drawing.Color]::Firebrick
            Log '  - 未检测到 Microsoft Office'
        }
        if ($script:WpsInfo.Found) {
            $lblWps.Text = ('已安装，版本 {0}' -f $script:WpsInfo.Version)
            $lblWps.ForeColor = [System.Drawing.Color]::Green
            Log ("  - 检测到 WPS Office，版本: {0}" -f $script:WpsInfo.Version)
            Log ("    安装目录: {0}" -f $script:WpsInfo.InstallPath)
        } else {
            $lblWps.Text = '未检测到'
            $lblWps.ForeColor = [System.Drawing.Color]::Firebrick
            Log '  - 未检测到 WPS Office'
        }

        # ---- Decide fix mode ----
        Log ''
        if (-not $script:MsInfo.Found -and -not $script:WpsInfo.Found) {
            Log '未检测到任何 Office 软件，无法进行修复，程序流程终止。'
            $lblModeHint.Text = '未检测到 Microsoft Office 和 WPS，无法修复'
            $lblModeHint.ForeColor = [System.Drawing.Color]::Firebrick
            [void][System.Windows.Forms.MessageBox]::Show($form,
                '未检测到 Microsoft Office 或 WPS Office，无需修复，工具即将退出。',
                '提示', [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
            $form.Close()
            return
        }
        elseif ($script:MsInfo.Found -and -not $script:WpsInfo.Found) {
            $script:FixMode = 'msoffice'
            $rbMs.Checked = $true
            $lblModeHint.Text = '仅检测到 Microsoft Office，已自动选择 MS Office 方案'
            $lblModeHint.ForeColor = [System.Drawing.Color]::Green
            Log '仅安装了 Microsoft Office，已自动选择 MS Office 修复方案。'
        }
        elseif ($script:WpsInfo.Found -and -not $script:MsInfo.Found) {
            $script:FixMode = 'wps'
            $rbWps.Checked = $true
            $lblModeHint.Text = '仅检测到 WPS Office，已自动选择 WPS 方案'
            $lblModeHint.ForeColor = [System.Drawing.Color]::Green
            Log '仅安装了 WPS Office，已自动选择 WPS 修复方案。'
        }
        else {
            $script:FixMode = ''
            $lblModeHint.Text = '双安装：请选择 1 或 2，再点击「② 执行修复」'
            $lblModeHint.ForeColor = [System.Drawing.Color]::DarkOrange
            $rbMs.Enabled = $true
            $rbWps.Enabled = $true
            $rbMs.Checked = $true
            Log '检测到 Microsoft Office 与 WPS 均已安装，请在上方选择修复方案。'
            [void][System.Windows.Forms.MessageBox]::Show($form,
                "检测到 Microsoft Office 和 WPS Office 均已安装。`r`n`r`nMicrosoft Office 版本: $($script:MsInfo.Version)`r`nWPS Office 版本: $($script:WpsInfo.Version)`r`n`r`n请在窗口中选择 [1] Microsoft Office 或 [2] WPS 修复方案，然后点击「② 执行修复」。",
                '方案选择 - 第 3/4 步', [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information)
        }

        if ($script:FixMode -ne '') { $btnFix.Enabled = $true }
        elseif ($rbMs.Enabled) { $btnFix.Enabled = $true }
    }
    catch {
        Log ('[错误] {0}' -f $_.Exception.Message)
        [void][System.Windows.Forms.MessageBox]::Show($form,
            ('执行过程中出现错误：{0}' -f $_.Exception.Message),
            '错误', [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnStep1.Enabled = $true
    }
})

# ------------------------------------------------------------
# Radio buttons keep $script:FixMode in sync (double-install case)
# ------------------------------------------------------------
$rbMs.Add_CheckedChanged({
    if ($rbMs.Checked -and $rbMs.Enabled) {
        $script:FixMode = 'msoffice'
        $lblModeHint.Text = '已选择: 1. Microsoft Office 修复（NullFile 方案）'
        $lblModeHint.ForeColor = [System.Drawing.Color]::DarkOrange
    }
})
$rbWps.Add_CheckedChanged({
    if ($rbWps.Checked -and $rbWps.Enabled) {
        $script:FixMode = 'wps'
        $lblModeHint.Text = '已选择: 2. WPS 修复（模板 FileName 方案）'
        $lblModeHint.ForeColor = [System.Drawing.Color]::DarkOrange
    }
})

# ------------------------------------------------------------
# Step 3 + 4 handler
# ------------------------------------------------------------
$btnFix.Add_Click({
    # Resolve mode (double-install case reads radio selection)
    if ($rbMs.Enabled -and $rbMs.Checked) { $script:FixMode = 'msoffice' }
    elseif ($rbWps.Enabled -and $rbWps.Checked) { $script:FixMode = 'wps' }

    if ([string]::IsNullOrEmpty($script:FixMode)) {
        [void][System.Windows.Forms.MessageBox]::Show($form, '请先选择修复方案（1 或 2）。',
            '提示', [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $btnFix.Enabled = $false
    $btnStep1.Enabled = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        Log ''
        Log '================ 第 3 步：执行修复 ================'
        if ($script:FixMode -eq 'msoffice') {
            Log '修复模式: Microsoft Office（NullFile 方案）'
            $fixCount = 0
            foreach ($m in $script:ModernExtMap) {
                if (Fix-MsOfficeExtension -Map $m) { $fixCount++ }
            }
        }
        else {
            Log '修复模式: WPS Office（模板 FileName 方案）'
            $templates = Prepare-WpsTemplates -WpsInfo $script:WpsInfo
            $fixCount = 0
            foreach ($m in $script:ModernExtMap) {
                if (Fix-WpsExtension -Map $m -TemplatePath $templates[$m.Ext]) { $fixCount++ }
            }
        }
        Log ("共修复 {0} / 3 个扩展名。" -f $fixCount)

        Log '正在刷新 ShellNew 缓存并广播更改...'
        [void](Refresh-OfficeShellNewCache)
        [void](Send-ShellNotify)
        Log '缓存已刷新，已发送系统通知。'

        # ---- Step 4: Verify ----
        Log ''
        Log '================ 第 4 步：验证修复结果 ================'
        $okCount = 0
        foreach ($m in $script:ModernExtMap) {
            if ($script:FixMode -eq 'msoffice') {
                if (Test-MsOfficeFix -Map $m) { Log ("  - [{0}] 验证通过" -f $m.Ext); $okCount++ }
                else { Log ("  - [{0}] 验证失败（ShellNew NullFile 或显示名缺失）" -f $m.Ext) }
            } else {
                if (Test-WpsFix -Map $m) { Log ("  - [{0}] 验证通过" -f $m.Ext); $okCount++ }
                else { Log ("  - [{0}] 验证失败（ShellNew FileName 缺失）" -f $m.Ext) }
            }
        }
        Log ("验证结果: {0} / 3 通过。" -f $okCount)

        Restart-Explorer

        $modeText = 'Microsoft Office（NullFile）'
        if ($script:FixMode -eq 'wps') { $modeText = 'WPS Office（模板 FileName）' }
        $summary = "修复完成！`r`n`r`n修复方案: {0}`r`n修复扩展名: {1} / 3`r`n验证通过: {2} / 3`r`n资源管理器已重启" -f $modeText, $fixCount, $okCount
        if ($script:FixMode -eq 'wps') {
            $summary += "`r`n模板目录: $($script:FixedTemplateDir)"
        }
        $summary += "`r`n`r`n请右键 -> 新建，确认 .xlsx / .pptx / .docx 菜单已正常显示。"
        Log $summary
        [void][System.Windows.Forms.MessageBox]::Show($form, $summary,
            '修复完成', [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch {
        Log ('[错误] {0}' -f $_.Exception.Message)
        [void][System.Windows.Forms.MessageBox]::Show($form,
            ('修复过程中出现错误：{0}' -f $_.Exception.Message),
            '错误', [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$btnExit.Add_Click({ $form.Close() })

Log '工具已就绪。点击「① 检测并清理旧右键菜单」开始修复流程。'
Log '流程: 清理旧项 -> 用户确认 -> 检测安装 -> 选择方案 -> 修复 -> 验证 -> 重启资源管理器'
Log ''

[void]$form.ShowDialog()
