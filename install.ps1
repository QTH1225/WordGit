[CmdletBinding()]
param(
    [switch]$Uninstall,
    [switch]$CheckOnly,
    [switch]$Plan,
    [switch]$LibraryOnly,
    [string]$SetupExecutable
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object Text.UTF8Encoding($false)
# Must match the [Guid]/[ProgId] attributes in src/AddIn.cs and src/HistoryPane.cs;
# tests/Installer.Tests.ps1 compares them with the built assembly.
$WordGitConnectClsid = '{C5707F76-F46F-4A11-800B-18E7D98377D3}'
$WordGitPaneClsid = '{F55EC3C1-B725-4ED8-9CA6-ED51D07C3361}'
$WordGitConnectProgId = 'WordGit.Connect'
$WordGitPaneProgId = 'WordGit.HistoryPane'

function Get-WordGitRegistryViews {
    if ([Environment]::Is64BitOperatingSystem) { [Microsoft.Win32.RegistryView]::Registry64 }
    [Microsoft.Win32.RegistryView]::Registry32
}

function Get-WordGitEnvironment {
    $wordPath = $null
    foreach ($view in (Get-WordGitRegistryViews)) {
        foreach ($hive in @([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryHive]::CurrentUser)) {
            $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hive, $view)
            try {
                $key = $base.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE')
                if ($key) {
                    try {
                        $candidate = ([string]$key.GetValue('')).Trim('"')
                        if ($candidate -and (Test-Path -LiteralPath $candidate -PathType Leaf)) { $wordPath = $candidate }
                    } finally { $key.Dispose() }
                }
            } finally { $base.Dispose() }
            if ($wordPath) { break }
        }
        if ($wordPath) { break }
    }
    if (-not $wordPath) { throw '未找到 Windows 桌面版 Microsoft Word。请先安装 Word。' }
    $reader = New-Object IO.BinaryReader([IO.File]::OpenRead($wordPath))
    try {
        if ($reader.ReadUInt16() -ne 0x5A4D) { throw 'Word 程序格式无效。' }
        $reader.BaseStream.Position = 0x3C
        $offset = $reader.ReadInt32()
        $reader.BaseStream.Position = $offset
        if ($reader.ReadUInt32() -ne 0x4550) { throw '无法识别 Word 程序架构。' }
        $machine = $reader.ReadUInt16()
    } finally { $reader.Dispose() }
    if ($machine -eq 0x8664) { $wordBits = 64 }
    elseif ($machine -eq 0x14C) { $wordBits = 32 }
    else { throw '当前安装程序支持 x86/x64 Word，尚未验证 ARM 原生 Word。' }

    $release = 0
    foreach ($view in (Get-WordGitRegistryViews)) {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $view)
        try {
            $key = $base.OpenSubKey('SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full')
            if ($key) {
                try { $release = [Math]::Max($release, [int]$key.GetValue('Release', 0)) }
                finally { $key.Dispose() }
            }
        } finally { $base.Dispose() }
    }
    if ($release -lt 528040) { throw '需要 .NET Framework 4.8 或更新版本。下载：https://dotnet.microsoft.com/download/dotnet-framework/net48' }

    $git = $null
    $explicit = [Environment]::GetEnvironmentVariable('WORDGIT_GIT')
    if (-not $explicit) { $explicit = [Environment]::GetEnvironmentVariable('WORDGIT_GIT', 'User') }
    if ($explicit) {
        if (-not [IO.Path]::IsPathRooted($explicit) -or -not (Test-Path -LiteralPath $explicit -PathType Leaf)) {
            throw 'WORDGIT_GIT 必须是有效 git.exe 的绝对路径。'
        }
        $git = $explicit
    } else {
        $command = Get-Command git.exe -ErrorAction SilentlyContinue
        if ($command) { $git = $command.Source }
        foreach ($folder in @($env:ProgramFiles, ${env:ProgramFiles(x86)}, (Join-Path $env:LOCALAPPDATA 'Programs'))) {
            if ($git -or -not $folder) { continue }
            $candidate = Join-Path $folder 'Git\cmd\git.exe'
            if (Test-Path -LiteralPath $candidate -PathType Leaf) { $git = $candidate }
        }
    }
    if (-not $git) { throw '未找到 Git for Windows。请先安装 2.51 或更新版本：https://git-scm.com/install/windows' }
    $gitVersionText = (& $git --version 2>&1 | Out-String).Trim()
    if ($LASTEXITCODE -ne 0 -or $gitVersionText -notmatch 'git version (\d+\.\d+\.\d+)') {
        throw 'Git 无法运行，请重新安装 Git for Windows。'
    }
    $gitVersion = [version]$Matches[1]
    if ($gitVersion -lt [version]'2.51.0') { throw "Git 版本为 $gitVersion，需要 2.51 或更新版本。" }
    [PSCustomObject]@{ WordPath = $wordPath; WordBits = $wordBits; GitPath = $git; GitVersion = $gitVersion.ToString(); FrameworkRelease = $release }
}

function Get-WordGitInstallDirectory {
    $parent = [IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA 'WordGit'))
    $target = [IO.Path]::GetFullPath((Join-Path $parent 'app'))
    if (-not $target.StartsWith($parent + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase) -or
        [IO.Path]::GetFileName($target) -ne 'app') { throw '安装目录校验失败。' }
    $target
}

function Set-WordGitRegistryValue($Base, [string]$Path, [string]$Name, $Value, [Microsoft.Win32.RegistryValueKind]$Kind = 'String') {
    $key = $Base.CreateSubKey($Path)
    try { $key.SetValue($Name, $Value, $Kind) } finally { $key.Dispose() }
}

function Register-WordGitClass($Base, [string]$ClassId, [string]$ProgId, [string]$Assembly, [string]$Version, [string]$CodeBase, [string]$Threading) {
    $class = "Software\Classes\CLSID\$ClassId"
    $server = "$class\InprocServer32"
    Set-WordGitRegistryValue $Base $class '' $ProgId
    Set-WordGitRegistryValue $Base "$class\ProgId" '' $ProgId
    Set-WordGitRegistryValue $Base "Software\Classes\$ProgId" '' $ProgId
    Set-WordGitRegistryValue $Base "Software\Classes\$ProgId\CLSID" '' $ClassId
    Set-WordGitRegistryValue $Base $server '' 'mscoree.dll'
    Set-WordGitRegistryValue $Base $server 'ThreadingModel' $Threading
    foreach ($path in @($server, "$server\$Version")) {
        Set-WordGitRegistryValue $Base $path 'Class' $ProgId
        Set-WordGitRegistryValue $Base $path 'Assembly' $Assembly
        Set-WordGitRegistryValue $Base $path 'RuntimeVersion' 'v4.0.30319'
        Set-WordGitRegistryValue $Base $path 'CodeBase' $CodeBase
    }
}

function Register-WordGitPlugin($Base, [string]$Assembly, [string]$Version, [string]$CodeBase) {
    Register-WordGitClass $Base $WordGitConnectClsid $WordGitConnectProgId $Assembly $Version $CodeBase 'Both'
    Register-WordGitClass $Base $WordGitPaneClsid $WordGitPaneProgId $Assembly $Version $CodeBase 'Apartment'
    $control = $Base.CreateSubKey("Software\Classes\CLSID\$WordGitPaneClsid\Control")
    $control.Dispose()
    $addin = "Software\Microsoft\Office\Word\Addins\$WordGitConnectProgId"
    Set-WordGitRegistryValue $Base $addin 'FriendlyName' 'WordGit'
    Set-WordGitRegistryValue $Base $addin 'Description' 'Word 本地 Git 版本管理'
    Set-WordGitRegistryValue $Base $addin 'LoadBehavior' 3 'DWord'
    # WPS 文字现在通过 JS 加载项工作；早期版本把同一个 COM 加载项也注册给 WPS，
    # 升级时清掉，否则 WPS 会同时加载两个同名功能区。
    Remove-WordGitLegacyWps $Base
}

function Remove-WordGitLegacyWps($Base) {
    $Base.DeleteSubKeyTree("Software\Kingsoft\Office\WPS\Addins\$WordGitConnectProgId", $false)
    $list = $Base.OpenSubKey('Software\Kingsoft\Office\WPS\AddinsWL', $true)
    if (-not $list) { return }
    try {
        if ($list.GetValueNames() -contains $WordGitConnectProgId) { $list.DeleteValue($WordGitConnectProgId, $false) }
    } finally { $list.Dispose() }
}

function Remove-WordGitRegistration($Base) {
    foreach ($path in @(
        "Software\Microsoft\Office\Word\Addins\$WordGitConnectProgId",
        "Software\Microsoft\Office\Word\AddinsData\$WordGitConnectProgId",
        "Software\Classes\CLSID\$WordGitConnectClsid",
        "Software\Classes\$WordGitConnectProgId",
        "Software\Classes\CLSID\$WordGitPaneClsid",
        "Software\Classes\$WordGitPaneProgId",
        'Software\Microsoft\Windows\CurrentVersion\Uninstall\WordGit'
    )) { $Base.DeleteSubKeyTree($path, $false) }
    try {
        $run = $Base.OpenSubKey('Software\Microsoft\Windows\CurrentVersion\Run', $true)
        if ($run) { try { $run.DeleteValue('WordGit', $false) } finally { $run.Dispose() } }
    } catch [System.Security.SecurityException] { }
    Remove-WordGitLegacyWps $Base
}

function Stop-WordGitKernel([string]$InstallDirectory) {
    $root = [IO.Path]::GetFullPath((Assert-WordGitAppDirectory $InstallDirectory)) + [IO.Path]::DirectorySeparatorChar
    foreach ($running in @(Get-Process -Name 'wordgit' -ErrorAction SilentlyContinue)) {
        $path = $null
        try { $path = $running.Path } catch [System.ComponentModel.Win32Exception] { continue } catch [System.InvalidOperationException] { continue }
        if ($path -and [IO.Path]::GetFullPath($path).StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) {
            try { Stop-Process -Id $running.Id -Force -ErrorAction Stop } catch { }
            try { [void]$running.WaitForExit(5000) } catch { }
        }
    }
}

function Register-WordGitKernel($Base, [string]$InstallDirectory) {
    $destination = Assert-WordGitAppDirectory $InstallDirectory
    $kernel = Join-Path $destination 'wordgit.exe'
    $run = $Base.CreateSubKey('Software\Microsoft\Windows\CurrentVersion\Run')
    try { $run.SetValue('WordGit', '"' + $kernel + '" serve', [Microsoft.Win32.RegistryValueKind]::String) }
    finally { $run.Dispose() }
    $kernel
}

# Starts the core now so the first ribbon click does not wait for it. A kernel left behind by an
# older build (for example a development copy) cannot serve the new endpoints, so it is replaced.
function Start-WordGitKernel([string]$Kernel) {
    if (-not (Test-Path -LiteralPath $Kernel -PathType Leaf)) { return }
    foreach ($running in @(Get-Process -Name 'wordgit' -ErrorAction SilentlyContinue)) {
        $path = $null
        try { $path = $running.Path } catch [System.ComponentModel.Win32Exception] { continue } catch [System.InvalidOperationException] { continue }
        if ($path -and [IO.Path]::GetFullPath($path) -eq [IO.Path]::GetFullPath($Kernel)) { return }
        try { Stop-Process -Id $running.Id -Force -ErrorAction Stop } catch { }
    }
    try {
        Start-Process -FilePath $Kernel -ArgumentList 'serve' -WindowStyle Hidden -ErrorAction Stop | Out-Null
    } catch {
        Write-Output '内核未能自动启动，首次使用功能区按钮时会再次尝试。'
    }
}

function Install-WordGitWpsAddin([string]$InstallDirectory) {
    $destination = Assert-WordGitAppDirectory $InstallDirectory
    $addin = Join-Path $destination 'addin'
    $kernel = Join-Path $destination 'wordgit.exe'
    if (-not (Test-Path -LiteralPath (Join-Path $addin 'ribbon.xml') -PathType Leaf)) { return }
    if (-not (Test-Path -LiteralPath $kernel -PathType Leaf)) { return }
    # The kernel knows where WPS keeps its add-ins; it also records the service file there, which is
    # what makes the add-in find the running service token.
    & $kernel install --addin $addin --no-autostart
    if ($LASTEXITCODE -ne 0) { Write-Output 'WPS 文字加载项未能自动更新，重新打开 WPS 后可再次运行安装程序。' }
}

function Test-WordGitPayload {
    $source = Join-Path $PSScriptRoot 'bin\WordGit.dll'
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw '缺少 bin\WordGit.dll。请完整解压安装包；从源码安装时先运行 build.ps1。' }
    $identity = [Reflection.AssemblyName]::GetAssemblyName($source)
    if ($identity.Name -ne 'WordGit') { throw '插件程序集名称不正确。' }
    $kernel = Join-Path $PSScriptRoot 'bin\wordgit.exe'
    if (-not (Test-Path -LiteralPath $kernel -PathType Leaf)) { throw '缺少 bin\wordgit.exe（Rust 内核）。请完整解压安装包；从源码安装时先运行 build.ps1。' }
    $manifestFile = Join-Path $PSScriptRoot 'manifest.json'
    if (Test-Path -LiteralPath $manifestFile) {
        $manifest = Get-Content -Raw -LiteralPath $manifestFile -Encoding UTF8 | ConvertFrom-Json
        if ($manifest.Version -ne $identity.Version.ToString()) { throw '安装包版本与插件版本不一致。' }
        $root = [IO.Path]::GetFullPath($PSScriptRoot) + [IO.Path]::DirectorySeparatorChar
        foreach ($entry in $manifest.Files) {
            $target = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot $entry.Path))
            if (-not $target.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) { throw '安装包包含无效文件路径。' }
            if (-not (Test-Path -LiteralPath $target -PathType Leaf) -or (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash -ne $entry.Sha256) {
                throw "安装文件校验失败：$($entry.Path)"
            }
        }
    }
    $identity
}

function Assert-WordGitAppDirectory([string]$Path) {
    $full = [IO.Path]::GetFullPath($Path).TrimEnd([IO.Path]::DirectorySeparatorChar)
    if ([IO.Path]::GetFileName($full) -ne 'app' -or [IO.Path]::GetFileName([IO.Path]::GetDirectoryName($full)) -ne 'WordGit') {
        throw '操作目录必须是 WordGit\\app；不会处理仓库或审阅稿目录。'
    }
    $full
}

function Install-WordGitFiles([string]$InstallDirectory, [string]$SourceDirectory, $Identity, [string]$SetupPath) {
    $destination = Assert-WordGitAppDirectory $InstallDirectory
    $source = Join-Path $SourceDirectory 'bin\WordGit.dll'
    $hash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
    $installed = Join-Path $destination ("WordGit." + $Identity.Version.ToString() + "." + $hash.Substring(0,12) + ".dll")
    New-Item -ItemType Directory -Path $destination -Force | Out-Null
    Copy-Item -LiteralPath $source -Destination $installed -Force
    if ((Get-FileHash -LiteralPath $installed -Algorithm SHA256).Hash -ne $hash) { throw '安装后的插件校验失败。' }
    $kernelSource = Join-Path $SourceDirectory 'bin\wordgit.exe'
    $kernelTarget = Join-Path $destination 'wordgit.exe'
    $kernelHash = (Get-FileHash -LiteralPath $kernelSource -Algorithm SHA256).Hash
    Copy-Item -LiteralPath $kernelSource -Destination $kernelTarget -Force
    if ((Get-FileHash -LiteralPath $kernelTarget -Algorithm SHA256).Hash -ne $kernelHash) { throw '安装后的内核校验失败。' }
    foreach ($name in @('install.ps1','Uninstall.cmd')) {
        $from = Join-Path $SourceDirectory $name
        $to = Join-Path $destination $name
        if ((Test-Path -LiteralPath $from -PathType Leaf) -and [IO.Path]::GetFullPath($from) -ne [IO.Path]::GetFullPath($to)) {
            Copy-Item -LiteralPath $from -Destination $to -Force
        }
    }
    # The WPS Writer add-in ships next to the kernel so `wordgit install` can register it.
    $addinSource = Join-Path $SourceDirectory 'addin-wps'
    $addinTarget = Join-Path $destination 'addin'
    if (Test-Path -LiteralPath (Join-Path $addinSource 'ribbon.xml') -PathType Leaf) {
        if (Test-Path -LiteralPath $addinTarget) { Remove-Item -LiteralPath $addinTarget -Recurse -Force }
        Copy-Item -LiteralPath $addinSource -Destination $addinTarget -Recurse -Force
    }
    if ($SetupPath -and (Test-Path -LiteralPath $SetupPath -PathType Leaf)) {
        $targetSetup = Join-Path $destination 'WordGit.Setup.exe'
        if ([IO.Path]::GetFullPath($SetupPath) -ne [IO.Path]::GetFullPath($targetSetup)) {
            Copy-Item -LiteralPath $SetupPath -Destination $targetSetup -Force
        }
    }
    $installed
}

function Remove-WordGitFiles([string]$InstallDirectory) {
    $destination = Assert-WordGitAppDirectory $InstallDirectory
    if (-not (Test-Path -LiteralPath $destination)) { return }
    Get-ChildItem -LiteralPath $destination -File | Where-Object {
        $_.Name -match '^WordGit\.\d+\.\d+\.\d+\.\d+\.[A-Fa-f0-9]{12}\.dll$' -or
        $_.Name -in @('WordGit.dll','wordgit.exe','install.ps1','Uninstall.cmd','WordGit.Setup.exe')
    } | ForEach-Object {
        try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop }
        catch [IO.IOException] { Write-Output '运行中的维护程序将在退出后清理。' }
        catch [UnauthorizedAccessException] { Write-Output '运行中的维护程序将在退出后清理。' }
    }
    $addin = Join-Path $destination 'addin'
    if (Test-Path -LiteralPath $addin) { try { Remove-Item -LiteralPath $addin -Recurse -Force -ErrorAction Stop } catch { Write-Output '保留仍在使用的加载项文件：addin' } }
    if (-not (Get-ChildItem -LiteralPath $destination -Force | Select-Object -First 1)) {
        Remove-Item -LiteralPath $destination -Force
    }
}

function Register-WordGitProgram($Base, [string]$InstallDirectory, [string]$Version) {
    $destination = Assert-WordGitAppDirectory $InstallDirectory
    $program = 'Software\Microsoft\Windows\CurrentVersion\Uninstall\WordGit'
    $installedSetup = Join-Path $destination 'WordGit.Setup.exe'
    $powerShell = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $command = if (Test-Path -LiteralPath $installedSetup -PathType Leaf) { '"' + $installedSetup + '" /uninstall' }
        else { '"' + $powerShell + '" -NoProfile -ExecutionPolicy Bypass -File "' + (Join-Path $destination 'install.ps1') + '" -Uninstall' }
    Set-WordGitRegistryValue $Base $program 'DisplayName' 'WordGit'
    Set-WordGitRegistryValue $Base $program 'DisplayVersion' $Version
    Set-WordGitRegistryValue $Base $program 'Publisher' 'WordGit'
    Set-WordGitRegistryValue $Base $program 'InstallLocation' $destination
    Set-WordGitRegistryValue $Base $program 'UninstallString' $command
    if (Test-Path -LiteralPath $installedSetup -PathType Leaf) { Set-WordGitRegistryValue $Base $program 'DisplayIcon' $installedSetup }
    Set-WordGitRegistryValue $Base $program 'NoModify' 1 'DWord'
    Set-WordGitRegistryValue $Base $program 'NoRepair' 1 'DWord'
}

if ($LibraryOnly) { return }
if ($Uninstall -and ($CheckOnly -or $Plan)) { throw '卸载和检查参数不能同时使用。' }
$installDir = Get-WordGitInstallDirectory
if ($Uninstall) {
    if (Get-Process -Name WINWORD -ErrorAction SilentlyContinue) { throw '请先保存并关闭所有 Word 窗口，再卸载 WordGit。' }
    foreach ($view in (Get-WordGitRegistryViews)) {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
        try { Remove-WordGitRegistration $base } finally { $base.Dispose() }
    }
    foreach ($running in @(Get-Process -Name 'wordgit' -ErrorAction SilentlyContinue)) {
        $path = $null
        try { $path = $running.Path } catch [System.ComponentModel.Win32Exception] { continue } catch [System.InvalidOperationException] { continue }
        if ($path -and [IO.Path]::GetFullPath($path).StartsWith([IO.Path]::GetFullPath($installDir) + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
            try { Stop-Process -Id $running.Id -Force -ErrorAction Stop } catch { }
        }
    }
    $kernel = Join-Path $installDir 'wordgit.exe'
    if (Test-Path -LiteralPath $kernel -PathType Leaf) { try { & $kernel uninstall | Out-Null } catch { } }
    Remove-WordGitFiles $installDir
    Write-Output '已卸载 WordGit（含 WPS 文字加载项）。文档、repositories 历史仓库、sessions 审阅稿和 logs 日志均保留。'
    return
}
$environmentInfo = Get-WordGitEnvironment
$identity = Test-WordGitPayload
$version = $identity.Version.ToString()
$source = Join-Path $PSScriptRoot 'bin\WordGit.dll'
$hash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
$installed = Join-Path $installDir "WordGit.$version.$($hash.Substring(0,12)).dll"
$view = if ($environmentInfo.WordBits -eq 64) { [Microsoft.Win32.RegistryView]::Registry64 } else { [Microsoft.Win32.RegistryView]::Registry32 }
if ($Plan) {
    [PSCustomObject]@{ Version = $version; WordBits = $environmentInfo.WordBits; WordPath = $environmentInfo.WordPath;
        GitPath = $environmentInfo.GitPath; GitVersion = $environmentInfo.GitVersion; FrameworkRelease = $environmentInfo.FrameworkRelease;
        RegistryView = $view.ToString(); InstallDirectory = $installDir; InstalledDll = $installed; Sha256 = $hash;
        WordRunning = [bool](Get-Process -Name WINWORD -ErrorAction SilentlyContinue); Scope = 'CurrentUser' } | ConvertTo-Json
    return
}
Write-Output "Microsoft Word：$($environmentInfo.WordBits) 位"
Write-Output "Git：$($environmentInfo.GitVersion)"
Write-Output '.NET Framework：4.8 或更新版本'
Write-Output "WordGit：$version"
if ($CheckOnly) { Write-Output '环境和安装文件检查通过。'; return }
if (Get-Process -Name WINWORD -ErrorAction SilentlyContinue) { throw '请先保存并关闭所有 Word 窗口，再安装或更新 WordGit。' }
if (-not $SetupExecutable) { $SetupExecutable = Join-Path $PSScriptRoot 'WordGit.Setup.exe' }
Stop-WordGitKernel $installDir
$installed = Install-WordGitFiles $installDir $PSScriptRoot $identity $SetupExecutable
$base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
try {
    Register-WordGitPlugin $base $identity.FullName $version ([Uri]$installed).AbsoluteUri
    Register-WordGitProgram $base $installDir $version
    $kernel = Register-WordGitKernel $base $installDir
} finally { $base.Dispose() }
Start-WordGitKernel $kernel
Install-WordGitWpsAddin $installDir
# Remove only obsolete versioned plugin binaries after registration succeeds.
Get-ChildItem -LiteralPath $installDir -File | Where-Object {
    $_.Name -match '^WordGit\.\d+\.\d+\.\d+\.\d+\.[A-Fa-f0-9]{12}\.dll$' -and $_.FullName -ne $installed
} | ForEach-Object {
    $obsolete = $_.Name
    try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop } catch { Write-Output "保留仍在使用的旧组件：$obsolete" }
}
Write-Output '安装完成。重新打开 Word，即可在顶部看到 WordGit 选项卡。'
Write-Output '升级保留原有版本历史。卸载入口位于 Windows 设置的应用列表中。'
