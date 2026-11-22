param([Parameter(Mandatory=$true)][string]$PackageRoot)
$ErrorActionPreference = 'Stop'
$PackageRoot = [IO.Path]::GetFullPath($PackageRoot)
. (Join-Path $PackageRoot 'install.ps1') -LibraryOnly

function Assert-InstallTest([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}
function Invoke-SetupProbe([string]$Mode) {
    $start = New-Object Diagnostics.ProcessStartInfo
    $start.FileName = Join-Path $PackageRoot 'WordGit.Setup.exe'
    $start.Arguments = $Mode
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $start.StandardOutputEncoding = [Text.Encoding]::UTF8
    $start.StandardErrorEncoding = [Text.Encoding]::UTF8
    $process = [Diagnostics.Process]::Start($start)
    try {
        $output = $process.StandardOutput.ReadToEndAsync()
        $errorText = $process.StandardError.ReadToEndAsync()
        $process.WaitForExit()
        Assert-InstallTest ($process.ExitCode -eq 0) ("Setup probe failed: " + $errorText.Result + $output.Result)
        $output.Result
    } finally { $process.Dispose() }
}
function Get-TestDataHashes([string]$Root) {
    $result = @{}
    if (Test-Path -LiteralPath $Root) {
        foreach ($file in Get-ChildItem -LiteralPath $Root -Recurse -File) {
            $result[$file.FullName.Substring($Root.Length)] = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
        }
    }
    $result
}
function Assert-SameHashes($Before, $After, [string]$Message) {
    Assert-InstallTest ($Before.Count -eq $After.Count) $Message
    foreach ($path in $Before.Keys) { Assert-InstallTest ($After.ContainsKey($path) -and $Before[$path] -eq $After[$path]) ($Message + ': ' + $path) }
}

$realRepositories = Join-Path $env:LOCALAPPDATA 'WordGit\repositories'
$realBefore = Get-TestDataHashes $realRepositories
$workspace = Join-Path ([IO.Path]::GetDirectoryName($PackageRoot)) ('it-' + [Guid]::NewGuid().ToString('N').Substring(0,12))
$wordgitRoot = Join-Path $workspace 'user\WordGit'
$app = Join-Path $wordgitRoot 'app'
$sentinels = @(
    (Join-Path $wordgitRoot 'repositories\legacy.git\objects\aa\old-object'),
    (Join-Path $wordgitRoot 'repositories\legacy.git\refs\heads\main'),
    (Join-Path $wordgitRoot 'repositories\legacy.git\refs\tags\旧版'),
    (Join-Path $wordgitRoot 'sessions\未完成的合并审阅.docx')
)
foreach ($file in $sentinels) {
    New-Item -ItemType Directory -Path ([IO.Path]::GetDirectoryName($file)) -Force | Out-Null
    [IO.File]::WriteAllBytes($file, [byte[]](0,1,10,13,255,82,5,90))
}
$historyBefore = Get-TestDataHashes (Join-Path $wordgitRoot 'repositories')
$sessionsBefore = Get-TestDataHashes (Join-Path $wordgitRoot 'sessions')
$identity = [Reflection.AssemblyName]::GetAssemblyName((Join-Path $PackageRoot 'bin\WordGit.dll'))
# install.ps1 registers fixed CLSIDs and ProgIDs; they must be the ones compiled into the plugin.
$assembly = [Reflection.Assembly]::LoadFrom((Join-Path $PackageRoot 'bin\WordGit.dll'))
foreach ($expected in @(@('WordGit.Connect', $WordGitConnectClsid, $WordGitConnectProgId), @('WordGit.HistoryPane', $WordGitPaneClsid, $WordGitPaneProgId))) {
    $type = $assembly.GetType($expected[0], $true)
    $progId = @($type.GetCustomAttributes([Runtime.InteropServices.ProgIdAttribute], $false))[0].Value
    Assert-InstallTest ((('{' + $type.GUID.ToString().ToUpperInvariant() + '}') -eq $expected[1]) -and ($progId -eq $expected[2])) ('install.ps1 registers a different COM identity for ' + $expected[0])
}
$registryParent = 'Software\WordGit\InstallerTests'
$registryName = $registryParent + '\' + [Guid]::NewGuid().ToString('N')
try {
    $verify = Invoke-SetupProbe '/verify'
    Assert-InstallTest ($verify -match 'PASS') 'Standalone setup payload verification failed'
    $installationPlan = Invoke-SetupProbe '/plan' | ConvertFrom-Json
    Assert-InstallTest ($installationPlan.Scope -eq 'CurrentUser' -and $installationPlan.Version -eq $identity.Version.ToString() -and $installationPlan.WordBits -in @(32,64)) 'Invalid setup plan'
    $installed = Install-WordGitFiles $app $PackageRoot $identity (Join-Path $PackageRoot 'WordGit.Setup.exe')
    $sourceHash = (Get-FileHash -LiteralPath (Join-Path $PackageRoot 'bin\WordGit.dll')).Hash
    Assert-InstallTest ((Get-FileHash -LiteralPath $installed).Hash -eq $sourceHash) 'Installed DLL mismatch'
    # Simulate an existing plugin and a second installation while keeping all user data.
    Copy-Item -LiteralPath $installed -Destination (Join-Path $app 'WordGit.0.8.0.0.0123456789AB.dll')
    $second = Install-WordGitFiles $app $PackageRoot $identity (Join-Path $PackageRoot 'WordGit.Setup.exe')
    Assert-InstallTest ($second -eq $installed) 'A repeated installation changed the target'
    Assert-SameHashes $historyBefore (Get-TestDataHashes (Join-Path $wordgitRoot 'repositories')) 'Upgrade changed history'
    Assert-SameHashes $sessionsBefore (Get-TestDataHashes (Join-Path $wordgitRoot 'sessions')) 'Upgrade changed review drafts'
    foreach ($view in (Get-WordGitRegistryViews)) {
        $hive = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
        $base = $null
        try {
            $base = $hive.CreateSubKey($registryName)
            Register-WordGitPlugin $base $identity.FullName $identity.Version.ToString() ([Uri]$installed).AbsoluteUri
            Register-WordGitProgram $base $app $identity.Version.ToString()
            $key = $base.OpenSubKey('Software\Classes\CLSID\{C5707F76-F46F-4A11-800B-18E7D98377D3}\InprocServer32')
            try {
                Assert-InstallTest ($key.GetValue('CodeBase') -eq ([Uri]$installed).AbsoluteUri -and $key.GetValue('ThreadingModel') -eq 'Both') 'Invalid COM registration'
            } finally { $key.Dispose() }
            $key = $base.OpenSubKey('Software\Microsoft\Office\Word\Addins\WordGit.Connect')
            try { Assert-InstallTest ($key.GetValue('LoadBehavior') -eq 3) 'Word add-in is not enabled' }
            finally { $key.Dispose() }
            $key = $base.OpenSubKey('Software\Microsoft\Windows\CurrentVersion\Uninstall\WordGit')
            try { Assert-InstallTest ($key.GetValue('UninstallString') -match 'WordGit.Setup.exe.* /uninstall') 'Missing installed-app uninstaller' }
            finally { $key.Dispose() }
            Remove-WordGitRegistration $base
            Assert-InstallTest ($null -eq $base.OpenSubKey('Software\Classes\WordGit.Connect')) 'COM registration remained after uninstall'
        } finally {
            if ($base) { $base.Dispose() }
            if ($registryName -notmatch '^Software\\WordGit\\InstallerTests\\[a-f0-9]{32}$') { throw 'Unsafe test registry path' }
            $hive.DeleteSubKeyTree($registryName, $false)
            # Remove the test container (and its parent) only when nothing else lives there.
            foreach ($container in @($registryParent, 'Software\WordGit')) {
                $key = $hive.OpenSubKey($container)
                $empty = $key -and $key.SubKeyCount -eq 0 -and $key.ValueCount -eq 0
                if ($key) { $key.Dispose() }
                if ($empty) { $hive.DeleteSubKey($container, $false) }
            }
            $hive.Dispose()
        }
    }
    Remove-WordGitFiles $app
    Assert-InstallTest (-not(Test-Path -LiteralPath $installed)) 'Uninstall left the installed plugin'
    Assert-InstallTest (-not(Test-Path -LiteralPath $app)) 'Uninstall left an empty app folder'
    Assert-SameHashes $historyBefore (Get-TestDataHashes (Join-Path $wordgitRoot 'repositories')) 'Uninstall changed history'
    Assert-SameHashes $sessionsBefore (Get-TestDataHashes (Join-Path $wordgitRoot 'sessions')) 'Uninstall changed review drafts'
    $blocked = $false
    try { Remove-WordGitFiles (Join-Path $wordgitRoot 'repositories') } catch { $blocked = $true }
    Assert-InstallTest $blocked 'Uninstaller accepted the repository directory'
    Assert-SameHashes $realBefore (Get-TestDataHashes $realRepositories) 'Installer probes changed real history'
    Write-Output 'PASS: standalone EXE, environment plan, install/upgrade/uninstall, x86/x64 registration, history preservation and real-repository hashes.'
} finally {
    $target = [IO.Path]::GetFullPath($workspace)
    $parent = [IO.Path]::GetFullPath([IO.Path]::GetDirectoryName($PackageRoot)) + [IO.Path]::DirectorySeparatorChar
    if ($target.StartsWith($parent,[StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($target) -match '^it-[a-f0-9]{12}$' -and (Test-Path -LiteralPath $target)) {
        Remove-Item -LiteralPath $target -Recurse -Force
    }
}
