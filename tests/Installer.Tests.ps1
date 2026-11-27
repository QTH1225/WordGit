param(
    [Parameter(Mandatory=$true)][string]$PackageRoot,
    [string]$PreviousPackageRoot,
    [string]$SetupSmoke,
    [string]$ScreenshotDirectory
)
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
        foreach ($file in Get-ChildItem -LiteralPath $Root -Recurse -File -Force) {
            $result[$file.FullName.Substring($Root.Length)] = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
        }
    }
    $result
}
function Assert-SameHashes($Before, $After, [string]$Message) {
    Assert-InstallTest ($Before.Count -eq $After.Count) $Message
    foreach ($path in $Before.Keys) { Assert-InstallTest ($After.ContainsKey($path) -and $Before[$path] -eq $After[$path]) ($Message + ': ' + $path) }
}

function New-TestDocument([string]$Path, [string]$Text) {
    $parts = @{
        '[Content_Types].xml' = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
        '_rels/.rels' = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
        'word/document.xml' = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>' + $Text + '</w:t></w:r></w:p></w:body></w:document>'
    }
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path }
    $zip = [IO.Compression.ZipFile]::Open($Path, [IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($name in $parts.Keys) {
            $writer = [IO.StreamWriter]::new($zip.CreateEntry($name).Open())
            try { $writer.Write($parts[$name]) } finally { $writer.Dispose() }
        }
    } finally { $zip.Dispose() }
}

function Start-TestService([string]$Kernel, [string]$Root) {
    $portProbe = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, 0)
    $portProbe.Start()
    try { $port = $portProbe.LocalEndpoint.Port } finally { $portProbe.Stop() }
    $info = Join-Path $Root 'service.json'
    Remove-Item -LiteralPath $info -ErrorAction SilentlyContinue
    $process = Start-Process -FilePath $Kernel -ArgumentList @('serve', '--port', $port) -WindowStyle Hidden -PassThru
    try {
        $timer = [Diagnostics.Stopwatch]::StartNew()
        do {
            if ($process.HasExited -or $timer.Elapsed.TotalSeconds -gt 20) { throw 'Isolated WordGit service did not start' }
            if (Test-Path -LiteralPath $info) {
                $record = Get-Content -LiteralPath $info -Encoding UTF8 -Raw | ConvertFrom-Json
                Assert-InstallTest ($record.port -eq $port) 'Isolated service returned the wrong port'
                return @{ Process = $process; Port = $port; Token = $record.token }
            }
            Start-Sleep -Milliseconds 50
        } while ($true)
    } catch {
        Stop-TestService @{ Process = $process }
        throw
    }
}

function Stop-TestService($Service) {
    if (-not $Service) { return }
    $process = $Service.Process
    try { if (-not $process.HasExited) { $process.Kill(); $process.WaitForExit() } }
    finally { $process.Dispose() }
}

function Invoke-TestService($Service, [string]$Endpoint, $Body) {
    $request = @{
        Uri = "http://127.0.0.1:$($Service.Port)/api/v1/$Endpoint"
        Headers = @{ 'X-WordGit-Token' = $Service.Token }
        ContentType = 'application/json; charset=utf-8'
        Body = [Text.Encoding]::UTF8.GetBytes(($Body | ConvertTo-Json -Compress))
        TimeoutSec = 20
    }
    $response = Invoke-WebRequest -UseBasicParsing -Method Post @request
    [Text.Encoding]::UTF8.GetString($response.RawContentStream.ToArray()) | ConvertFrom-Json
}

$realRepositories = Join-Path $env:LOCALAPPDATA 'WordGit\repositories'
$realBefore = Get-TestDataHashes $realRepositories
$workspace = Join-Path ([IO.Path]::GetDirectoryName($PackageRoot)) ('it-' + [Guid]::NewGuid().ToString('N').Substring(0,12))
$wordgitRoot = Join-Path $workspace 'user\WordGit'
$app = Join-Path $wordgitRoot 'app'
$dataRoot = Join-Path $workspace 'configured-storage'
$repositories = Join-Path $dataRoot 'repositories'
$sessions = Join-Path $wordgitRoot 'sessions'
$config = Join-Path $wordgitRoot 'data-directory.json'
$sentinels = @(
    (Join-Path $repositories 'legacy.git\objects\aa\old-object'),
    (Join-Path $repositories 'legacy.git\refs\heads\main'),
    (Join-Path $repositories 'legacy.git\refs\tags\旧版'),
    (Join-Path $sessions '未完成的合并审阅.docx')
)
foreach ($file in $sentinels) {
    New-Item -ItemType Directory -Path ([IO.Path]::GetDirectoryName($file)) -Force | Out-Null
    [IO.File]::WriteAllBytes($file, [byte[]](0,1,10,13,255,82,5,90))
}
[IO.File]::WriteAllText($config, (@{ directory = $dataRoot } | ConvertTo-Json -Compress))
$configHash = (Get-FileHash -LiteralPath $config).Hash
$sessionsBefore = Get-TestDataHashes $sessions
$identity = [Reflection.AssemblyName]::GetAssemblyName((Join-Path $PackageRoot 'bin\WordGit.dll'))
# install.ps1 registers fixed CLSIDs and ProgIDs; they must be the ones compiled into the plugin.
$assembly = [Reflection.Assembly]::LoadFrom((Join-Path $PackageRoot 'bin\WordGit.dll'))
foreach ($expected in @(@('WordGit.Connect', $WordGitConnectClsid, $WordGitConnectProgId), @('WordGit.HistoryPane', $WordGitPaneClsid, $WordGitPaneProgId))) {
    $type = $assembly.GetType($expected[0], $true)
    $progId = @($type.GetCustomAttributes([Runtime.InteropServices.ProgIdAttribute], $false))[0].Value
    Assert-InstallTest ((('{' + $type.GUID.ToString().ToUpperInvariant() + '}') -eq $expected[1]) -and ($progId -eq $expected[2])) ('install.ps1 registers a different COM identity for ' + $expected[0])
}
$registryName = 'WordGit-InstallerTest-' + [Guid]::NewGuid().ToString('N')
# An actual CurrentVersion\Run name can be retained by Windows during recursive
# cleanup even inside a fixture. Exercise the same value writes under a test key.
$WordGitStartupKey = 'Software\Microsoft\Windows\CurrentVersion\StartupFixture'
$originalHome = $env:WORDGIT_HOME
$originalAppData = $env:APPDATA
$service = $null
try {
    $verify = Invoke-SetupProbe '/verify'
    Assert-InstallTest ($verify -match 'PASS') 'Standalone setup payload verification failed'
    $installationPlan = Invoke-SetupProbe '/plan' | ConvertFrom-Json
    Assert-InstallTest ($installationPlan.Scope -eq 'CurrentUser' -and $installationPlan.Version -eq $identity.Version.ToString()) 'Invalid setup plan'
    Assert-InstallTest ($installationPlan.WordBits -in @(32,64) -and $installationPlan.RegistryView -eq ('Registry' + $installationPlan.WordBits) -and
        (Test-Path -LiteralPath $installationPlan.WordPath -PathType Leaf)) 'Invalid Word path or architecture in setup plan'
    $env:WORDGIT_HOME = $wordgitRoot
    $env:APPDATA = Join-Path $workspace 'roaming'
    Add-Type -AssemblyName System.IO.Compression.FileSystem, System.IO.Compression
    $kernelPath = Join-Path $app 'wordgit.exe'
    $previous = if ($PreviousPackageRoot) { [IO.Path]::GetFullPath($PreviousPackageRoot) } else { $PackageRoot }
    $previousIdentity = Test-WordGitPayload $previous
    $first = Install-WordGitFiles $app $previous $previousIdentity (Join-Path $previous 'WordGit.Setup.exe')
    Assert-InstallTest (Test-Path -LiteralPath $first -PathType Leaf) 'Initial DLL was not installed'
    Install-WordGitWpsAddin $app
    $service = Start-TestService $kernelPath $wordgitRoot
    $document = Join-Path $workspace '测试稿件.docx'
    New-TestDocument $document 'First submission'
    $olderHash = (Get-FileHash -LiteralPath $document).Hash
    $older = (Invoke-TestService $service 'commit' @{ document = $document; author = 'Installer test'; message = 'First submission' }).commit
    New-TestDocument $document 'Current submission'
    $currentHash = (Get-FileHash -LiteralPath $document).Hash
    $current = (Invoke-TestService $service 'commit' @{ document = $document; author = 'Installer test'; message = 'Current submission' }).commit
    Assert-InstallTest ($older -and $current -and $older -ne $current) 'Initial installation did not create two revisions'
    Stop-TestService $service
    $service = $null
    $historyBefore = Get-TestDataHashes $repositories
    if ($SetupSmoke) {
        $appBefore = Get-TestDataHashes $app
        $wpsBefore = Get-TestDataHashes $env:APPDATA
        & $SetupSmoke $ScreenshotDirectory
        Assert-InstallTest ($LASTEXITCODE -eq 0) 'Setup disclaimer UI test failed'
        Assert-SameHashes $appBefore (Get-TestDataHashes $app) 'Cancellation changed the installed add-in'
        Assert-SameHashes $wpsBefore (Get-TestDataHashes $env:APPDATA) 'Cancellation changed the WPS add-in'
        Assert-SameHashes $historyBefore (Get-TestDataHashes $repositories) 'Cancellation changed history'
        Assert-SameHashes $sessionsBefore (Get-TestDataHashes $sessions) 'Cancellation changed review drafts'
        Assert-InstallTest ((Get-FileHash -LiteralPath $config).Hash -eq $configHash) 'Cancellation changed the selected data directory'
        Assert-InstallTest ((Get-FileHash -LiteralPath $document).Hash -eq $currentHash) 'Cancellation changed the current draft'
    }

    $installed = Install-WordGitFiles $app $PackageRoot $identity (Join-Path $PackageRoot 'WordGit.Setup.exe')
    Install-WordGitWpsAddin $app
    $sourceHash = (Get-FileHash -LiteralPath (Join-Path $PackageRoot 'bin\WordGit.dll')).Hash
    Assert-InstallTest ((Get-FileHash -LiteralPath $installed).Hash -eq $sourceHash) 'Installed DLL mismatch'
    $wpsAddins = Join-Path $env:APPDATA 'kingsoft\wps\jsaddons'
    $wpsVersion = $identity.Version.ToString(3)
    $wpsConfig = Get-Content -LiteralPath (Join-Path $wpsAddins "WordGit_$wpsVersion\wordgit-config.json") -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-InstallTest ($wpsConfig.home -eq $wordgitRoot -and $wpsConfig.version -eq $wpsVersion) 'WPS did not upgrade to the isolated add-in'
    [xml]$wpsPublish = Get-Content -LiteralPath (Join-Path $wpsAddins 'publish.xml') -Raw -Encoding UTF8
    Assert-InstallTest ($wpsPublish.SelectNodes('/jsplugins/jsplugin[@name="WordGit"]').Count -eq 1) 'WPS upgrade duplicated the add-in'
    Assert-SameHashes $historyBefore (Get-TestDataHashes $repositories) 'Upgrade changed history'
    Assert-SameHashes $sessionsBefore (Get-TestDataHashes $sessions) 'Upgrade changed review drafts'
    Assert-InstallTest ((Get-FileHash -LiteralPath $config).Hash -eq $configHash) 'Upgrade changed the selected data directory'
    foreach ($view in (Get-WordGitRegistryViews)) {
        $hive = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
        $software = $base = $null
        try {
            $software = $hive.OpenSubKey('Software', $true)
            $base = $software.CreateSubKey($registryName)
            # 早期版本把同一个 COM 加载项也注册给 WPS；安装程序必须清掉它（WPS 现在用 JS 加载项）。
            $wpsListPath = 'Software\Kingsoft\Office\WPS\AddinsWL'
            Set-WordGitRegistryValue $base $wpsListPath 'Example.OtherAddIn' ''
            Set-WordGitRegistryValue $base $wpsListPath $WordGitConnectProgId ''
            $legacy = $base.CreateSubKey("Software\Kingsoft\Office\WPS\Addins\$WordGitConnectProgId")
            try { $legacy.SetValue('LoadBehavior', 3, [Microsoft.Win32.RegistryValueKind]::DWord) } finally { $legacy.Dispose() }
            Register-WordGitPlugin $base $previousIdentity.FullName $previousIdentity.Version.ToString() ([Uri]$first).AbsoluteUri
            Register-WordGitPlugin $base $identity.FullName $identity.Version.ToString() ([Uri]$installed).AbsoluteUri
            Register-WordGitProgram $base $app $identity.Version.ToString()
            Set-WordGitRegistryValue $base $WordGitStartupKey 'WordGit' ('"' + (Join-Path $app 'wordgit.exe') + '" serve')
            [void](Register-WordGitKernel $base $app)
            $key = $base.OpenSubKey($WordGitStartupKey)
            try {
                $launcher = Join-Path $app 'wordgit-launcher.exe'
                Assert-InstallTest ($key.GetValue('WordGit') -eq ('"' + $launcher + '"') -and (Test-Path -LiteralPath $launcher)) 'Login still launches the console executable'
            } finally { $key.Dispose() }
            $key = $base.OpenSubKey('Software\Classes\CLSID\{C5707F76-F46F-4A11-800B-18E7D98377D3}\InprocServer32')
            try {
                Assert-InstallTest ($key.GetValue('CodeBase') -eq ([Uri]$installed).AbsoluteUri -and $key.GetValue('ThreadingModel') -eq 'Both') 'Invalid COM registration'
            } finally { $key.Dispose() }
            $key = $base.OpenSubKey('Software\Microsoft\Office\Word\Addins\WordGit.Connect')
            try { Assert-InstallTest ($key.GetValue('LoadBehavior') -eq 3) 'Word add-in is not enabled' }
            finally { $key.Dispose() }
            Assert-InstallTest ($null -eq $base.OpenSubKey('Software\Kingsoft\Office\WPS\Addins\WordGit.Connect')) 'Installer kept the legacy WPS COM add-in'
            $key = $base.OpenSubKey($wpsListPath)
            try { Assert-InstallTest ($key -and -not($key.GetValueNames() -contains $WordGitConnectProgId)) 'Installer kept the legacy WPS add-in list entry' }
            finally { if ($key) { $key.Dispose() } }
            $key = $base.OpenSubKey('Software\Microsoft\Windows\CurrentVersion\Uninstall\WordGit')
            try { Assert-InstallTest ($key.GetValue('UninstallString') -match 'WordGit.Setup.exe.* /uninstall') 'Missing installed-app uninstaller' }
            finally { $key.Dispose() }
            Remove-WordGitRegistration $base
            Assert-InstallTest ($null -eq $base.OpenSubKey('Software\Classes\WordGit.Connect')) 'COM registration remained after uninstall'
            Assert-InstallTest ($null -eq $base.OpenSubKey('Software\Microsoft\Office\Word\AddinsData\WordGit.Connect')) 'Word load-state entry remained after uninstall'
            Assert-InstallTest ($null -eq $base.OpenSubKey('Software\Kingsoft\Office\WPS\Addins\WordGit.Connect')) 'WPS registration remained after uninstall'
            $key = $base.OpenSubKey($wpsListPath)
            try { Assert-InstallTest ($key -and @($key.GetValueNames()).Count -eq 1 -and
                $key.GetValueNames() -contains 'Example.OtherAddIn') 'Uninstall removed an unrelated WPS add-in' }
            finally { if ($key) { $key.Dispose() } }
        } finally {
            if ($base) { $base.Dispose() }
            try {
                if ($software) { $software.DeleteSubKeyTree($registryName, $false) }
            } finally {
                if ($software) { $software.Dispose() }
                $hive.Dispose()
            }
        }
    }
    $service = Start-TestService $kernelPath $wordgitRoot
    $revisionRequest = @{ document = $document; revision = $older; author = 'Installer test' }
    $preview = Invoke-TestService $service 'version/preview' $revisionRequest
    Assert-InstallTest ((Get-FileHash -LiteralPath $preview.file).Hash -eq $olderHash) 'Upgraded preview differed from the original DOCX'
    Assert-InstallTest ((Get-FileHash -LiteralPath $document).Hash -eq $currentHash) 'Preview modified the current draft'
    $restorePlan = (Invoke-TestService $service 'version/restore' $revisionRequest).plan
    Assert-InstallTest (-not [string]::IsNullOrEmpty($restorePlan)) 'Restore plan not created'
    $null = Invoke-TestService $service 'plan/apply' @{ plan = $restorePlan }
    Assert-InstallTest ((Get-FileHash -LiteralPath $document).Hash -eq $olderHash) 'Restored DOCX differed from the original'
    $preview = Invoke-TestService $service 'version/preview' @{ document = $document; revision = $current }
    Assert-InstallTest ((Get-FileHash -LiteralPath $preview.file).Hash -eq $currentHash) 'The newer revision was lost after restoring an older version'
    Stop-TestService $service
    $service = $null
    $uninstallHistoryBefore = Get-TestDataHashes $repositories
    $uninstallSessionsBefore = Get-TestDataHashes $sessions
    Remove-WordGitFiles $app
    Assert-InstallTest (-not(Test-Path -LiteralPath $installed)) 'Uninstall left the installed plugin'
    Assert-InstallTest (-not(Test-Path -LiteralPath $app)) 'Uninstall left an empty app folder'
    Assert-SameHashes $uninstallHistoryBefore (Get-TestDataHashes $repositories) 'Uninstall changed history'
    Assert-SameHashes $uninstallSessionsBefore (Get-TestDataHashes $sessions) 'Uninstall changed review drafts'
    $blocked = $false
    try { Remove-WordGitFiles $repositories } catch { $blocked = $true }
    Assert-InstallTest $blocked 'Uninstaller accepted the repository directory'
    Assert-SameHashes $realBefore (Get-TestDataHashes $realRepositories) 'Installer probes changed real history'
    Write-Output ("PASS: Windows install, upgrade ({0} -> {1}), x86/x64 registration, WPS add-in replacement, byte-identical DOCX preview and restore, history preservation." -f $previousIdentity.Version, $identity.Version)
} finally {
    Stop-TestService $service
    $env:WORDGIT_HOME = $originalHome
    $env:APPDATA = $originalAppData
    $target = [IO.Path]::GetFullPath($workspace)
    $parent = [IO.Path]::GetFullPath([IO.Path]::GetDirectoryName($PackageRoot)) + [IO.Path]::DirectorySeparatorChar
    if ($target.StartsWith($parent,[StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($target) -match '^it-[a-f0-9]{12}$' -and (Test-Path -LiteralPath $target)) {
        Remove-Item -LiteralPath $target -Recurse -Force
    }
}
