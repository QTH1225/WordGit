[CmdletBinding()]
param([switch]$Test, [switch]$WordTest, [switch]$NoBuild)
$ErrorActionPreference = 'Stop'
if (-not $NoBuild) { & (Join-Path $PSScriptRoot 'build.ps1') -Test:$Test -UiTest:$Test -WordTest:$WordTest }
$sourceDll = Join-Path $PSScriptRoot 'bin\WordGit.dll'
$identity = [Reflection.AssemblyName]::GetAssemblyName($sourceDll)
$version = $identity.Version.ToString(3)
$bin = Join-Path $PSScriptRoot 'bin'
$work = Join-Path $bin ('package-' + [Guid]::NewGuid().ToString('N').Substring(0,12))
$package = Join-Path $work ("WordGit-" + $version)
$dist = Join-Path $PSScriptRoot 'dist'
New-Item -ItemType Directory -Path $package,$dist -Force | Out-Null
try {
    $allowlist = @('bin/WordGit.dll','install.ps1','Install.cmd','Uninstall.cmd','Check.cmd','README.md',
        'docs/USAGE.md','docs/DEVELOPMENT.md','docs/RELEASING.md')
    $files = @()
    $number = 0
    foreach ($name in $allowlist) {
        $from = Join-Path $PSScriptRoot $name
        $to = Join-Path $package $name
        if (-not(Test-Path -LiteralPath $from -PathType Leaf)) { throw "发布文件缺失：$name" }
        New-Item -ItemType Directory -Path ([IO.Path]::GetDirectoryName($to)) -Force | Out-Null
        Copy-Item -LiteralPath $from -Destination $to -Force
        if ([IO.Path]::GetExtension($to) -eq '.ps1') { [IO.File]::WriteAllText($to,[IO.File]::ReadAllText($to),[Text.UTF8Encoding]::new($true)) }
        $files += [PSCustomObject]@{ Path=$name; Sha256=(Get-FileHash -LiteralPath $to -Algorithm SHA256).Hash; Resource=('WordGit.Payload.' + $number.ToString('000')) }
        $number++
    }
    $manifest = [PSCustomObject]@{ Version=$identity.Version.ToString(); Files=$files }
    [IO.File]::WriteAllText((Join-Path $package 'manifest.json'),($manifest | ConvertTo-Json -Depth 5),[Text.UTF8Encoding]::new($true))
    $generated = Join-Path $work 'Version.cs'
    [IO.File]::WriteAllText($generated,('[assembly:System.Reflection.AssemblyVersion("' + $identity.Version.ToString() + '")][assembly:System.Reflection.AssemblyFileVersion("' + $identity.Version.ToString() + '")]'),[Text.UTF8Encoding]::new($true))
    $framework = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
    if (-not(Test-Path -LiteralPath $framework)) { $framework = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319' }
    $compiler = Join-Path $framework 'csc.exe'
    $options = @('/nologo','/codepage:65001','/langversion:5','/optimize+','/platform:anycpu',
        '/reference:System.Windows.Forms.dll','/reference:System.Drawing.dll','/reference:System.Web.Extensions.dll')
    $resources = @('/resource:' + (Join-Path $package 'manifest.json') + ',WordGit.Manifest')
    foreach ($file in $files) { $resources += '/resource:' + (Join-Path $package $file.Path) + ',' + $file.Resource }
    $setup = Join-Path $package 'WordGit.Setup.exe'
    & $compiler @options @resources /target:winexe ('/win32manifest:'+(Join-Path $PSScriptRoot 'packaging\setup.manifest')) ("/out:$setup") (Join-Path $PSScriptRoot 'packaging\Setup.cs') $generated
    if ($LASTEXITCODE -ne 0) { throw '安装程序构建失败。' }
    if ($Test) {
        $smoke = Join-Path $work 'SetupSmoke.exe'
        & $compiler @options @resources /target:exe /main:WordGitSetup.SetupSmoke ("/out:$smoke") (Join-Path $PSScriptRoot 'packaging\Setup.cs') (Join-Path $PSScriptRoot 'packaging\SetupSmoke.cs') $generated
        if ($LASTEXITCODE -ne 0) { throw '安装界面测试构建失败。' }
        & $smoke $bin
        if ($LASTEXITCODE -ne 0) { throw '安装界面测试失败。' }
        $windowsPowerShell = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
        & $windowsPowerShell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'tests\Installer.Tests.ps1') -PackageRoot $package
        if ($LASTEXITCODE -ne 0) { throw '安装器集成测试失败。' }
    }
    $exe = Join-Path $dist ("WordGit-" + $version + "-Setup.exe")
    $zip = Join-Path $dist ("WordGit-" + $version + "-windows.zip")
    Copy-Item -LiteralPath $setup -Destination $exe -Force
    Compress-Archive -LiteralPath $package -DestinationPath $zip -CompressionLevel Optimal -Force
    $sums = @($exe,$zip) | ForEach-Object { (Get-FileHash -LiteralPath $_ -Algorithm SHA256).Hash.ToLowerInvariant() + '  ' + [IO.Path]::GetFileName($_) }
    [IO.File]::WriteAllLines((Join-Path $dist 'SHA256SUMS.txt'),$sums,[Text.Encoding]::ASCII)
    Write-Output $exe
    Write-Output $zip
    Write-Output (Join-Path $dist 'SHA256SUMS.txt')
} finally {
    $target = [IO.Path]::GetFullPath($work)
    $parent = [IO.Path]::GetFullPath($bin) + [IO.Path]::DirectorySeparatorChar
    if ($target.StartsWith($parent,[StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($target) -match '^package-[a-f0-9]{12}$' -and (Test-Path -LiteralPath $target)) {
        Remove-Item -LiteralPath $target -Recurse -Force
    }
}
