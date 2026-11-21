param([switch]$Test, [switch]$WordTest, [switch]$UiTest)
$ErrorActionPreference = 'Stop'
$framework = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
if (-not (Test-Path -LiteralPath $framework)) { $framework = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319' }
$compiler = Join-Path $framework 'csc.exe'
if (-not (Test-Path -LiteralPath $compiler)) { throw '需要 Windows .NET Framework 4.8。' }
$bin = Join-Path $PSScriptRoot 'bin'
New-Item -ItemType Directory -Path $bin -Force | Out-Null
$sources = @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'src') -Filter '*.cs' | Select-Object -ExpandProperty FullName)
$references = @('System.Windows.Forms.dll','System.Drawing.dll','System.IO.Compression.dll','Microsoft.CSharp.dll')
$options = @('/nologo','/codepage:65001','/langversion:5','/optimize+') + @($references | ForEach-Object { "/reference:$_" })
& $compiler @options /target:library "/out:$bin\WordGit.dll" @sources
if ($LASTEXITCODE -ne 0) { throw '插件构建失败。' }
Write-Output "$bin\WordGit.dll"
if ($Test) {
    $testSource = Join-Path $PSScriptRoot 'tests\Smoke.cs'
    & $compiler @options /target:exe "/out:$bin\WordGit.Smoke.exe" @sources $testSource
    if ($LASTEXITCODE -ne 0) { throw '测试构建失败。' }
    & "$bin\WordGit.Smoke.exe"
    if ($LASTEXITCODE -ne 0) { throw '测试失败。' }
}
if ($WordTest) {
    $testSource = Join-Path $PSScriptRoot 'tests\WordSmoke.cs'
    & $compiler @options /debug:pdbonly /target:exe "/out:$bin\WordGit.WordSmoke.exe" @sources $testSource
    if ($LASTEXITCODE -ne 0) { throw 'Word 测试构建失败。' }
    & "$bin\WordGit.WordSmoke.exe"
    if ($LASTEXITCODE -ne 0) { throw 'Word 实机测试失败。' }
}
if ($UiTest) {
    $testSource = Join-Path $PSScriptRoot 'tests\UiSmoke.cs'
    & $compiler @options /target:exe "/out:$bin\WordGit.UiSmoke.exe" @sources $testSource
    if ($LASTEXITCODE -ne 0) { throw '界面测试构建失败。' }
    & "$bin\WordGit.UiSmoke.exe" $bin
    if ($LASTEXITCODE -ne 0) { throw '界面布局测试失败。' }
}
