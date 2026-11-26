param([switch]$Test, [switch]$WordTest, [switch]$UiTest, [switch]$KernelTest, [switch]$PaneTest, [switch]$NoRust)
$ErrorActionPreference = 'Stop'
$framework = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
if (-not (Test-Path -LiteralPath $framework)) { $framework = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319' }
$compiler = Join-Path $framework 'csc.exe'
if (-not (Test-Path -LiteralPath $compiler)) { throw '需要 Windows .NET Framework 4.8。' }
$bin = Join-Path $PSScriptRoot 'bin'
New-Item -ItemType Directory -Path $bin -Force | Out-Null
$sources = @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'src') -Filter '*.cs' | Select-Object -ExpandProperty FullName)
$support = Join-Path $PSScriptRoot 'tests\TestSupport.cs'
$references = @('System.Windows.Forms.dll','System.Drawing.dll','System.IO.Compression.dll','Microsoft.CSharp.dll','System.Net.Http.dll','System.Web.Extensions.dll')
$options = @('/nologo','/codepage:65001','/langversion:5','/optimize+') + @($references | ForEach-Object { "/reference:$_" })
& $compiler @options /target:library "/out:$bin\WordGit.dll" @sources
if ($LASTEXITCODE -ne 0) { throw '插件构建失败。' }
Write-Output "$bin\WordGit.dll"

# The history, branches and remote work live in the Rust core; the add-in only talks to it.
$kernel = Join-Path $bin 'wordgit.exe'
if (-not $NoRust) {
    $cargoPath = $null
    $cargo = Get-Command cargo -ErrorAction SilentlyContinue
    if ($cargo) { $cargoPath = $cargo.Source }
    if (-not $cargoPath) {
        $candidate = Join-Path $env:USERPROFILE '.cargo\bin\cargo.exe'
        if (Test-Path -LiteralPath $candidate) { $cargoPath = $candidate }
    }
    if (-not $cargoPath) { throw '需要 Rust 工具链（cargo）来构建 wordgit 内核。' }
    Push-Location $PSScriptRoot
    try {
        & $cargoPath build --locked --release --package wordgit
        if ($LASTEXITCODE -ne 0) { throw '内核构建失败。' }
    } finally { Pop-Location }
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'target\release\wordgit.exe') -Destination $kernel -Force
    Write-Output $kernel
}

# Each test program is built from all sources plus the shared test helpers. Every requested
# program runs even if an earlier one fails, and the failures are reported together.
$failures = New-Object System.Collections.Generic.List[string]
function Invoke-TestProgram([string]$Name, [string]$Label, [string[]]$Arguments = @(), [string[]]$Extra = @()) {
    $exe = Join-Path $bin ('WordGit.' + $Name + '.exe')
    & $compiler @options @Extra /target:exe "/out:$exe" @sources $support (Join-Path $PSScriptRoot ('tests\' + $Name + '.cs'))
    if ($LASTEXITCODE -ne 0) { $failures.Add($Label + '（构建）'); return }
    & $exe @Arguments
    if ($LASTEXITCODE -ne 0) { $failures.Add($Label) }
}
if ($Test) { Invoke-TestProgram 'UpdatesSmoke' '更新检查测试' }
if ($KernelTest) {
    if (-not (Test-Path -LiteralPath $kernel)) { throw '缺少 bin\wordgit.exe，请先不带 -NoRust 构建。' }
    $env:WORDGIT_SERVICE_EXE = $kernel
    try { Invoke-TestProgram 'KernelSmoke' '内核与 Word 联调测试' } finally { Remove-Item Env:\WORDGIT_SERVICE_EXE -ErrorAction SilentlyContinue }
}
if ($PaneTest) { Invoke-TestProgram 'PaneProbe' '侧栏首帧测试' }
if ($UiTest) { Invoke-TestProgram 'UiSmoke' '界面布局测试' -Arguments @($bin) }
if ($failures.Count -gt 0) { throw ('测试失败：' + ($failures -join '、')) }