# Cross-implementation check: the original C# history code (src/) and the Rust core must use the
# same repository location, produce identical snapshot trees for the same file, and read each
# other's commits byte for byte. Run from the repository root with PowerShell 7:
#   pwsh -NoProfile -File tests/rust-compat.ps1
$ErrorActionPreference = 'Stop'
# Both programs print UTF-8; read native output as UTF-8 regardless of the console code page.
[Console]::OutputEncoding = New-Object Text.UTF8Encoding($false)
$OutputEncoding = New-Object Text.UTF8Encoding($false)
$root = Split-Path -Parent $PSScriptRoot
$framework = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
$compiler = Join-Path $framework 'csc.exe'
$bin = Join-Path $root 'bin'
New-Item -ItemType Directory -Path $bin -Force | Out-Null
$sources = @(Get-ChildItem -LiteralPath (Join-Path $root 'src') -Filter '*.cs' | Select-Object -ExpandProperty FullName)
$references = @('System.Windows.Forms.dll','System.Drawing.dll','System.IO.Compression.dll','Microsoft.CSharp.dll','System.Net.Http.dll','System.Web.Extensions.dll')
$csharp = Join-Path $bin 'WordGit.RustCompat.exe'
& $compiler /nologo /codepage:65001 /langversion:5 /optimize+ @($references | ForEach-Object { "/reference:$_" }) /target:exe "/out:$csharp" @sources (Join-Path $PSScriptRoot 'RustCompat.cs')
if ($LASTEXITCODE -ne 0) { throw 'C# 兼容性程序构建失败。' }
cargo build -q -p wordgit
if ($LASTEXITCODE -ne 0) { throw 'Rust 构建失败。' }
$rust = (cargo metadata --format-version 1 --no-deps | ConvertFrom-Json).target_directory
$rust = Join-Path $rust 'debug\wordgit.exe'

$temp = Join-Path ([IO.Path]::GetTempPath()) ('wordgit-compat-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $temp | Out-Null
$env:WORDGIT_HOME = Join-Path $temp 'home'
$storage = Join-Path $env:WORDGIT_HOME 'repositories'
$document = Join-Path $temp '兼容 测试.docx'
$failures = New-Object System.Collections.Generic.List[string]
function Check([bool]$condition, [string]$label) {
    if ($condition) { Write-Output "PASS: $label" } else { Write-Output "FAIL: $label"; $failures.Add($label) }
}
function Fixture([string]$text, [byte[]]$image, [string]$path) {
    Add-Type -AssemblyName System.IO.Compression
    $stream = [IO.File]::Open($path, 'CreateNew')
    $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Create)
    foreach ($part in @(@('[Content_Types].xml', '<Types/>'), @('_rels/.rels', '<Relationships/>'),
            @('word/document.xml', "<w:document><w:body><w:p><w:t>$text</w:t></w:p></w:body></w:document>"))) {
        $entry = $zip.CreateEntry($part[0])
        $writer = New-Object IO.StreamWriter($entry.Open(), (New-Object Text.UTF8Encoding($false)))
        $writer.Write($part[1]); $writer.Dispose()
    }
    $entry = $zip.CreateEntry('word/media/image1.png', [System.IO.Compression.CompressionLevel]::NoCompression)
    $output = $entry.Open(); $output.Write($image, 0, $image.Length); $output.Dispose()
    $zip.Dispose(); $stream.Dispose()
}
function Rust([string[]]$arguments) {
    $output = & $rust @arguments
    if ($LASTEXITCODE -ne 0) { throw "wordgit $($arguments -join ' ') 失败" }
    return ($output -join "`n")
}
function CSharp([string[]]$arguments) {
    $output = & $csharp @arguments
    if ($LASTEXITCODE -ne 0) { throw "C# $($arguments -join ' ') 失败" }
    return ($output -join "`n").Trim()
}
function GitIn([string]$repository, [string[]]$arguments) {
    $output = & git "--git-dir=$repository" @arguments
    if ($LASTEXITCODE -ne 0) { throw "git $($arguments -join ' ') 失败" }
    return ($output -join "`n").Trim()
}

try {
    $image = New-Object byte[] (400 * 1024)
    (New-Object Random 417).NextBytes($image)
    Fixture '第一版：Rust 提交' $image $document

    $status = Rust @('status', $document) | ConvertFrom-Json
    $location = CSharp @('location', $storage, $document)
    Check ($status.repository -eq $location) "仓库位置一致（$([IO.Path]::GetFileName($location))）"

    $first = (Rust @('commit', $document, '-m', 'Rust 初稿', '--author', 'Rust 客户端') | ConvertFrom-Json).commit
    $head = CSharp @('head', $storage, $document)
    Check ($head -eq "main $first") 'C# 读取 Rust 提交的分支与 HEAD'
    $out1 = Join-Path $temp 'csharp-extract-1.docx'
    CSharp @('extract', $storage, $document, $first, $out1) | Out-Null
    Check ((Get-FileHash $out1).Hash -eq (Get-FileHash $document).Hash) 'C# 逐字节还原 Rust 提交的文档'
    $prepared = (CSharp @('prepare', $storage, $document, $document)).Split(' ')
    $rustTree = GitIn $location @('rev-parse', "$first^{tree}")
    Check ($prepared[0] -eq $rustTree) "同一文件的分块快照树一致（$($rustTree.Substring(0, 12))）"
    Check ($prepared[1] -eq (Get-FileHash $document -Algorithm SHA256).Hash.ToLowerInvariant()) '整文件 SHA-256 一致'

    Remove-Item -LiteralPath $document
    Fixture '第二版：C# 提交，图片不变' $image $document
    $second = CSharp @('commit', $storage, $document, 'C# 第二版')
    $log = Rust @('log', $document)
    Check (($log -split "`n").Count -eq 2 -and $log.Contains('C# 第二版') -and $log.Contains('C# 客户端')) 'Rust 读取 C# 提交的历史'
    $preview = (Rust @('call', 'version/preview', (ConvertTo-Json -Compress @{ document = $document; revision = $second }))) | ConvertFrom-Json
    Check ((Get-FileHash $preview.file).Hash -eq (Get-FileHash $document).Hash) 'Rust 逐字节还原 C# 提交的文档'
    $status = Rust @('status', $document) | ConvertFrom-Json
    Check ($status.workingState -eq 'clean' -and $status.head -eq $second) 'Rust 判断工作稿与 C# 提交一致'
    $objects = GitIn $location @('count-objects', '-v')
    Write-Output $objects
}
finally {
    Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item Env:\WORDGIT_HOME -ErrorAction SilentlyContinue
}
if ($failures.Count -gt 0) { throw ('兼容性检查失败：' + ($failures -join '、')) }
Write-Output '全部兼容性检查通过。'
