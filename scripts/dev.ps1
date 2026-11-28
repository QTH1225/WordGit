<#
.SYNOPSIS
Run WordGit's Linux development tasks in the container defined by docker/dev/Dockerfile.

.DESCRIPTION
The Windows Word add-in, Setup.exe and wordgit.exe are still built natively with scripts\build.ps1
and scripts\package.ps1. This script covers the Linux side, with the same commands as scripts/dev.sh:

  .\scripts\dev.ps1 check       encoding, WPS add-in, clippy, tests and Linux package checks (what CI runs)
  .\scripts\dev.ps1 package     static x86_64 Linux release package into dist\
  .\scripts\dev.ps1 shell       interactive shell in the container (the default)
  .\scripts\dev.ps1 image       only build the image
  .\scripts\dev.ps1 <command>   any command, e.g. .\scripts\dev.ps1 cargo test -p wordgit-core

Optional download mirrors: WORDGIT_DOCKER_REGISTRY, WORDGIT_APT_MIRROR, WORDGIT_NODE_MIRROR
(see docs\DEVELOPMENT.md).
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$command = @($args)

function Read-RepositoryFile([string]$name) { [IO.File]::ReadAllText((Join-Path $root $name)) }

$toolchain = Read-RepositoryFile 'rust-toolchain.toml'
if ($toolchain -notmatch '(?m)^channel\s*=\s*"([^"]+)"') { throw 'rust-toolchain.toml does not name a channel.' }
$rust = $Matches[1]
$node = (Read-RepositoryFile 'toolchains/.node-version').Trim()
$python = (Read-RepositoryFile 'toolchains/.python-version').Trim()

# Same tag as scripts/dev.sh: the first 12 hex digits of SHA-256 over the version inputs without CR.
$inputs = @('docker/dev/Dockerfile', 'rust-toolchain.toml', 'toolchains/.node-version', 'toolchains/.python-version')
$text = (($inputs | ForEach-Object { Read-RepositoryFile $_ }) -join '').Replace("`r", '')
$sha = [Security.Cryptography.SHA256]::Create()
$digest = -join ($sha.ComputeHash([Text.UTF8Encoding]::new($false).GetBytes($text)) | ForEach-Object { $_.ToString('x2') })
$tag = $digest.Substring(0, 12)
$image = "wordgit-dev:$tag"

if (-not (docker images -q $image)) {
    $build = @('build', '-f', (Join-Path $root 'docker\dev\Dockerfile'), '-t', $image,
        '--build-arg', "RUST_VERSION=$rust", '--build-arg', "NODE_VERSION=$node", '--build-arg', "PYTHON_VERSION=$python")
    if ($env:WORDGIT_DOCKER_REGISTRY) { $build += @('--build-arg', "REGISTRY=$env:WORDGIT_DOCKER_REGISTRY") }
    if ($env:WORDGIT_APT_MIRROR) { $build += @('--build-arg', "APT_MIRROR=$env:WORDGIT_APT_MIRROR") }
    if ($env:WORDGIT_NODE_MIRROR) { $build += @('--build-arg', "NODE_MIRROR=$env:WORDGIT_NODE_MIRROR") }
    & docker @build $root
    if ($LASTEXITCODE -ne 0) { throw 'Building the development image failed.' }
    # Images from older version files are no longer used.
    foreach ($old in @(docker images wordgit-dev --format '{{.Tag}}')) {
        if ($old -and $old -ne $tag) { docker rmi "wordgit-dev:$old" | Out-Null }
    }
}

$first = if ($command.Count -gt 0) { $command[0] } else { 'shell' }
switch ($first) {
    'image' { exit 0 }
    { $_ -in 'check', 'package' } { $command = @('bash', 'scripts/linux-tasks.sh') + $command }
    'shell' { $command = @('bash') }
}

$run = @('run', '--rm', '--init', '-v', "${root}:/work", '-v', 'wordgit-dev-cache:/cache', '-w', '/work')
if (-not [Console]::IsInputRedirected -and -not [Console]::IsOutputRedirected) { $run += '-it' }
& docker @run $image @command
exit $LASTEXITCODE
