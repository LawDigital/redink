#requires -version 5.1
# Produce a self-contained source snapshot without customer data or credentials.
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
if (-not (Test-Path -LiteralPath (Join-Path $root 'redink_server\web.py') -PathType Leaf)) {
    throw 'Run this script from an extracted Red Ink Usage Server source tree (redink_server\web.py is missing).'
}
$desktop = [Environment]::GetFolderPath([Environment+SpecialFolder]::DesktopDirectory)
if ([string]::IsNullOrWhiteSpace($desktop) -or -not (Test-Path -LiteralPath $desktop -PathType Container)) {
    throw 'Windows desktop folder could not be located.'
}

# Allowlist: all runnable source, assets, tests, build configuration and documentation.
# Never copy runtime data, the private server.toml, tokens, secrets, certs or logs.
$directories = @('redink_server','tests','scripts','packaging','build','config','examples','docs','validation')
$topLevel = @('.gitignore','README.md','RedInkUsageServer.spec','run_server.py',
              'build-windows.bat','build-windows.ps1','requirements.txt','requirements-build.txt')
$skipDirectoryNames = @('.git','.venv','venv','env','__pycache__','.pytest_cache','.mypy_cache',
                        '.ruff_cache','dist','data','config-files','backups','logs','node_modules',
                        'bin','obj','.vs','tmp','temp','_internal')
$skipFileNames = @('server.toml','secrets.json','secrets.toml','.env','server.lock','usage.sqlite3',
                   'desktop.ini','thumbs.db','license.pem','private-key.pem')
$blockedExtensions = @('.zip','.7z','.rar','.exe','.dll','.pfx','.p12','.p8','.pem','.key','.cer',
                       '.crt','.sqlite','.sqlite3','.db','.log','.sig','.pyc','.pyo','.bak','.tmp')
$files = New-Object System.Collections.Generic.List[System.IO.FileInfo]
foreach ($name in $topLevel) {
    $path = Join-Path $root $name
    if (Test-Path -LiteralPath $path -PathType Leaf) { $files.Add((Get-Item -LiteralPath $path)) }
}
foreach ($name in $directories) {
    $folder = Join-Path $root $name
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) { continue }
    foreach ($file in (Get-ChildItem -LiteralPath $folder -File -Recurse -Force -ErrorAction Stop)) {
        $relative = $file.FullName.Substring($root.Length).TrimStart('\','/')
        $parts = $relative -split '[\\/]'
        $skip = $false
        foreach ($part in $parts[0..($parts.Length - 2)]) {
            if ($skipDirectoryNames -contains $part.ToLowerInvariant()) { $skip = $true; break }
        }
        if ($skip -or ($skipFileNames -contains $file.Name.ToLowerInvariant()) -or
            ($blockedExtensions -contains $file.Extension.ToLowerInvariant()) -or
            $file.Name.EndsWith('~') -or $file.Attributes.HasFlag([System.IO.FileAttributes]::ReparsePoint)) { continue }
        $files.Add($file)
    }
}
if ($files.Count -lt 20) { throw "Source tree appears incomplete ($($files.Count) files). Archive not created." }
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outputName = 'RedInkUsageServer_Source_' + $stamp + '.zip'
$outputPath = Join-Path $desktop $outputName
$tempPath = Join-Path $desktop ('.' + $outputName + '.incomplete')
if (Test-Path -LiteralPath $tempPath) { throw "Temporary file already exists: $tempPath" }
$stream = $null
$archive = $null
try {
    $stream = New-Object System.IO.FileStream($tempPath,[System.IO.FileMode]::CreateNew,[System.IO.FileAccess]::ReadWrite)
    $archive = New-Object System.IO.Compression.ZipArchive($stream,[System.IO.Compression.ZipArchiveMode]::Create,$false)
    foreach ($file in ($files | Sort-Object FullName -Unique)) {
        $relative = $file.FullName.Substring($root.Length).TrimStart('\','/').Replace('\','/')
        $entry = 'Red Ink Usage Server/' + $relative
        [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
            $archive, $file.FullName, $entry,[System.IO.Compression.CompressionLevel]::Optimal) | Out-Null
    }
    $manifest = $archive.CreateEntry('Red Ink Usage Server/SOURCE_ARCHIVE_INFO.txt')
    $writer = New-Object System.IO.StreamWriter($manifest.Open(),[System.Text.Encoding]::UTF8)
    try {
        $writer.WriteLine('Red Ink Usage Server source archive')
        $writer.WriteLine('Created: ' + (Get-Date).ToString('u'))
        $writer.WriteLine('Files: ' + $files.Count)
        $writer.WriteLine('Excluded: data, secrets.json, server.toml, certs, keys, logs, SQLite, virtual environments, binaries and compiled distributions.')
        $writer.WriteLine('This archive is source code, NOT a ready-to-run installation.')
    } finally { $writer.Dispose() }
} finally {
    if ($null -ne $archive) { $archive.Dispose() }
    if ($null -ne $stream) { $stream.Dispose() }
}
try {
    # Reopen to verify the ZIP before placing it at its final name.
    $verify = [System.IO.Compression.ZipFile]::OpenRead($tempPath)
    try {
        if ($verify.Entries.Count -lt 21) { throw 'Archive validation failed: missing entries.' }
        if (-not ($verify.Entries | Where-Object { $_.FullName -eq 'Red Ink Usage Server/redink_server/web.py' })) {
            throw 'Archive validation failed: web.py was not included.'
        }
    } finally { $verify.Dispose() }
    Move-Item -LiteralPath $tempPath -Destination $outputPath -ErrorAction Stop
    Write-Host "Source archive created: $outputPath" -ForegroundColor Green
    Write-Host "Included files: $($files.Count). No runtime secrets or customer data."
} catch {
    Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
    throw
}
