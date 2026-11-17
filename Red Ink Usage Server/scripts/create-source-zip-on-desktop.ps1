#requires -version 5.1
# Produce a self-contained source snapshot without customer data or credentials.
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
$desktop = [Environment]::GetFolderPath([Environment+SpecialFolder]::DesktopDirectory)
if ([string]::IsNullOrWhiteSpace($desktop) -or -not (Test-Path -LiteralPath $desktop -PathType Container)) {
    throw 'Windows desktop folder could not be located.'
}

# Keep the build inputs explicitly and prune generated or private trees before recursing.
$directories = @('redink_server','tests','scripts','packaging','config','examples','docs','validation')
$required = @('README.md','INSTALL_PYTHON.md','RedInkUsageServer.spec','run_server.py',
              'build-windows.bat','build-windows.ps1','build-python.bat','requirements.txt',
              'initialize-server.bat','run-server.bat',
              'requirements-build.txt','build\red-ink.ico','build\version_info.txt',
              'redink_server\web.py','redink_server\__init__.py','config\server.example.toml',
              'scripts\build_support.py','scripts\release_payload.py','scripts\build-python-release.py',
              'scripts\collect-third-party-notices.py','create-source-zip-on-desktop.bat',
              'scripts\create-source-zip-on-desktop.ps1','tests\data\counter-vb-vectors-v2.json',
              'tests\data\schema-v1.sql')
$optional = @('.gitignore','Tasks.md','LICENSE','LICENSE.txt','LICENSE.md','COPYING','NOTICE','NOTICE.txt')
$skipDirectoryNames = @('.git','.venv','venv','env','__pycache__','cache','.cache','.pytest_cache',
                        '.mypy_cache','.ruff_cache','.tox','htmlcov','dist','data','config-files',
                        'backups','logs','log-archives','reports','tls','certificates','node_modules',
                        'bin','obj','.vs','tmp','temp','_internal')
$skipFileNames = @('server.toml','secrets.json','secrets.toml','server.lock','feature_taxonomy.json',
                   'desktop.ini','thumbs.db','build-info.json','build-dependencies.txt','SHA256SUMS.txt')
$blockedExtensions = @('.zip','.7z','.rar','.exe','.dll','.pfx','.p12','.p8','.pem','.key','.cer',
                       '.crt','.sqlite','.sqlite3','.db','.log',
                       '.sig','.pyc','.pyo','.bak','.tmp','.whl')

function Test-LinkedItem($item) {
    $item.Attributes.HasFlag([System.IO.FileAttributes]::ReparsePoint)
}

function Get-SourceFiles([System.IO.DirectoryInfo]$folder) {
    foreach ($item in (Get-ChildItem -LiteralPath $folder.FullName -Force)) {
        if (Test-LinkedItem $item) { continue }
        $relative = $item.FullName.Substring($root.Length).TrimStart('\','/')
        if ($item.PSIsContainer) {
            # These example/fixture trees contain versioned inputs.
            $isFixture = $relative -in @('tests\data','examples\config-files')
            if (-not $isFixture -and ($item.Name -in $skipDirectoryNames -or
                $item.Name -like '.venv-*' -or $item.Name -like 'venv-*' -or
                $item.Name -like 'pyinstaller-*')) { continue }
            Get-SourceFiles $item
        } elseif ($item.Name -notin $skipFileNames -and $item.Extension -notin $blockedExtensions -and
                  $item.Name -notmatch '(?i)(^\.env(?:\.|$)|^\.coverage(?:\.|$)|\.(?:log|db|sqlite3?)(?:[.-]|$)|~$)') {
            $item
        }
    }
}

$files = @(
    foreach ($name in ($required + $optional)) {
        $path = Join-Path $root $name
        if (-not (Test-Path -LiteralPath $path)) {
            if ($name -in $required) { throw "Required source file missing: $name" }
            continue
        }
        $item = Get-Item -LiteralPath $path -Force
        if ($item.PSIsContainer -or (Test-LinkedItem $item) -or
            ($item.DirectoryName -ne $root -and (Test-LinkedItem $item.Directory))) {
            throw "Expected a regular source file: $name"
        }
        $item
    }
    foreach ($name in $directories) {
        $path = Join-Path $root $name
        if (-not (Test-Path -LiteralPath $path -PathType Container)) { continue }
        $folder = Get-Item -LiteralPath $path -Force
        if (Test-LinkedItem $folder) { throw "Linked source directory: $name" }
        Get-SourceFiles $folder
    }
)
$files = @($files | Sort-Object FullName -Unique)
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
$outputName = 'RedInkUsageServer_Source_' + $stamp + '.zip'
$outputPath = Join-Path $desktop $outputName
$tempPath = Join-Path $desktop ('.' + [Guid]::NewGuid().ToString('N') + '.incomplete')
$prefix = 'Red Ink Usage Server/'
$infoName = $prefix + 'SOURCE_ARCHIVE_INFO.txt'
$versionFile = Join-Path $root 'redink_server\__init__.py'
$version = Select-String -LiteralPath $versionFile -Pattern '^__version__\s*=\s*["'']([^"'']+)["'']'
if (-not $version) { throw 'Source version missing in redink_server\__init__.py.' }
$versionNumber = $version.Matches[0].Groups[1].Value
$seen = @{}
try {
    $stream = New-Object System.IO.FileStream($tempPath,[System.IO.FileMode]::CreateNew,[System.IO.FileAccess]::ReadWrite)
    try {
        $archive = New-Object System.IO.Compression.ZipArchive($stream,[System.IO.Compression.ZipArchiveMode]::Create,$true)
        try {
            foreach ($file in $files) {
                $relative = $file.FullName.Substring($root.Length).TrimStart('\','/').Replace('\','/')
                $entry = $prefix + $relative
                $seen.Add($entry, $file.Length)
                [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                    $archive, $file.FullName, $entry,[System.IO.Compression.CompressionLevel]::Optimal) | Out-Null
            }
            $manifest = $archive.CreateEntry($infoName)
            $writer = New-Object System.IO.StreamWriter($manifest.Open(),[System.Text.Encoding]::UTF8)
            try {
                $writer.WriteLine('Red Ink Usage Server source archive')
                $writer.WriteLine('Version: ' + $versionNumber)
                $writer.WriteLine('Created UTC: ' + [DateTime]::UtcNow.ToString('u'))
                $writer.WriteLine('Source files: ' + $files.Count)
                $writer.WriteLine('Python 3.11+; see README.md for the Windows and Python builds.')
                $writer.WriteLine('Excluded: local deployment state, keys, certificates, logs, databases, virtual environments, caches and generated builds.')
            } finally { $writer.Dispose() }
        } finally { $archive.Dispose() }
    } finally { $stream.Dispose() }

    $verify = [System.IO.Compression.ZipFile]::OpenRead($tempPath)
    try {
        if ($verify.Entries.Count -ne ($files.Count + 1) -or $null -eq $verify.GetEntry($infoName)) {
            throw 'Archive validation failed: unexpected file inventory.'
        }
        foreach ($name in $seen.Keys) {
            $entry = $verify.GetEntry($name)
            if ($null -eq $entry -or $entry.Length -ne $seen[$name]) {
                throw "Archive validation failed: $name"
            }
        }
    } finally { $verify.Dispose() }
    [System.IO.File]::Move($tempPath, $outputPath)
    Write-Host "Source archive created: $outputPath" -ForegroundColor Green
    Write-Host "Included source files: $($files.Count)"
} finally {
    if (Test-Path -LiteralPath $tempPath) {
        Remove-Item -LiteralPath $tempPath -Force
    }
}
