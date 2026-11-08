param([string]$Configuration = 'Release')
$ErrorActionPreference = 'Stop'
Push-Location $PSScriptRoot
try {
    python tools/inventory.py
    if ($LASTEXITCODE -ne 0) { throw 'Source inventory failed.' }
    dotnet build Sandy.slnx -c $Configuration --nologo -m:1 -p:UseSharedCompilation=false
    if ($LASTEXITCODE -ne 0) { throw 'Build failed.' }
    dotnet test tests/Sandy.Tests -c $Configuration --no-build --no-restore --nologo
    if ($LASTEXITCODE -ne 0) { throw 'Regression tests failed.' }
    dotnet "src/Sandy.Desktop/bin/$Configuration/net10.0-windows/Sandy.Desktop.dll" --smoke-forms artifacts/form-previews
    if ($LASTEXITCODE -ne 0) { throw 'Form smoke test failed.' }
    $renders = Get-Content artifacts/form-previews/results.json -Raw | ConvertFrom-Json
    if (@($renders).Count -ne 21) { throw 'Expected all 21 forms to render.' }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $package = [IO.Compression.ZipFile]::OpenRead((Join-Path $PSScriptRoot "src/Sandy.VisualStudio/bin/$Configuration/net10.0-windows8.0/Sandy.VisualStudio.vsix"))
    try {
        foreach ($entry in @('Desktop/Sandy.Desktop.exe', 'Desktop/Sandy.Desktop.dll', 'Desktop/Sandy.Desktop.runtimeconfig.json', 'Desktop/Export-LegacyMdb.ps1', 'Libraries/SliceAndDice.xml')) {
            if (-not $package.GetEntry($entry)) { throw "VSIX is missing $entry" }
        }
    }
    finally { $package.Dispose() }
    Write-Host 'Build, regression tests, 21 form renders and VSIX payload checks passed. Live Visual Studio behavior is a separate check.'
}
finally { Pop-Location }
