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
    dotnet "src/Sandy.Desktop/bin/$Configuration/net10.0-windows10.0.17763.0/Sandy.Desktop.dll" --smoke-ui artifacts/blazor-verification
    if ($LASTEXITCODE -ne 0) { throw 'Form smoke test failed.' }
    $renders = Get-Content artifacts/blazor-verification/results.json -Raw | ConvertFrom-Json
    if (-not $renders.Workbench -or @($renders.Forms).Count -ne 20 -or -not $renders.DraftSwitching -or -not $renders.SaveReload -or -not $renders.RepeatedPrompts -or -not $renders.ChoicePrompt -or -not $renders.LibraryMutation -or -not $renders.DatabaseGeneration -or -not $renders.NativeShellRegistration -or -not $renders.LegacyObfuscation -or -not $renders.CommandReference) { throw 'Blazor workspace/dialog checks failed.' }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $package = [IO.Compression.ZipFile]::OpenRead((Join-Path $PSScriptRoot "src/Sandy.VisualStudio/bin/$Configuration/net10.0-windows10.0.17763.0/Sandy.VisualStudio.vsix"))
    try {
        foreach ($entry in @('Desktop/Sandy.Desktop.exe', 'Desktop/Sandy.Desktop.dll', 'Desktop/Sandy.Desktop.runtimeconfig.json', 'Desktop/Export-LegacyMdb.ps1', 'Libraries/SliceAndDice.xml', 'Desktop/Sandy.Blazor.dll', 'Desktop/wwwroot/index.html', 'Desktop/wwwroot/app.css', 'Desktop/wwwroot/app.js', 'Desktop/runtimes/win-x64/native/WebView2Loader.dll')) {
            if (-not $package.GetEntry($entry)) { throw "VSIX is missing $entry" }
        }
    }
    finally { $package.Dispose() }
    Write-Host 'Build, regression tests, Blazor workspace plus 20 dialogs and VSIX payload checks passed. Live Visual Studio behavior is a separate check.'
}
finally { Pop-Location }
