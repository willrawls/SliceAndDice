param([string]$Configuration = 'Release', [string]$MetXStringsRoot = '')
$ErrorActionPreference = 'Stop'
Push-Location $PSScriptRoot
try {
    $metXProperties = @()
    if ($MetXStringsRoot) { $metXProperties = @("-p:MetXStringsRoot=$MetXStringsRoot") }
    dotnet build SandyCSharp10.slnx -c $Configuration --nologo @metXProperties
    if ($LASTEXITCODE -ne 0) { throw 'Solution build failed.' }
    dotnet test tests/Sandy.Core.Tests/Sandy.Core.Tests.csproj -c $Configuration --no-restore --nologo @metXProperties
    if ($LASTEXITCODE -ne 0) { throw '.NET 10 engine tests failed.' }
    dotnet test tests/Sandy.Core.Tests/Sandy.Core.Tests.csproj -c $Configuration --no-restore --nologo -p:CoreTargetFramework=netstandard2.0 @metXProperties
    if ($LASTEXITCODE -ne 0) { throw '.NET Standard 2.0 engine tests failed.' }
    dotnet test tests/Sandy.Core.Tests/Sandy.Core.Tests.csproj -c $Configuration --no-restore --nologo -p:CoreTargetFramework=netstandard2.1 @metXProperties
    if ($LASTEXITCODE -ne 0) { throw '.NET Standard 2.1 engine tests failed.' }
    $vsix = Join-Path $PSScriptRoot "src/Sandy.VisualStudio/bin/$Configuration/net10.0-windows8.0/Sandy.VisualStudio.vsix"
    if (-not (Test-Path -LiteralPath $vsix)) { throw 'Expected VSIX package is missing.' }
    $resolvedMetXRoot = if ($MetXStringsRoot) { [IO.Path]::GetFullPath($MetXStringsRoot) } else { [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\xlg\MetX\MetX.Standard.Strings\bin\Debug')) }
    foreach ($framework in @('net10.0', 'netstandard2.0', 'netstandard2.1')) {
        $sourceHash = (Get-FileHash -LiteralPath (Join-Path $resolvedMetXRoot "$framework/MetX.Standard.Strings.dll") -Algorithm SHA256).Hash
        $copiedHash = (Get-FileHash -LiteralPath (Join-Path $PSScriptRoot "src/Sandy.Core/bin/$Configuration/$framework/MetX.Standard.Strings.dll") -Algorithm SHA256).Hash
        if ($sourceHash -ne $copiedHash) { throw "Incorrect MetX assembly copied for $framework." }
    }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::OpenRead($vsix)
    try {
        $entry = $archive.GetEntry('MetX.Standard.Strings.dll')
        if ($null -eq $entry) { throw 'MetX.Standard.Strings.dll is missing from the VSIX.' }
        $stream = $entry.Open()
        $sha = [Security.Cryptography.SHA256]::Create()
        try { $packageHash = [BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-', '') }
        finally { $sha.Dispose(); $stream.Dispose() }
        $runtimeHash = (Get-FileHash -LiteralPath (Join-Path $resolvedMetXRoot 'net10.0/MetX.Standard.Strings.dll') -Algorithm SHA256).Hash
        if ($packageHash -ne $runtimeHash) { throw 'The VSIX does not contain the matching .NET 10 MetX assembly.' }
    }
    finally { $archive.Dispose() }
    Write-Host 'Verified matching MetX references for all three targets and .NET 10 MetX assembly inside the VSIX.'
    Write-Host "Verified build and engine tests. VSIX: $vsix"
    Write-Host 'Live IDE checks are separate: see docs/VISUAL-STUDIO-VERIFICATION.md for recorded results and docs/COMPATIBILITY.md for rerun steps.'
}
finally { Pop-Location }
