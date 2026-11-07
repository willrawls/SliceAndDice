param(
    [string]$Configuration = 'Release',
    [ValidatePattern('^[A-Za-z][A-Za-z0-9]*$')]
    [string]$RootSuffix = 'SandyCSharp10',
    [string]$MetXStringsRoot = '',
    [switch]$Launch
)
$ErrorActionPreference = 'Stop'
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
if (-not (Test-Path -LiteralPath $vswhere)) { throw 'Visual Studio Installer vswhere.exe was not found.' }
$instances = & $vswhere -latest -products '*' -version '[18.0,)' -format json | ConvertFrom-Json
if ($LASTEXITCODE -ne 0 -or -not $instances) { throw 'A complete Visual Studio 2026 installation is required.' }
$instance = @($instances)[0]
$msbuild = Join-Path $instance.installationPath 'MSBuild\Current\Bin\MSBuild.exe'
if (-not (Test-Path -LiteralPath $msbuild)) { throw "MSBuild.exe was not found in $($instance.installationPath)." }
$running = Get-CimInstance Win32_Process -Filter "Name='devenv.exe'" | Where-Object { $_.CommandLine -match ('(?i)/rootsuffix\s+' + [regex]::Escape($RootSuffix) + '(\s|$)') }
if ($running) { throw "Close the $RootSuffix experimental IDE before redeploying. Other Visual Studio profiles may remain open." }
$project = Join-Path $PSScriptRoot 'src/Sandy.VisualStudio/Sandy.VisualStudio.csproj'
$arguments = @($project, '/restore', '/t:Rebuild', "/p:Configuration=$Configuration", '/p:DeployExtension=true',
    "/p:DeployTargetInstanceId=$($instance.instanceId)", "/p:VSSDKTargetPlatformRegRootSuffix=$RootSuffix", '/v:minimal', '/nologo')
if ($MetXStringsRoot) { $arguments += "/p:MetXStringsRoot=$MetXStringsRoot" }
& $msbuild @arguments
if ($LASTEXITCODE -ne 0) { throw "Experimental deployment failed (exit $LASTEXITCODE)." }
Write-Host "Deployed Sandy to $($instance.displayName), experimental profile $RootSuffix."
if ($Launch) {
    $devenv = Join-Path $instance.installationPath 'Common7/IDE/devenv.exe'
    Start-Process -FilePath $devenv -ArgumentList @('/RootSuffix', $RootSuffix) | Out-Null
}
