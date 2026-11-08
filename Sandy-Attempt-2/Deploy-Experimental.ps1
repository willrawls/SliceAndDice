param([switch]$Launch)
$ErrorActionPreference = 'Stop'
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
$installation = & $vswhere -latest -products '*' -version '[18.0,)' -format json | ConvertFrom-Json
if (-not $installation) { throw 'Visual Studio 2026 with a .NET 10 extensibility host is required.' }
$instance = @($installation)[0]
$msbuild = Join-Path $instance.installationPath 'MSBuild/Current/Bin/MSBuild.exe'
$project = Join-Path $PSScriptRoot 'src/Sandy.VisualStudio/Sandy.VisualStudio.csproj'
& $msbuild $project /restore /t:Rebuild /p:Configuration=Release /p:DeployExtension=true "/p:DeployTargetInstanceId=$($instance.instanceId)" /p:VSSDKTargetPlatformRegRootSuffix=SandyAttempt2 /m:1 /v:minimal /nologo
if ($LASTEXITCODE -ne 0) { throw 'Experimental VSIX deployment failed.' }
if ($Launch) { Start-Process -FilePath (Join-Path $instance.installationPath 'Common7/IDE/devenv.exe') -ArgumentList @('/RootSuffix', 'SandyAttempt2') -WindowStyle Normal }
