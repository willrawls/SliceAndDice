param([switch]$Launch)
$ErrorActionPreference='Stop'
$vswhere=Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
$instances=@(& $vswhere -latest -products '*' -version '[18.0,)' -format json | ConvertFrom-Json)
if($instances.Count -eq 0){throw 'Visual Studio 2026 with a .NET 10 extension host is required.'}
if(Get-Process devenv -ErrorAction SilentlyContinue){throw 'Close Visual Studio before installing the extension.'}
$instance=$instances[0]
$vsix=Join-Path $PSScriptRoot 'src/Sandy.VisualStudio/bin/Release/net10.0-windows10.0.17763.0/Sandy.VisualStudio.vsix'
if(-not(Test-Path -LiteralPath $vsix)){throw 'Run Verify.ps1 to build the VSIX first.'}
$artifacts=Join-Path $PSScriptRoot 'artifacts'
[void][IO.Directory]::CreateDirectory($artifacts)
$log=Join-Path $artifacts 'vsix-install.log'
$msbuild=Join-Path $instance.installationPath 'MSBuild/Current/Bin/MSBuild.exe'
$project=Join-Path $PSScriptRoot 'src/Sandy.VisualStudio/Sandy.VisualStudio.csproj'
# An explicitly empty global suffix selects the normal profile. Omitting it defaults to Exp.
# The native VSIXInstaller currently rejects the Extensibility manifest's multiple architectures.
& $msbuild $project /restore /t:Rebuild /p:Configuration=Release /p:DeployExtension=true "/p:DeployTargetInstanceId=$($instance.instanceId)" /p:VSSDKTargetPlatformRegRootSuffix= /m:1 /v:minimal /nologo /fl "/flp:logfile=$log;verbosity=normal"
if($LASTEXITCODE -ne 0){throw "SDK extension installation returned $LASTEXITCODE. See $log"}
Write-Host "Installed Sandy into Visual Studio instance $($instance.instanceId). Log: $log"
if($Launch){Start-Process -FilePath (Join-Path $instance.installationPath 'Common7/IDE/devenv.exe') -WindowStyle Normal}
