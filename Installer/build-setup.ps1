[CmdletBinding()]
param(
    [string]$PublishDirectory,
    [string]$OutputDirectory,
    [string]$Version
)

$ErrorActionPreference = 'Stop'

$repositoryRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($Version)) {
    [xml]$project = Get-Content -Raw -LiteralPath (Join-Path $repositoryRoot 'SpotLight.csproj')
    $namespace = New-Object System.Xml.XmlNamespaceManager($project.NameTable)
    $namespace.AddNamespace('msb', 'http://schemas.microsoft.com/developer/msbuild/2003')
    $versionNode = $project.SelectSingleNode('//msb:PropertyGroup/msb:Version', $namespace)
    if ($null -eq $versionNode) {
        throw 'SpotLight.csproj does not contain a Version property.'
    }

    $Version = $versionNode.InnerText.Trim()
}

if ($Version -notmatch '^\d+\.\d+\.\d+\.\d+$') {
    throw "Installer version '$Version' must contain four numeric components."
}

if ([string]::IsNullOrWhiteSpace($PublishDirectory)) {
    $releasePublish = Join-Path $repositoryRoot 'bin\Release\app.publish'
    $debugPublish = Join-Path $repositoryRoot 'bin\Debug\app.publish'
    $PublishDirectory = if (Test-Path -LiteralPath $releasePublish) {
        $releasePublish
    }
    elseif (Test-Path -LiteralPath $debugPublish) {
        $debugPublish
    }
    else {
        throw 'Publish the VSTO project before building the single-file installer.'
    }
}

if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $repositoryRoot 'dist'
}

$publishRoot = (Resolve-Path -LiteralPath $PublishDirectory).Path.TrimEnd('\')
$requiredFiles = @(
    (Join-Path $publishRoot 'setup.exe'),
    (Join-Path $publishRoot 'SpotLight.vsto')
)
foreach ($requiredFile in $requiredFiles) {
    if (-not (Test-Path -LiteralPath $requiredFile -PathType Leaf)) {
        throw "Required publish file is missing: $requiredFile"
    }
}

[xml]$deploymentManifest = Get-Content -Raw -LiteralPath (Join-Path $publishRoot 'SpotLight.vsto')
$publishedVersion = $deploymentManifest.assembly.assemblyIdentity.version
if ($publishedVersion -ne $Version) {
    throw "Published VSTO version '$publishedVersion' does not match project version '$Version'. Publish the project again before building the installer."
}

$compilerCandidates = @(
    "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
    "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe"
)
$compiler = $compilerCandidates | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1
if ($null -eq $compiler) {
    throw '.NET Framework C# compiler was not found.'
}

New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$outputPath = Join-Path (Resolve-Path -LiteralPath $OutputDirectory).Path 'SpotLightSetup.exe'
$sourcePath = Join-Path $PSScriptRoot 'SpotLightSetup.cs'
$generatedDirectory = Join-Path $repositoryRoot 'obj\Installer'
New-Item -ItemType Directory -Path $generatedDirectory -Force | Out-Null
$versionSourcePath = Join-Path $generatedDirectory 'SpotLightSetup.Version.g.cs'
$versionSource = @"
using System.Reflection;

[assembly: AssemblyTitle("SpotLight Setup")]
[assembly: AssemblyProduct("SpotLight")]
[assembly: AssemblyVersion("$Version")]
[assembly: AssemblyFileVersion("$Version")]
[assembly: AssemblyInformationalVersion("$Version")]
"@
[IO.File]::WriteAllText($versionSourcePath, $versionSource, [Text.UTF8Encoding]::new($false))

$compilerArguments = [System.Collections.Generic.List[string]]::new()
$compilerArguments.Add('/nologo')
$compilerArguments.Add('/target:winexe')
$compilerArguments.Add('/platform:anycpu')
$compilerArguments.Add('/optimize+')
$compilerArguments.Add("/out:$outputPath")
$compilerArguments.Add('/reference:System.dll')
$compilerArguments.Add('/reference:System.Core.dll')
$compilerArguments.Add('/reference:System.Windows.Forms.dll')
$compilerArguments.Add($sourcePath)
$compilerArguments.Add($versionSourcePath)

$payloadFiles = Get-ChildItem -LiteralPath $publishRoot -Recurse -File | Sort-Object FullName
foreach ($payloadFile in $payloadFiles) {
    $relativePath = $payloadFile.FullName.Substring($publishRoot.Length + 1)
    $encodedPath = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($relativePath))
    $encodedPath = $encodedPath.TrimEnd('=').Replace('+', '-').Replace('/', '_')
    $resourceName = "SpotLight.Payload.$encodedPath"
    $compilerArguments.Add("/resource:$($payloadFile.FullName),$resourceName")
}

& $compiler $compilerArguments
if ($LASTEXITCODE -ne 0) {
    throw "Installer compilation failed with exit code $LASTEXITCODE."
}

& $outputPath --verify
if ($LASTEXITCODE -ne 0) {
    throw 'The generated installer failed its embedded-payload verification.'
}

$result = Get-Item -LiteralPath $outputPath
$hash = (Get-FileHash -LiteralPath $outputPath -Algorithm SHA256).Hash
[PSCustomObject]@{
    Installer = $result.FullName
    Version = $Version
    Size = $result.Length
    PayloadFiles = $payloadFiles.Count
    SHA256 = $hash
}
