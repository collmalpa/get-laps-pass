Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ReleaseVersion = '2.0.0'

$requiredPs2ExeVersion = [version]'1.0.18'
$requiredPesterVersion = [version]'3.4.0'
$requiredAnalyzerVersion = [version]'1.25.0'
$minimumPassedTests = 50

function Resolve-BuildModule {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [version]$RequiredVersion
    )

    $matchingModules = @(
        Get-Module -ListAvailable -Name $Name |
            Where-Object { $_.Version -eq $RequiredVersion }
    )
    if ($matchingModules.Count -eq 0) {
        throw ("Required module {0} version {1} is unavailable." -f $Name, $RequiredVersion)
    }

    $programFilesPath = [Environment]::GetFolderPath(
        [Environment+SpecialFolder]::ProgramFiles)
    $programFilesModuleRoot = [IO.Path]::GetFullPath(
        (Join-Path -Path $programFilesPath -ChildPath 'WindowsPowerShell/Modules'))

    $orderedModules = @(
        $matchingModules |
            Sort-Object @{
                Expression = {
                    if ($_.Path.StartsWith(
                        $programFilesModuleRoot,
                        [StringComparison]::OrdinalIgnoreCase)) {
                        0
                    } else {
                        1
                    }
                }
            }, @{
                Expression = { $_.Path }
            }
    )

    return $orderedModules[0]
}

function Test-DistPathInvariant {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot,

        [Parameter(Mandatory = $true)]
        [string]$DistPath
    )

    $trimCharacters = [char[]]@(
        [IO.Path]::DirectorySeparatorChar,
        [IO.Path]::AltDirectorySeparatorChar)
    $normalizedRepositoryRoot = [IO.Path]::GetFullPath(
        $RepositoryRoot).TrimEnd($trimCharacters)
    $normalizedDistPath = [IO.Path]::GetFullPath(
        $DistPath).TrimEnd($trimCharacters)
    $distParent = [IO.Directory]::GetParent($normalizedDistPath)

    if ($null -eq $distParent) {
        return $false
    }

    return (
        [string]::Equals(
            $distParent.FullName.TrimEnd($trimCharacters),
            $normalizedRepositoryRoot,
            [StringComparison]::OrdinalIgnoreCase) -and
        [string]::Equals(
            [IO.Path]::GetFileName($normalizedDistPath),
            'dist',
            [StringComparison]::Ordinal)
    )
}

function Test-SafeBuildOutputDirectory {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot,

        [Parameter(Mandatory = $true)]
        [string]$DistPath
    )

    if (-not (Test-DistPathInvariant -RepositoryRoot $RepositoryRoot -DistPath $DistPath)) {
        return $false
    }
    if (-not (Test-Path -LiteralPath $DistPath -PathType Container)) {
        return $false
    }

    $distItem = Get-Item -LiteralPath $DistPath -Force -ErrorAction Stop
    if (-not $distItem.PSIsContainer) {
        return $false
    }

    return (
        ($distItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -eq 0
    )
}

function Test-KnownPs2ExePipelineControlArtifact {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$InputObject
    )

    if ($null -eq $InputObject) {
        return $false
    }
    if (
        $InputObject.GetType().FullName -cne
        'System.Management.Automation.StopUpstreamCommandsException'
    ) {
        return $false
    }

    $targetSite = $InputObject.TargetSite
    if ($null -eq $targetSite -or $null -eq $targetSite.DeclaringType) {
        return $false
    }

    return (
        $targetSite.DeclaringType.FullName -ceq
        'Microsoft.PowerShell.Commands.SelectObjectCommand' -and
        $targetSite.Name -ceq 'ProcessRecord' -and
        ([string]$InputObject.RequestingCommandProcessor) -ceq 'Select-Object'
    )
}

function Get-PeImageMetadata {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $imageBytes = $null
    try {
        $imageBytes = [IO.File]::ReadAllBytes($Path)
        if ($imageBytes.Length -lt 256) {
            throw 'The generated executable is too small to contain valid PE metadata.'
        }

        $peOffset = [BitConverter]::ToInt32($imageBytes, 0x3c)
        if ($peOffset -lt 0 -or ($peOffset + 94) -ge $imageBytes.Length) {
            throw 'The generated executable has an invalid PE header offset.'
        }

        if (
            $imageBytes[$peOffset] -ne 0x50 -or
            $imageBytes[$peOffset + 1] -ne 0x45 -or
            $imageBytes[$peOffset + 2] -ne 0 -or
            $imageBytes[$peOffset + 3] -ne 0
        ) {
            throw 'The generated executable does not contain a valid PE signature.'
        }

        $machine = [BitConverter]::ToUInt16($imageBytes, $peOffset + 4)
        $optionalHeaderOffset = $peOffset + 24
        $optionalHeaderMagic = [BitConverter]::ToUInt16(
            $imageBytes,
            $optionalHeaderOffset)
        $subsystem = [BitConverter]::ToUInt16(
            $imageBytes,
            $optionalHeaderOffset + 68)

        return [pscustomobject]@{
            Machine = $machine
            OptionalHeaderMagic = $optionalHeaderMagic
            Subsystem = $subsystem
        }
    } finally {
        if ($null -ne $imageBytes) {
            [Array]::Clear($imageBytes, 0, $imageBytes.Length)
        }
    }
}

$distPath = $null
$distPathVerified = $false

try {
    if ($PSVersionTable.PSEdition -ne 'Desktop') {
        throw 'The build requires Windows PowerShell Desktop edition.'
    }
    if (
        $PSVersionTable.PSVersion.Major -ne 5 -or
        $PSVersionTable.PSVersion.Minor -ne 1
    ) {
        throw 'The build requires Windows PowerShell 5.1.'
    }
    if (-not [Environment]::Is64BitProcess) {
        throw 'The x64 build requires a 64-bit Windows PowerShell process.'
    }
    if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        throw 'The repository root could not be resolved from the build script.'
    }

    $parsedReleaseVersion = $null
    if (
        -not [version]::TryParse($ReleaseVersion, [ref]$parsedReleaseVersion) -or
        $ReleaseVersion.Split('.').Count -ne 3
    ) {
        throw 'ReleaseVersion must contain exactly three numeric components.'
    }

    $repositoryRoot = [IO.Path]::GetFullPath($PSScriptRoot)
    $fileVersion = '{0}.0' -f $ReleaseVersion
    $releaseDirectoryName = 'Get-LAPS-pass-{0}' -f $ReleaseVersion
    $zipFileName = '{0}.zip' -f $releaseDirectoryName
    $zipHashFileName = '{0}.sha256' -f $zipFileName

    $requiredRelativePaths = @(
        'build.ps1',
        'Get_LAPS_pass.ps1',
        'GetLapsPass.Core.psm1',
        'tests/GetLapsPass.Core.Tests.ps1',
        'PSScriptAnalyzerSettings.psd1',
        'config.example.json',
        'README.md',
        'LICENSE'
    )
    $sourcePaths = @{}
    foreach ($relativePath in $requiredRelativePaths) {
        $sourcePath = [IO.Path]::GetFullPath(
            (Join-Path -Path $repositoryRoot -ChildPath $relativePath))
        if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
            throw ("Required input file is missing: {0}" -f $relativePath)
        }

        $sourcePaths[$relativePath] = $sourcePath
    }

    $ps2ExeModule = Resolve-BuildModule -Name 'ps2exe' -RequiredVersion $requiredPs2ExeVersion
    $pesterModule = Resolve-BuildModule -Name 'Pester' -RequiredVersion $requiredPesterVersion
    $analyzerModule = Resolve-BuildModule -Name 'PSScriptAnalyzer' -RequiredVersion $requiredAnalyzerVersion

    Import-Module -Name $ps2ExeModule.Path -Force -ErrorAction Stop
    Import-Module -Name $pesterModule.Path -Force -ErrorAction Stop
    Import-Module -Name $analyzerModule.Path -Force -ErrorAction Stop

    $invokePs2ExeCommand = Get-Command -Name 'Invoke-ps2exe' -Module 'ps2exe' -CommandType Function -ErrorAction Stop
    $invokePesterCommand = Get-Command -Name 'Invoke-Pester' -Module 'Pester' -CommandType Function -ErrorAction Stop
    $invokeAnalyzerCommand = Get-Command -Name 'Invoke-ScriptAnalyzer' -Module 'PSScriptAnalyzer' -CommandType Cmdlet -ErrorAction Stop

    $pesterResult = & $invokePesterCommand -Script $sourcePaths['tests/GetLapsPass.Core.Tests.ps1'] -PassThru
    $inconclusiveCount = @(
        $pesterResult.TestResult |
            Where-Object { $_.Result -eq 'Inconclusive' }
    ).Count

    if (
        $pesterResult.FailedCount -ne 0 -or
        $pesterResult.SkippedCount -ne 0 -or
        $pesterResult.PendingCount -ne 0 -or
        $inconclusiveCount -ne 0 -or
        $pesterResult.PassedCount -lt $minimumPassedTests -or
        $pesterResult.TotalCount -ne $pesterResult.PassedCount
    ) {
        $pesterFailure = (
            'Pester gate failed: Passed={0}, Failed={1}, Skipped={2}, ' +
            'Pending={3}, Inconclusive={4}, Total={5}.')
        throw ($pesterFailure -f
            $pesterResult.PassedCount,
            $pesterResult.FailedCount,
            $pesterResult.SkippedCount,
            $pesterResult.PendingCount,
            $inconclusiveCount,
            $pesterResult.TotalCount)
    }

    $analysisTargets = @(
        $sourcePaths['build.ps1'],
        $sourcePaths['Get_LAPS_pass.ps1'],
        $sourcePaths['GetLapsPass.Core.psm1'],
        $sourcePaths['tests/GetLapsPass.Core.Tests.ps1']
    )
    $analysisFindings = @(
        foreach ($analysisTarget in $analysisTargets) {
            & $invokeAnalyzerCommand -Path $analysisTarget -Settings $sourcePaths['PSScriptAnalyzerSettings.psd1'] -ErrorAction Stop
        }
    )
    if ($analysisFindings.Count -ne 0) {
        throw ('PSScriptAnalyzer gate failed with {0} finding(s).' -f $analysisFindings.Count)
    }

    $distPath = [IO.Path]::GetFullPath(
        (Join-Path -Path $repositoryRoot -ChildPath 'dist'))
    if (-not (Test-DistPathInvariant -RepositoryRoot $repositoryRoot -DistPath $distPath)) {
        throw 'The build output path failed the repository child-path safety check.'
    }
    $distPathVerified = $true

    if (Test-Path -LiteralPath $distPath) {
        if (-not (Test-SafeBuildOutputDirectory -RepositoryRoot $repositoryRoot -DistPath $distPath)) {
            throw 'The existing dist path is not a safe non-reparse directory and will not be deleted.'
        }

        Remove-Item -LiteralPath $distPath -Recurse -Force -ErrorAction Stop
    }

    $releaseDirectory = Join-Path -Path $distPath -ChildPath $releaseDirectoryName
    $zipPath = Join-Path -Path $distPath -ChildPath $zipFileName
    $zipHashPath = Join-Path -Path $distPath -ChildPath $zipHashFileName
    $null = New-Item -Path $releaseDirectory -ItemType Directory -Force -ErrorAction Stop

    $outputExePath = Join-Path -Path $releaseDirectory -ChildPath 'Get_LAPS_pass.exe'
    if (Test-Path -LiteralPath $outputExePath) {
        throw 'The intended output executable already exists before compilation.'
    }

    $compileErrorItems = @()
    $compileParameters = @{
        inputFile = $sourcePaths['Get_LAPS_pass.ps1']
        outputFile = $outputExePath
        noConsole = $true
        x64 = $true
        STA = $true
        supportOS = $true
        title = 'Get-LAPS-pass'
        description = 'Windows LAPS credential retrieval and RDP launcher'
        product = 'Get-LAPS-pass'
        version = $fileVersion
    }
    $null = & $invokePs2ExeCommand @compileParameters -ErrorAction Continue -ErrorVariable +compileErrorItems

    if ($compileErrorItems.Count -ne 4) {
        throw (
            'PS2EXE compilation captured an unexpected number of ' +
            ('error-stream objects: {0}.' -f $compileErrorItems.Count)
        )
    }
    foreach ($compileErrorItem in $compileErrorItems) {
        if (-not (Test-KnownPs2ExePipelineControlArtifact -InputObject $compileErrorItem)) {
            throw 'PS2EXE compilation captured an unexpected error-stream object.'
        }
    }
    if (-not (Test-Path -LiteralPath $outputExePath -PathType Leaf)) {
        throw 'PS2EXE did not create the expected executable.'
    }
    if ((Get-Item -LiteralPath $outputExePath -ErrorAction Stop).Length -eq 0) {
        throw 'PS2EXE created an empty executable.'
    }

    $sidecarRelativePaths = @(
        'GetLapsPass.Core.psm1',
        'config.example.json',
        'README.md',
        'LICENSE'
    )
    foreach ($sidecarRelativePath in $sidecarRelativePaths) {
        $sidecarFileName = [IO.Path]::GetFileName($sidecarRelativePath)
        $destinationPath = Join-Path -Path $releaseDirectory -ChildPath $sidecarFileName
        Copy-Item -LiteralPath $sourcePaths[$sidecarRelativePath] -Destination $destinationPath -ErrorAction Stop
    }

    $unexpectedConfigurations = @(
        Get-ChildItem -LiteralPath $distPath -Filter 'config.json' -File -Recurse -Force -ErrorAction Stop
    )
    if ($unexpectedConfigurations.Count -ne 0) {
        throw 'A local config.json was found in the build output.'
    }

    $versionInfo = [Diagnostics.FileVersionInfo]::GetVersionInfo($outputExePath)
    if ($versionInfo.FileVersion -cne $fileVersion) {
        throw 'The generated executable has an unexpected file version.'
    }
    if ($versionInfo.ProductVersion -cne $fileVersion) {
        throw 'The generated executable has an unexpected product version.'
    }
    if ($versionInfo.ProductName -cne 'Get-LAPS-pass') {
        throw 'The generated executable has an unexpected product name.'
    }
    # PS2EXE 1.0.18 maps -title to FileDescription and -description to Comments.
    if ($versionInfo.FileDescription -cne 'Get-LAPS-pass') {
        throw 'The generated executable has an unexpected file description.'
    }
    if (
        $versionInfo.Comments -cne
        'Windows LAPS credential retrieval and RDP launcher'
    ) {
        throw 'The generated executable has unexpected description metadata.'
    }

    $peMetadata = Get-PeImageMetadata -Path $outputExePath
    if ($peMetadata.Machine -ne 0x8664) {
        throw 'The generated executable is not an AMD64/x64 PE image.'
    }
    if ($peMetadata.OptionalHeaderMagic -ne 0x020b) {
        throw 'The generated executable does not use the expected PE32+ format.'
    }
    if ($peMetadata.Subsystem -ne 2) {
        throw 'The generated executable is not a Windows GUI application.'
    }

    $signature = Get-AuthenticodeSignature -LiteralPath $outputExePath -ErrorAction Stop
    $signatureStatus = [string]$signature.Status

    foreach ($sidecarRelativePath in $sidecarRelativePaths) {
        $sidecarFileName = [IO.Path]::GetFileName($sidecarRelativePath)
        $packagedSidecarPath = Join-Path -Path $releaseDirectory -ChildPath $sidecarFileName
        $sourceHash = Get-FileHash -LiteralPath $sourcePaths[$sidecarRelativePath] -Algorithm SHA256 -ErrorAction Stop
        $packagedHash = Get-FileHash -LiteralPath $packagedSidecarPath -Algorithm SHA256 -ErrorAction Stop
        if ($sourceHash.Hash -cne $packagedHash.Hash) {
            throw ('Packaged sidecar hash does not match its source: {0}' -f $sidecarFileName)
        }
    }

    $payloadFileNames = @(
        'Get_LAPS_pass.exe',
        'GetLapsPass.Core.psm1',
        'config.example.json',
        'README.md',
        'LICENSE'
    )
    $hashLines = @(
        foreach ($payloadFileName in ($payloadFileNames | Sort-Object)) {
            $payloadPath = Join-Path -Path $releaseDirectory -ChildPath $payloadFileName
            $payloadHash = Get-FileHash -LiteralPath $payloadPath -Algorithm SHA256 -ErrorAction Stop
            '{0}  {1}' -f $payloadHash.Hash, $payloadFileName
        }
    )
    $hashManifestPath = Join-Path -Path $releaseDirectory -ChildPath 'SHA256SUMS.txt'
    [IO.File]::WriteAllLines(
        $hashManifestPath,
        [string[]]$hashLines,
        [Text.Encoding]::ASCII)

    Compress-Archive -LiteralPath $releaseDirectory -DestinationPath $zipPath -CompressionLevel Optimal -Force -ErrorAction Stop

    $zipHash = Get-FileHash -LiteralPath $zipPath -Algorithm SHA256 -ErrorAction Stop
    $zipHashLine = '{0}  {1}{2}' -f $zipHash.Hash, $zipFileName, [Environment]::NewLine
    [IO.File]::WriteAllText(
        $zipHashPath,
        $zipHashLine,
        [Text.Encoding]::ASCII)

    Write-Output ''
    Write-Output 'Build succeeded.'
    Write-Output ('Release version: {0}' -f $ReleaseVersion)
    Write-Output ('PowerShell version: {0}' -f $PSVersionTable.PSVersion)
    Write-Output ('PS2EXE: {0} ({1})' -f $ps2ExeModule.Version, $ps2ExeModule.Path)
    Write-Output ('Pester: {0} ({1})' -f $pesterModule.Version, $pesterModule.Path)
    Write-Output ('PSScriptAnalyzer: {0} ({1})' -f $analyzerModule.Version, $analyzerModule.Path)
    Write-Output ('Tests passed: {0}' -f $pesterResult.PassedCount)
    Write-Output ('Analyzer findings: {0}' -f $analysisFindings.Count)
    Write-Output ('EXE signature status: {0}' -f $signatureStatus)
    Write-Output ('EXE path: {0}' -f $outputExePath)
    Write-Output ('Release directory: {0}' -f $releaseDirectory)
    Write-Output ('ZIP path: {0}' -f $zipPath)
    Write-Output ('ZIP SHA-256: {0}' -f $zipHash.Hash)

    exit 0
} catch {
    $buildFailureMessage = $_.Exception.Message
    $cleanupFailureMessage = $null

    if (
        $distPathVerified -and
        $null -ne $distPath -and
        (Test-Path -LiteralPath $distPath)
    ) {
        try {
            if (-not (Test-SafeBuildOutputDirectory -RepositoryRoot $repositoryRoot -DistPath $distPath)) {
                $cleanupFailureMessage = 'The existing dist path is not a safe non-reparse directory and was not deleted.'
            } else {
                Remove-Item -LiteralPath $distPath -Recurse -Force -ErrorAction Stop
            }
        } catch {
            $cleanupFailureMessage = $_.Exception.Message
        }
    }

    Write-Error -Message ('Build failed: {0}' -f $buildFailureMessage) -ErrorAction Continue
    if ($null -ne $cleanupFailureMessage) {
        Write-Warning ('Incomplete build output could not be removed: {0}' -f $cleanupFailureMessage)
    }

    exit 1
}
