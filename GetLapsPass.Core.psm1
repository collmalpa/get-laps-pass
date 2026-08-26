function Test-GetLapsPassConfiguration {
    param ([AllowNull()][psobject]$Configuration)

    if ($null -eq $Configuration) {
        return $false
    }

    $propertyNames = @($Configuration.PSObject.Properties.Name)
    return (
        $propertyNames -contains 'SearchTemplate' -and
        $propertyNames -contains 'UserForConnect' -and
        $Configuration.SearchTemplate -is [string] -and
        $Configuration.UserForConnect -is [string] -and
        -not [string]::IsNullOrWhiteSpace($Configuration.UserForConnect)
    )
}

function Get-LapsErrorMessage {
    param ([string]$Category)

    switch ($Category) {
        'ModuleUnavailable' {
            return "The Windows LAPS PowerShell module is unavailable."
        }
        'ComputerNotFound' {
            return "The computer was not found in Active Directory."
        }
        'AccessDenied' {
            return "You are not authorized to read or decrypt this LAPS password."
        }
        'PasswordUnavailable' {
            return "The LAPS password is unavailable or could not be decrypted."
        }
        default {
            return "An unexpected error occurred while retrieving the LAPS credential."
        }
    }
}

function New-LapsRetrievalException {
    param ([string]$Category)

    $exception = New-Object System.InvalidOperationException(
        (Get-LapsErrorMessage -Category $Category))
    $exception.Data['LapsErrorCategory'] = $Category
    return $exception
}

function Get-LapsExceptionCategory {
    param ([System.Management.Automation.ErrorRecord]$ErrorRecord)

    $exception = $ErrorRecord.Exception
    while ($null -ne $exception) {
        $exceptionTypeName = $exception.GetType().FullName
        if ($exceptionTypeName -eq 'System.Management.Automation.CommandNotFoundException') {
            return 'ModuleUnavailable'
        }
        if ($exceptionTypeName -eq 'Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException') {
            return 'ComputerNotFound'
        }
        if (
            $exception -is [System.UnauthorizedAccessException] -or
            $exception -is [System.Security.SecurityException] -or
            $exception.HResult -eq -2147024891
        ) {
            return 'AccessDenied'
        }

        $exception = $exception.InnerException
    }

    $errorCategory = [string]$ErrorRecord.CategoryInfo.Category
    switch ($errorCategory) {
        'ObjectNotFound' {
            return 'ComputerNotFound'
        }
        'PermissionDenied' {
            return 'AccessDenied'
        }
        'SecurityError' {
            return 'AccessDenied'
        }
    }

    if ($ErrorRecord.FullyQualifiedErrorId -eq 'CommandNotFoundException') {
        return 'ModuleUnavailable'
    }

    return 'Unexpected'
}

function Resolve-LapsAccountSelection {
    param (
        [psobject]$LapsResult,
        [string]$FallbackAccount
    )

    $lapsAccount = [string]$LapsResult.Account
    if ([string]::IsNullOrWhiteSpace($lapsAccount)) {
        $rdpAccount = $FallbackAccount.Trim()
        $accountSource = 'ConfigFallback'
    } else {
        $rdpAccount = $lapsAccount.Trim()
        $accountSource = 'LAPS'
    }

    return [pscustomobject]@{
        RdpAccount = $rdpAccount
        AccountSource = $accountSource
    }
}

function ConvertTo-LapsCredential {
    param (
        [string]$RequestedHostname,
        [AllowNull()][AllowEmptyCollection()][object[]]$LapsResults,
        [string]$FallbackAccount
    )

    $normalizedResults = @($LapsResults)
    if ($normalizedResults.Count -eq 0) {
        throw (New-LapsRetrievalException -Category 'ComputerNotFound')
    }
    if ($normalizedResults.Count -ne 1) {
        throw (New-LapsRetrievalException -Category 'Unexpected')
    }

    $lapsResult = $normalizedResults[0]
    $decryptionStatus = [string]$lapsResult.DecryptionStatus
    if ($decryptionStatus -eq 'Unauthorized') {
        throw (New-LapsRetrievalException -Category 'AccessDenied')
    }
    if ($decryptionStatus -notin @('Success', 'NotApplicable')) {
        throw (New-LapsRetrievalException -Category 'PasswordUnavailable')
    }
    if ([string]::IsNullOrEmpty([string]$lapsResult.Password)) {
        throw (New-LapsRetrievalException -Category 'PasswordUnavailable')
    }

    $accountSelection = Resolve-LapsAccountSelection `
        -LapsResult $lapsResult `
        -FallbackAccount $FallbackAccount

    return [pscustomobject]@{
        RequestedHostname = $RequestedHostname
        LapsResult = $lapsResult
        RdpAccount = $accountSelection.RdpAccount
        AccountSource = $accountSelection.AccountSource
        ExpirationTimestamp = $lapsResult.ExpirationTimestamp
    }
}

function Test-RetrievedCredentialMatchesHostname {
    param (
        [AllowNull()][psobject]$Credential,
        [string]$Hostname
    )

    if ($null -eq $Credential) {
        return $false
    }

    return $Credential.RequestedHostname -ceq $Hostname
}

function New-RdpFileContent {
    param (
        [string]$Hostname,
        [string]$Username,
        [bool]$RedirectDrive
    )

    $driveStoreRedirect = if ($RedirectDrive) { 'C:\;' } else { '' }
    $rdpContent = @"
screen mode id:i:2
use multimon:i:0
session bpp:i:32
winposstr:s:0,1,0,0,800,600
compression:i:1
keyboardhook:i:2
audiomode:i:0
redirectprinters:i:0
redirectclipboard:i:1
redirectsmartcards:i:0
drivestoredirect:s:$driveStoreRedirect
full address:s:$Hostname
username:s:$Username
prompt for credentials:i:0
authentication level:i:2
enablecredsspsupport:i:1
"@
    return $rdpContent
}

function Get-ClipboardTextFingerprint {
    param ([string]$Text)

    $textBytes = $null
    $sha256 = $null
    try {
        $textBytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
        $sha256 = [System.Security.Cryptography.SHA256]::Create()
        $fingerprint = $sha256.ComputeHash($textBytes)
        return ,$fingerprint
    } finally {
        if ($null -ne $textBytes) {
            [System.Array]::Clear($textBytes, 0, $textBytes.Length)
        }
        if ($null -ne $sha256) {
            $sha256.Dispose()
        }
    }
}

Export-ModuleMember -Function @(
    'Test-GetLapsPassConfiguration',
    'Get-LapsErrorMessage',
    'New-LapsRetrievalException',
    'Get-LapsExceptionCategory',
    'Resolve-LapsAccountSelection',
    'ConvertTo-LapsCredential',
    'Test-RetrievedCredentialMatchesHostname',
    'New-RdpFileContent',
    'Get-ClipboardTextFingerprint'
)
