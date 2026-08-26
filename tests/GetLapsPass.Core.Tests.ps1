$modulePath = Join-Path `
    -Path (Split-Path -Parent $PSScriptRoot) `
    -ChildPath 'GetLapsPass.Core.psm1'

Remove-Module -Name 'GetLapsPass.Core' -Force -ErrorAction SilentlyContinue
Import-Module -Name $modulePath -Force -ErrorAction Stop

$script:SyntheticPassword = 'SYNTHETIC-' + [guid]::NewGuid().ToString('N')
$script:SyntheticExpiration = [datetime]'2030-01-02T03:04:05'

function New-SyntheticLapsResult {
    param (
        [string]$DecryptionStatus = 'Success',
        [AllowNull()]$Password = $script:SyntheticPassword,
        [AllowNull()]$Account = 'LocalAdmin',
        [AllowNull()]$ExpirationTimestamp = $script:SyntheticExpiration
    )

    return [pscustomobject]@{
        Account = $Account
        Password = $Password
        ExpirationTimestamp = $ExpirationTimestamp
        DecryptionStatus = $DecryptionStatus
    }
}

function Get-ThrownLapsCategory {
    param ([scriptblock]$Action)

    try {
        $null = & $Action
        return $null
    } catch {
        return [string]$_.Exception.Data['LapsErrorCategory']
    }
}

function New-SyntheticErrorRecord {
    param (
        [System.Exception]$Exception,
        [System.Management.Automation.ErrorCategory]$Category,
        [string]$ErrorId = 'SyntheticError'
    )

    return New-Object System.Management.Automation.ErrorRecord -ArgumentList @(
        $Exception,
        $ErrorId,
        $Category,
        $null)
}

Describe 'Test-GetLapsPassConfiguration' {
    It 'rejects a null configuration' {
        Test-GetLapsPassConfiguration -Configuration $null | Should Be $false
    }

    It 'rejects a configuration without SearchTemplate' {
        $configuration = [pscustomobject]@{ UserForConnect = 'Administrator' }
        Test-GetLapsPassConfiguration -Configuration $configuration | Should Be $false
    }

    It 'rejects a configuration without UserForConnect' {
        $configuration = [pscustomobject]@{ SearchTemplate = '' }
        Test-GetLapsPassConfiguration -Configuration $configuration | Should Be $false
    }

    It 'rejects a non-string SearchTemplate' {
        $configuration = [pscustomobject]@{
            SearchTemplate = 123
            UserForConnect = 'Administrator'
        }
        Test-GetLapsPassConfiguration -Configuration $configuration | Should Be $false
    }

    It 'rejects a non-string UserForConnect' {
        $configuration = [pscustomobject]@{
            SearchTemplate = ''
            UserForConnect = 123
        }
        Test-GetLapsPassConfiguration -Configuration $configuration | Should Be $false
    }

    It 'rejects a blank UserForConnect' {
        $configuration = [pscustomobject]@{
            SearchTemplate = ''
            UserForConnect = '   '
        }
        Test-GetLapsPassConfiguration -Configuration $configuration | Should Be $false
    }

    It 'accepts an empty SearchTemplate' {
        $configuration = [pscustomobject]@{
            SearchTemplate = ''
            UserForConnect = 'Administrator'
        }
        Test-GetLapsPassConfiguration -Configuration $configuration | Should Be $true
    }

    It 'accepts a valid configuration' {
        $configuration = [pscustomobject]@{
            SearchTemplate = 'TEST-'
            UserForConnect = 'Administrator'
        }
        Test-GetLapsPassConfiguration -Configuration $configuration | Should Be $true
    }

    It 'allows extra properties' {
        $configuration = [pscustomobject]@{
            SearchTemplate = ''
            UserForConnect = 'Administrator'
            AdditionalSetting = 'synthetic'
        }
        Test-GetLapsPassConfiguration -Configuration $configuration | Should Be $true
    }
}

Describe 'Resolve-LapsAccountSelection' {
    It 'trims and selects a populated LAPS account' {
        $lapsResult = [pscustomobject]@{ Account = '  LocalAdmin  ' }
        $selection = Resolve-LapsAccountSelection `
            -LapsResult $lapsResult `
            -FallbackAccount 'FallbackAdmin'

        $selection.RdpAccount | Should Be 'LocalAdmin'
        $selection.AccountSource | Should Be 'LAPS'
    }

    It 'uses and trims the fallback for a whitespace LAPS account' {
        $lapsResult = [pscustomobject]@{ Account = '   ' }
        $selection = Resolve-LapsAccountSelection `
            -LapsResult $lapsResult `
            -FallbackAccount '  FallbackAdmin  '

        $selection.RdpAccount | Should Be 'FallbackAdmin'
        $selection.AccountSource | Should Be 'ConfigFallback'
    }
}

Describe 'ConvertTo-LapsCredential validation' {
    It 'classifies zero results as ComputerNotFound' {
        $category = Get-ThrownLapsCategory {
            ConvertTo-LapsCredential `
                -RequestedHostname 'TEST-HOST-001' `
                -LapsResults @() `
                -FallbackAccount 'FallbackAdmin'
        }
        $category | Should Be 'ComputerNotFound'
    }

    It 'classifies multiple results as Unexpected' {
        $results = @(
            (New-SyntheticLapsResult),
            (New-SyntheticLapsResult)
        )
        $category = Get-ThrownLapsCategory {
            ConvertTo-LapsCredential `
                -RequestedHostname 'TEST-HOST-001' `
                -LapsResults $results `
                -FallbackAccount 'FallbackAdmin'
        }
        $category | Should Be 'Unexpected'
    }

    It 'classifies Unauthorized as AccessDenied' {
        $result = New-SyntheticLapsResult -DecryptionStatus 'Unauthorized'
        $category = Get-ThrownLapsCategory {
            ConvertTo-LapsCredential `
                -RequestedHostname 'TEST-HOST-001' `
                -LapsResults @($result) `
                -FallbackAccount 'FallbackAdmin'
        }
        $category | Should Be 'AccessDenied'
    }

    It 'classifies an unsupported decryption status as PasswordUnavailable' {
        $result = New-SyntheticLapsResult -DecryptionStatus 'DecryptionFailed'
        $category = Get-ThrownLapsCategory {
            ConvertTo-LapsCredential `
                -RequestedHostname 'TEST-HOST-001' `
                -LapsResults @($result) `
                -FallbackAccount 'FallbackAdmin'
        }
        $category | Should Be 'PasswordUnavailable'
    }

    It 'classifies a null password as PasswordUnavailable' {
        $result = New-SyntheticLapsResult -Password $null
        $category = Get-ThrownLapsCategory {
            ConvertTo-LapsCredential `
                -RequestedHostname 'TEST-HOST-001' `
                -LapsResults @($result) `
                -FallbackAccount 'FallbackAdmin'
        }
        $category | Should Be 'PasswordUnavailable'
    }

    It 'classifies an empty password as PasswordUnavailable' {
        $result = New-SyntheticLapsResult -Password ''
        $category = Get-ThrownLapsCategory {
            ConvertTo-LapsCredential `
                -RequestedHostname 'TEST-HOST-001' `
                -LapsResults @($result) `
                -FallbackAccount 'FallbackAdmin'
        }
        $category | Should Be 'PasswordUnavailable'
    }

    It 'accepts Success' {
        $result = New-SyntheticLapsResult -DecryptionStatus 'Success'
        $credential = ConvertTo-LapsCredential `
            -RequestedHostname 'TEST-HOST-001' `
            -LapsResults @($result) `
            -FallbackAccount 'FallbackAdmin'

        $credential.AccountSource | Should Be 'LAPS'
    }

    It 'accepts NotApplicable' {
        $result = New-SyntheticLapsResult -DecryptionStatus 'NotApplicable'
        $credential = ConvertTo-LapsCredential `
            -RequestedHostname 'TEST-HOST-001' `
            -LapsResults @($result) `
            -FallbackAccount 'FallbackAdmin'

        $credential.AccountSource | Should Be 'LAPS'
    }
}

Describe 'ConvertTo-LapsCredential output' {
    It 'uses a trimmed LAPS account and records its source' {
        $result = New-SyntheticLapsResult -Account '  LocalAdmin  '
        $credential = ConvertTo-LapsCredential `
            -RequestedHostname 'TEST-HOST-001' `
            -LapsResults @($result) `
            -FallbackAccount 'FallbackAdmin'

        $credential.RdpAccount | Should Be 'LocalAdmin'
        $credential.AccountSource | Should Be 'LAPS'
    }

    It 'uses the trimmed fallback and records its source for a whitespace account' {
        $result = New-SyntheticLapsResult -Account '   '
        $credential = ConvertTo-LapsCredential `
            -RequestedHostname 'TEST-HOST-001' `
            -LapsResults @($result) `
            -FallbackAccount '  FallbackAdmin  '

        $credential.RdpAccount | Should Be 'FallbackAdmin'
        $credential.AccountSource | Should Be 'ConfigFallback'
    }

    It 'preserves RequestedHostname' {
        $result = New-SyntheticLapsResult
        $credential = ConvertTo-LapsCredential `
            -RequestedHostname 'TEST-HOST-001' `
            -LapsResults @($result) `
            -FallbackAccount 'FallbackAdmin'

        $credential.RequestedHostname | Should Be 'TEST-HOST-001'
    }

    It 'preserves ExpirationTimestamp' {
        $result = New-SyntheticLapsResult
        $credential = ConvertTo-LapsCredential `
            -RequestedHostname 'TEST-HOST-001' `
            -LapsResults @($result) `
            -FallbackAccount 'FallbackAdmin'

        $credential.ExpirationTimestamp | Should Be $script:SyntheticExpiration
    }

    It 'preserves the original LapsResult object reference' {
        $result = New-SyntheticLapsResult
        $credential = ConvertTo-LapsCredential `
            -RequestedHostname 'TEST-HOST-001' `
            -LapsResults @($result) `
            -FallbackAccount 'FallbackAdmin'

        [object]::ReferenceEquals($result, $credential.LapsResult) | Should Be $true
    }

    It 'returns exactly the expected five properties' {
        $result = New-SyntheticLapsResult
        $credential = ConvertTo-LapsCredential `
            -RequestedHostname 'TEST-HOST-001' `
            -LapsResults @($result) `
            -FallbackAccount 'FallbackAdmin'
        $propertyNames = @($credential.PSObject.Properties.Name)

        $propertyNames.Count | Should Be 5
        ($propertyNames -join ',') | Should Be `
            'RequestedHostname,LapsResult,RdpAccount,AccountSource,ExpirationTimestamp'
    }

    It 'does not create a top-level Password property' {
        $result = New-SyntheticLapsResult
        $credential = ConvertTo-LapsCredential `
            -RequestedHostname 'TEST-HOST-001' `
            -LapsResults @($result) `
            -FallbackAccount 'FallbackAdmin'

        @($credential.PSObject.Properties.Name) -contains 'Password' | Should Be $false
    }
}

Describe 'Get-LapsErrorMessage' {
    It 'returns the fixed ModuleUnavailable message' {
        Get-LapsErrorMessage -Category 'ModuleUnavailable' |
            Should Be 'The Windows LAPS PowerShell module is unavailable.'
    }

    It 'returns the fixed ComputerNotFound message' {
        Get-LapsErrorMessage -Category 'ComputerNotFound' |
            Should Be 'The computer was not found in Active Directory.'
    }

    It 'returns the fixed AccessDenied message' {
        Get-LapsErrorMessage -Category 'AccessDenied' |
            Should Be 'You are not authorized to read or decrypt this LAPS password.'
    }

    It 'returns the fixed PasswordUnavailable message' {
        Get-LapsErrorMessage -Category 'PasswordUnavailable' |
            Should Be 'The LAPS password is unavailable or could not be decrypted.'
    }

    It 'returns the fixed Unexpected message for unknown categories' {
        Get-LapsErrorMessage -Category 'SyntheticUnknown' |
            Should Be 'An unexpected error occurred while retrieving the LAPS credential.'
    }
}

Describe 'Get-LapsExceptionCategory' {
    It 'classifies CommandNotFoundException as ModuleUnavailable' {
        $exception = New-Object System.Management.Automation.CommandNotFoundException(
            'Synthetic command was unavailable.')
        $errorRecord = New-SyntheticErrorRecord `
            -Exception $exception `
            -Category ([System.Management.Automation.ErrorCategory]::NotSpecified)

        Get-LapsExceptionCategory -ErrorRecord $errorRecord | Should Be 'ModuleUnavailable'
    }

    It 'classifies UnauthorizedAccessException as AccessDenied' {
        $exception = New-Object System.UnauthorizedAccessException(
            'Synthetic access failure.')
        $errorRecord = New-SyntheticErrorRecord `
            -Exception $exception `
            -Category ([System.Management.Automation.ErrorCategory]::NotSpecified)

        Get-LapsExceptionCategory -ErrorRecord $errorRecord | Should Be 'AccessDenied'
    }

    It 'classifies PermissionDenied as AccessDenied' {
        $exception = New-Object System.InvalidOperationException('Synthetic failure.')
        $errorRecord = New-SyntheticErrorRecord `
            -Exception $exception `
            -Category ([System.Management.Automation.ErrorCategory]::PermissionDenied)

        Get-LapsExceptionCategory -ErrorRecord $errorRecord | Should Be 'AccessDenied'
    }

    It 'classifies SecurityError as AccessDenied' {
        $exception = New-Object System.InvalidOperationException('Synthetic failure.')
        $errorRecord = New-SyntheticErrorRecord `
            -Exception $exception `
            -Category ([System.Management.Automation.ErrorCategory]::SecurityError)

        Get-LapsExceptionCategory -ErrorRecord $errorRecord | Should Be 'AccessDenied'
    }

    It 'classifies ObjectNotFound as ComputerNotFound' {
        $exception = New-Object System.InvalidOperationException('Synthetic failure.')
        $errorRecord = New-SyntheticErrorRecord `
            -Exception $exception `
            -Category ([System.Management.Automation.ErrorCategory]::ObjectNotFound)

        Get-LapsExceptionCategory -ErrorRecord $errorRecord | Should Be 'ComputerNotFound'
    }

    It 'leaves AuthenticationError as Unexpected' {
        $exception = New-Object System.InvalidOperationException('Synthetic failure.')
        $errorRecord = New-SyntheticErrorRecord `
            -Exception $exception `
            -Category ([System.Management.Automation.ErrorCategory]::AuthenticationError)

        Get-LapsExceptionCategory -ErrorRecord $errorRecord | Should Be 'Unexpected'
    }

    It 'classifies an unknown error as Unexpected' {
        $exception = New-Object System.InvalidOperationException('Synthetic failure.')
        $errorRecord = New-SyntheticErrorRecord `
            -Exception $exception `
            -Category ([System.Management.Automation.ErrorCategory]::NotSpecified)

        Get-LapsExceptionCategory -ErrorRecord $errorRecord | Should Be 'Unexpected'
    }
}

Describe 'Test-RetrievedCredentialMatchesHostname' {
    It 'matches the exact hostname' {
        $credential = [pscustomobject]@{ RequestedHostname = 'TEST-HOST-001' }
        Test-RetrievedCredentialMatchesHostname `
            -Credential $credential `
            -Hostname 'TEST-HOST-001' | Should Be $true
    }

    It 'rejects a different hostname' {
        $credential = [pscustomobject]@{ RequestedHostname = 'TEST-HOST-001' }
        Test-RetrievedCredentialMatchesHostname `
            -Credential $credential `
            -Hostname 'TEST-HOST-002' | Should Be $false
    }

    It 'rejects a case-only hostname difference' {
        $credential = [pscustomobject]@{ RequestedHostname = 'TEST-HOST-001' }
        Test-RetrievedCredentialMatchesHostname `
            -Credential $credential `
            -Hostname 'test-host-001' | Should Be $false
    }

    It 'rejects a null credential' {
        Test-RetrievedCredentialMatchesHostname `
            -Credential $null `
            -Hostname 'TEST-HOST-001' | Should Be $false
    }
}

Describe 'New-RdpFileContent' {
    It 'preserves the exact RDP content with C drive redirection enabled' {
        $content = New-RdpFileContent `
            -Hostname 'TEST-HOST-001' `
            -Username 'TEST-HOST-001\LocalAdmin' `
            -RedirectDrive $true
        $expected = @"
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
drivestoredirect:s:C:\;
full address:s:TEST-HOST-001
username:s:TEST-HOST-001\LocalAdmin
prompt for credentials:i:0
authentication level:i:2
enablecredsspsupport:i:1
"@

        [string]::Equals($content, $expected, [System.StringComparison]::Ordinal) |
            Should Be $true
    }

    It 'preserves the exact RDP content with drive redirection disabled' {
        $content = New-RdpFileContent `
            -Hostname 'TEST-HOST-001' `
            -Username 'TEST-HOST-001\LocalAdmin' `
            -RedirectDrive $false
        $expected = @"
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
drivestoredirect:s:
full address:s:TEST-HOST-001
username:s:TEST-HOST-001\LocalAdmin
prompt for credentials:i:0
authentication level:i:2
enablecredsspsupport:i:1
"@

        [string]::Equals($content, $expected, [System.StringComparison]::Ordinal) |
            Should Be $true
    }

    It 'does not include the synthetic password sentinel' {
        $content = New-RdpFileContent `
            -Hostname 'TEST-HOST-001' `
            -Username 'TEST-HOST-001\LocalAdmin' `
            -RedirectDrive $true

        $content.Contains($script:SyntheticPassword) | Should Be $false
    }
}

Describe 'Get-ClipboardTextFingerprint' {
    It 'returns one 32-byte array object' {
        $fingerprint = Get-ClipboardTextFingerprint -Text 'synthetic text'

        $fingerprint.GetType().FullName | Should Be 'System.Byte[]'
        $fingerprint.Length | Should Be 32
    }

    It 'returns identical bytes for identical text' {
        $first = Get-ClipboardTextFingerprint -Text 'synthetic text'
        $second = Get-ClipboardTextFingerprint -Text 'synthetic text'

        [Convert]::ToBase64String($first) | Should Be ([Convert]::ToBase64String($second))
    }

    It 'returns different bytes for different text' {
        $first = Get-ClipboardTextFingerprint -Text 'synthetic text one'
        $second = Get-ClipboardTextFingerprint -Text 'synthetic text two'

        [Convert]::ToBase64String($first) |
            Should Not Be ([Convert]::ToBase64String($second))
    }

    It 'is deterministic for Unicode text' {
        $unicodeText = ([string][char]0x03A9) + ([string][char]0x4E2D)
        $first = Get-ClipboardTextFingerprint -Text $unicodeText
        $second = Get-ClipboardTextFingerprint -Text $unicodeText

        [Convert]::ToBase64String($first) | Should Be ([Convert]::ToBase64String($second))
    }

    It 'supports an empty string' {
        $fingerprint = Get-ClipboardTextFingerprint -Text ''

        $fingerprint.GetType().FullName | Should Be 'System.Byte[]'
        $fingerprint.Length | Should Be 32
    }
}
