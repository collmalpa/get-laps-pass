# Load required .NET assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Load the Windows Credential Manager API wrapper
$credentialManagerSource = @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;

public static class LapsRdpCredentialManager
{
    private const uint CredentialTypeGeneric = 1;
    private const uint CredentialPersistSession = 1;
    private const int ErrorNotFound = 1168;

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeFileTime
    {
        public uint LowDateTime;
        public uint HighDateTime;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct NativeCredential
    {
        public uint Flags;
        public uint Type;
        public IntPtr TargetName;
        public IntPtr Comment;
        public NativeFileTime LastWritten;
        public uint CredentialBlobSize;
        public IntPtr CredentialBlob;
        public uint Persist;
        public uint AttributeCount;
        public IntPtr Attributes;
        public IntPtr TargetAlias;
        public IntPtr UserName;
    }

    public sealed class GenericCredentialMetadata
    {
        public string Comment { get; private set; }

        internal GenericCredentialMetadata(string comment)
        {
            Comment = comment;
        }

        public override string ToString()
        {
            return "GenericCredentialMetadata";
        }
    }

    [DllImport("Advapi32.dll", EntryPoint = "CredReadW", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CredRead(
        string targetName,
        uint type,
        uint flags,
        out IntPtr credentialPointer);

    [DllImport("Advapi32.dll", EntryPoint = "CredWriteW", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CredWrite(
        ref NativeCredential credential,
        uint flags);

    [DllImport("Advapi32.dll", EntryPoint = "CredDeleteW", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CredDelete(
        string targetName,
        uint type,
        uint flags);

    [DllImport("Advapi32.dll", EntryPoint = "CredFree", SetLastError = false)]
    private static extern void CredFree(IntPtr buffer);

    public static GenericCredentialMetadata ReadGenericCredentialMetadata(string targetName)
    {
        ValidateTargetName(targetName);

        IntPtr credentialPointer;
        if (!CredRead(targetName, CredentialTypeGeneric, 0, out credentialPointer))
        {
            int errorCode = Marshal.GetLastWin32Error();
            if (errorCode == ErrorNotFound)
            {
                return null;
            }

            throw new Win32Exception(errorCode, "Unable to inspect the Windows credential.");
        }

        try
        {
            if (credentialPointer == IntPtr.Zero)
            {
                throw new InvalidOperationException("Windows returned invalid credential metadata.");
            }

            NativeCredential credential = (NativeCredential)Marshal.PtrToStructure(
                credentialPointer,
                typeof(NativeCredential));

            if (credential.Type != CredentialTypeGeneric)
            {
                throw new InvalidOperationException("Windows returned an unexpected credential type.");
            }

            string comment = credential.Comment == IntPtr.Zero
                ? null
                : Marshal.PtrToStringUni(credential.Comment);
            return new GenericCredentialMetadata(comment);
        }
        finally
        {
            CredFree(credentialPointer);
        }
    }

    public static void WriteTemporaryGenericCredential(
        string targetName,
        string userName,
        string password,
        string ownershipMarker)
    {
        ValidateTargetName(targetName);

        if (String.IsNullOrEmpty(userName))
        {
            throw new ArgumentException("A credential user name is required.", "userName");
        }

        if (password == null)
        {
            throw new ArgumentNullException("password");
        }

        if (String.IsNullOrEmpty(ownershipMarker))
        {
            throw new ArgumentException("An ownership marker is required.", "ownershipMarker");
        }

        byte[] passwordBytes = null;
        IntPtr targetPointer = IntPtr.Zero;
        IntPtr commentPointer = IntPtr.Zero;
        IntPtr userPointer = IntPtr.Zero;
        IntPtr passwordPointer = IntPtr.Zero;

        try
        {
            passwordBytes = Encoding.Unicode.GetBytes(password);
            targetPointer = Marshal.StringToCoTaskMemUni(targetName);
            commentPointer = Marshal.StringToCoTaskMemUni(ownershipMarker);
            userPointer = Marshal.StringToCoTaskMemUni(userName);

            if (passwordBytes.Length > 0)
            {
                passwordPointer = Marshal.AllocCoTaskMem(passwordBytes.Length);
                Marshal.Copy(passwordBytes, 0, passwordPointer, passwordBytes.Length);
            }

            NativeCredential credential = new NativeCredential();
            credential.Type = CredentialTypeGeneric;
            credential.TargetName = targetPointer;
            credential.Comment = commentPointer;
            credential.CredentialBlobSize = (uint)passwordBytes.Length;
            credential.CredentialBlob = passwordPointer;
            credential.Persist = CredentialPersistSession;
            credential.UserName = userPointer;

            if (!CredWrite(ref credential, 0))
            {
                int errorCode = Marshal.GetLastWin32Error();
                throw new Win32Exception(errorCode, "Unable to write the temporary Windows credential.");
            }
        }
        finally
        {
            if (passwordPointer != IntPtr.Zero)
            {
                int passwordLength = passwordBytes == null ? 0 : passwordBytes.Length;
                for (int index = 0; index < passwordLength; index++)
                {
                    Marshal.WriteByte(passwordPointer, index, 0);
                }

                Marshal.FreeCoTaskMem(passwordPointer);
            }

            if (passwordBytes != null)
            {
                Array.Clear(passwordBytes, 0, passwordBytes.Length);
            }

            FreeString(targetPointer);
            FreeString(commentPointer);
            FreeString(userPointer);
        }
    }

    public static void DeleteGenericCredential(string targetName)
    {
        ValidateTargetName(targetName);

        if (!CredDelete(targetName, CredentialTypeGeneric, 0))
        {
            int errorCode = Marshal.GetLastWin32Error();
            if (errorCode == ErrorNotFound)
            {
                return;
            }

            throw new Win32Exception(errorCode, "Unable to delete the temporary Windows credential.");
        }
    }

    private static void ValidateTargetName(string targetName)
    {
        if (String.IsNullOrEmpty(targetName))
        {
            throw new ArgumentException("A credential target is required.", "targetName");
        }
    }

    private static void FreeString(IntPtr pointer)
    {
        if (pointer != IntPtr.Zero)
        {
            Marshal.FreeCoTaskMem(pointer);
        }
    }
}
'@

try {
    if (-not ('LapsRdpCredentialManager' -as [type])) {
        Add-Type -TypeDefinition $credentialManagerSource -Language CSharp -ErrorAction Stop
    }
} catch {
    [void][System.Windows.Forms.MessageBox]::Show(
        "The Windows Credential Manager integration could not be initialized.",
        "Startup Error",
        "OK",
        "Error")
    return
}

# Load configuration relative to the script/application directory
$applicationDirectory = if ($PSScriptRoot) {
    $PSScriptRoot
} else {
    [System.AppDomain]::CurrentDomain.BaseDirectory
}

$coreModulePath = Join-Path -Path $applicationDirectory -ChildPath 'GetLapsPass.Core.psm1'
try {
    Import-Module -Name $coreModulePath -ErrorAction Stop
} catch {
    [void][System.Windows.Forms.MessageBox]::Show(
        "Unable to initialize the Get-LAPS-pass core module. Ensure GetLapsPass.Core.psm1 exists beside the application.",
        "Startup Error",
        "OK",
        "Error")
    return
}

$configPath = Join-Path -Path $applicationDirectory -ChildPath 'config.json'

try {
    if (-not (Test-Path -LiteralPath $configPath -PathType Leaf)) {
        throw (New-Object System.IO.FileNotFoundException)
    }

    $config = Get-Content -Raw -LiteralPath $configPath -ErrorAction Stop |
        ConvertFrom-Json -ErrorAction Stop

    if (-not (Test-GetLapsPassConfiguration -Configuration $config)) {
        throw (New-Object System.FormatException)
    }
} catch {
    [void][System.Windows.Forms.MessageBox]::Show(
        "Unable to load config.json. Ensure the file exists beside the application and contains valid SearchTemplate and UserForConnect string values.",
        "Configuration Error",
        "OK",
        "Error")
    return
}

# -------- FUNCTIONS --------

function Show-LapsRetrievalError {
    param ([System.Management.Automation.ErrorRecord]$errorRecord)

    $category = 'Unexpected'
    if (
        $null -ne $errorRecord.Exception -and
        $errorRecord.Exception.Data.Contains('LapsErrorCategory')
    ) {
        $category = [string]$errorRecord.Exception.Data['LapsErrorCategory']
    }

    [void][System.Windows.Forms.MessageBox]::Show(
        (Get-LapsErrorMessage -category $category),
        "LAPS Error",
        "OK",
        "Error")
}

function Get-LapsCredential {
    param ([string]$hostname)

    try {
        $lapsCommand = Get-Command `
            -Name 'Get-LapsADPassword' `
            -CommandType Cmdlet `
            -ErrorAction Stop
    } catch {
        throw (New-LapsRetrievalException -category 'ModuleUnavailable')
    }

    if ($null -eq $lapsCommand) {
        throw (New-LapsRetrievalException -category 'ModuleUnavailable')
    }
    $lapsCommand = $null

    try {
        $lapsResults = @(
            Get-LapsADPassword -Identity $hostname -AsPlainText -ErrorAction Stop
        )
    } catch {
        $category = Get-LapsExceptionCategory -errorRecord $_
        throw (New-LapsRetrievalException -category $category)
    }

    return (ConvertTo-LapsCredential `
        -RequestedHostname $hostname `
        -LapsResults $lapsResults `
        -FallbackAccount $config.UserForConnect)
}

function Invoke-OwnedClipboardCleanup {
    param ([bool]$retryOnFailure = $false)

    if ($null -ne $script:ClipboardCleanupTimer) {
        $script:ClipboardCleanupTimer.Stop()
    }

    if ($null -eq $script:OwnedClipboardFingerprint) {
        $script:ClipboardCleanupRetryCount = 0
        return $true
    }

    $currentClipboardText = $null
    $currentFingerprint = $null
    try {
        if ([System.Windows.Forms.Clipboard]::ContainsText(
            [System.Windows.Forms.TextDataFormat]::UnicodeText)) {
            $currentClipboardText = [System.Windows.Forms.Clipboard]::GetText(
                [System.Windows.Forms.TextDataFormat]::UnicodeText)
            $currentFingerprint = Get-ClipboardTextFingerprint -text $currentClipboardText

            $fingerprintsMatch =
                $script:OwnedClipboardFingerprint.Length -eq $currentFingerprint.Length
            if ($fingerprintsMatch) {
                for ($index = 0; $index -lt $currentFingerprint.Length; $index++) {
                    if (
                        $script:OwnedClipboardFingerprint[$index] -ne
                        $currentFingerprint[$index]
                    ) {
                        $fingerprintsMatch = $false
                        break
                    }
                }
            }

            if ($fingerprintsMatch) {
                [System.Windows.Forms.Clipboard]::Clear()
            }
        }
    } catch {
        if (
            $retryOnFailure -and
            $null -ne $script:ClipboardCleanupTimer -and
            $script:ClipboardCleanupRetryCount -lt $script:ClipboardCleanupMaxRetries
        ) {
            $script:ClipboardCleanupRetryCount++
            $script:ClipboardCleanupTimer.Interval = 1000
            $script:ClipboardCleanupTimer.Start()
        }

        return $false
    } finally {
        $currentClipboardText = $null
        if ($null -ne $currentFingerprint) {
            [System.Array]::Clear($currentFingerprint, 0, $currentFingerprint.Length)
            $currentFingerprint = $null
        }
    }

    [System.Array]::Clear(
        $script:OwnedClipboardFingerprint,
        0,
        $script:OwnedClipboardFingerprint.Length)
    $script:OwnedClipboardFingerprint = $null
    $script:ClipboardCleanupRetryCount = 0
    if ($null -ne $script:ClipboardCleanupTimer) {
        $script:ClipboardCleanupTimer.Interval = 30000
    }

    return $true
}

function Clear-RetrievedCredential {
    $script:RetrievedCredential = $null

    if ($null -ne $PassOutput -and -not $PassOutput.IsDisposed) {
        $PassOutput.Clear()
        $PassOutput.UseSystemPasswordChar = $true
    }
    if ($null -ne $RdpAccountOutput -and -not $RdpAccountOutput.IsDisposed) {
        $RdpAccountOutput.Clear()
    }
    if ($null -ne $ExpirationOutput -and -not $ExpirationOutput.IsDisposed) {
        $ExpirationOutput.Clear()
    }
    if ($null -ne $CopyButton -and -not $CopyButton.IsDisposed) {
        $CopyButton.Enabled = $false
    }
    if ($null -ne $ShowPasswordButton -and -not $ShowPasswordButton.IsDisposed) {
        $ShowPasswordButton.Enabled = $false
        $ShowPasswordButton.Text = "Show"
    }
}

function Set-RetrievedCredential {
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        'PSUseShouldProcessForStateChangingFunctions',
        '',
        Justification = 'Updates only in-memory application and UI state.')]
    param ([psobject]$RetrievedState)

    $script:RetrievedCredential = $RetrievedState
    $PassOutput.Text = [string]$RetrievedState.LapsResult.Password
    $PassOutput.UseSystemPasswordChar = $true
    $RdpAccountOutput.Text = '{0}\{1}' -f `
        $RetrievedState.RequestedHostname,
        $RetrievedState.RdpAccount

    if ($null -eq $RetrievedState.ExpirationTimestamp) {
        $ExpirationOutput.Text = "Not available"
    } else {
        $ExpirationOutput.Text = ([datetime]$RetrievedState.ExpirationTimestamp).ToString(
            'g',
            [System.Globalization.CultureInfo]::CurrentCulture)
    }

    $CopyButton.Enabled = $true
    $ShowPasswordButton.Enabled = $true
    $ShowPasswordButton.Text = "Show"
}

function Test-Hostname {
    param ([string]$hostname)
    try {
        Resolve-DnsName -Name $hostname -ErrorAction Stop | Out-Null
        return $true
    } catch {
        return $false
    }
}

function Connect-RDP {
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        'PSAvoidUsingPlainTextForPassword',
        '',
        Justification = 'Plaintext comes from Get-LapsADPassword -AsPlainText and is required by the existing Credential Manager integration.')]
    param (
        [string]$hostname,
        [string]$account,
        [string]$password,
        [bool]$redirectDrive
    )

    $user = "$hostname\$account"
    $credentialTarget = "TERMSRV/$hostname"
    $credentialMarker = $null
    $rdpFile = $null
    $cleanupWarnings = @()

    try {
        $existingCredentialMetadata = [LapsRdpCredentialManager]::ReadGenericCredentialMetadata(
            $credentialTarget)

        if ($null -ne $existingCredentialMetadata) {
            [void][System.Windows.Forms.MessageBox]::Show(
                "A saved RDP credential already exists for this host and was left untouched. Remove it manually from Windows Credential Manager before trying again.",
                "Saved Credential Exists",
                "OK",
                "Error")
            return
        }

        $rdpFileName = '{0}.rdp' -f [guid]::NewGuid().ToString('N')
        $rdpFile = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath $rdpFileName
        $rdpContent = New-RdpFileContent `
            -Hostname $hostname `
            -Username $user `
            -RedirectDrive $redirectDrive
        [System.IO.File]::WriteAllText(
            $rdpFile,
            $rdpContent,
            [System.Text.Encoding]::Unicode)
        $rdpFileArgument = '"{0}"' -f $rdpFile

        $credentialMarker = 'Get-LAPS-pass:{0}' -f [guid]::NewGuid().ToString('D')
        [LapsRdpCredentialManager]::WriteTemporaryGenericCredential(
            $credentialTarget,
            $user,
            $password,
            $credentialMarker)

        Start-Process -FilePath "mstsc.exe" -ArgumentList $rdpFileArgument -Wait -ErrorAction Stop
    } finally {
        if ($null -ne $credentialMarker) {
            $currentCredentialMetadata = $null
            try {
                $currentCredentialMetadata = [LapsRdpCredentialManager]::ReadGenericCredentialMetadata(
                    $credentialTarget)

                if ($null -eq $currentCredentialMetadata) {
                    # The temporary credential is already absent.
                } elseif ([string]::Equals(
                    $currentCredentialMetadata.Comment,
                    $credentialMarker,
                    [System.StringComparison]::Ordinal)) {
                    [LapsRdpCredentialManager]::DeleteGenericCredential($credentialTarget)
                } else {
                    $cleanupWarnings += "The RDP credential changed externally and was left untouched."
                }
            } catch {
                $cleanupWarnings += "The temporary RDP credential could not be verified or removed."
            } finally {
                $currentCredentialMetadata = $null
            }
        }

        $credentialMarker = $null

        if ($rdpFile -and [System.IO.File]::Exists($rdpFile)) {
            try {
                Remove-Item -LiteralPath $rdpFile -Force -ErrorAction Stop
            } catch {
                $cleanupWarnings += "The temporary RDP file could not be removed."
            }
        }

        if ($cleanupWarnings.Count -gt 0) {
            [void][System.Windows.Forms.MessageBox]::Show(
                ($cleanupWarnings -join [System.Environment]::NewLine),
                "Cleanup Warning",
                "OK",
                "Warning")
        }
    }
}

# -------- FORM DESIGN --------

$script:RetrievedCredential = $null
$script:OwnedClipboardFingerprint = $null
$script:ClipboardCleanupRetryCount = 0
$script:ClipboardCleanupMaxRetries = 3
$script:ClipboardCleanupTimer = New-Object System.Windows.Forms.Timer
$script:ClipboardCleanupTimer.Interval = 30000
$script:ClipboardCleanupTimer.Add_Tick({
    [void](Invoke-OwnedClipboardCleanup -retryOnFailure $true)
})

$Form = New-Object System.Windows.Forms.Form
$Form.ClientSize = '305,210'
$Form.Text = "Get LAPS pass"
$Form.FormBorderStyle = 'FixedSingle'
$Form.StartPosition = "CenterScreen"
$Form.KeyPreview = $true
$Form.TopMost = $false

# Enter/Escape key support
$Form.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") { $GetPasswordButton.PerformClick() }
    elseif ($_.KeyCode -eq "Escape") { $Form.Close() }
})

# Hostname label
$Label = [System.Windows.Forms.Label]@{
    Text = "Enter hostname:"
    Location = New-Object System.Drawing.Point(20, 13)
    Size = New-Object System.Drawing.Size(130, 20)
    Font = New-Object System.Drawing.Font('Tahoma',12)
}
$Form.Controls.Add($Label)

# Hostname textbox
$InputTextbox = [System.Windows.Forms.TextBox]@{
    Location = New-Object System.Drawing.Point(150,10)
    Size = New-Object System.Drawing.Size(135,25)
    Font = New-Object System.Drawing.Font('Tahoma',12)
    Text = $config.SearchTemplate
}
$InputTextbox.SelectionStart = $InputTextbox.Text.Length
$Form.Controls.Add($InputTextbox)

# Get Password button
$GetPasswordButton = [System.Windows.Forms.Button]@{
    Text = "Get Password"
    Location = New-Object System.Drawing.Point(20,45)
    Size = New-Object System.Drawing.Size(120,26)
    Font = New-Object System.Drawing.Font('Tahoma',11)
}
$GetPasswordButton.Add_Click({
    $hostname = $InputTextbox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($hostname)) {
        [void][System.Windows.Forms.MessageBox]::Show(
            "Enter a hostname before retrieving a password.",
            "Invalid Hostname",
            "OK",
            "Error")
        return
    }

    $retrievedCredential = $null
    try {
        Clear-RetrievedCredential
        $retrievedCredential = Get-LapsCredential -hostname $hostname
        Set-RetrievedCredential -RetrievedState $retrievedCredential
    } catch {
        Clear-RetrievedCredential
        Show-LapsRetrievalError -errorRecord $_
    } finally {
        $retrievedCredential = $null
    }
})
$Form.Controls.Add($GetPasswordButton)

# Copy button
$CopyButton = [System.Windows.Forms.Button]@{
    Text = "Copy"
    Location = New-Object System.Drawing.Point(150,45)
    Size = New-Object System.Drawing.Size(65,26)
    Font = New-Object System.Drawing.Font('Tahoma',10)
    Enabled = $false
}
$CopyButton.Add_Click({
    if (
        $null -eq $script:RetrievedCredential -or
        [string]::IsNullOrEmpty(
            [string]$script:RetrievedCredential.LapsResult.Password)
    ) {
        return
    }

    $copiedText = $null
    $newFingerprint = $null
    try {
        $copiedText = [string]$script:RetrievedCredential.LapsResult.Password
        $newFingerprint = Get-ClipboardTextFingerprint -text $copiedText
        [System.Windows.Forms.Clipboard]::SetText($copiedText)

        if ($null -ne $script:OwnedClipboardFingerprint) {
            [System.Array]::Clear(
                $script:OwnedClipboardFingerprint,
                0,
                $script:OwnedClipboardFingerprint.Length)
        }
        $script:OwnedClipboardFingerprint = $newFingerprint
        $newFingerprint = $null
        $script:ClipboardCleanupRetryCount = 0
        $script:ClipboardCleanupTimer.Stop()
        $script:ClipboardCleanupTimer.Interval = 30000
        $script:ClipboardCleanupTimer.Start()
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show(
            "The clipboard is currently unavailable.",
            "Clipboard Error",
            "OK",
            "Warning")
    } finally {
        $copiedText = $null
        if ($null -ne $newFingerprint) {
            [System.Array]::Clear($newFingerprint, 0, $newFingerprint.Length)
            $newFingerprint = $null
        }
    }
})
$Form.Controls.Add($CopyButton)

# Show/Hide Password button
$ShowPasswordButton = [System.Windows.Forms.Button]@{
    Text = "Show"
    Location = New-Object System.Drawing.Point(220,45)
    Size = New-Object System.Drawing.Size(65,26)
    Font = New-Object System.Drawing.Font('Tahoma',10)
    Enabled = $false
}
$ShowPasswordButton.Add_Click({
    if ($null -eq $script:RetrievedCredential) {
        return
    }

    $PassOutput.UseSystemPasswordChar = -not $PassOutput.UseSystemPasswordChar
    $ShowPasswordButton.Text = if ($PassOutput.UseSystemPasswordChar) {
        "Show"
    } else {
        "Hide"
    }
})
$Form.Controls.Add($ShowPasswordButton)

# Password label
$PassLabel = [System.Windows.Forms.Label]@{
    Text = "Admin Password:"
    Location = New-Object System.Drawing.Point(20,85)
    Size = New-Object System.Drawing.Size(130, 25)
    Font = New-Object System.Drawing.Font('Tahoma',11)
}
$Form.Controls.Add($PassLabel)

# Password output textbox
$PassOutput = [System.Windows.Forms.TextBox]@{
    ReadOnly = $true
    UseSystemPasswordChar = $true
    Location = New-Object System.Drawing.Point(150,80)
    Size = New-Object System.Drawing.Size(135,25)
    Font = New-Object System.Drawing.Font('Tahoma',11)
}
$Form.Controls.Add($PassOutput)

# Effective RDP account
$RdpAccountLabel = [System.Windows.Forms.Label]@{
    Text = "RDP account:"
    Location = New-Object System.Drawing.Point(20,115)
    Size = New-Object System.Drawing.Size(130,25)
    Font = New-Object System.Drawing.Font('Tahoma',11)
}
$Form.Controls.Add($RdpAccountLabel)

$RdpAccountOutput = [System.Windows.Forms.TextBox]@{
    ReadOnly = $true
    Location = New-Object System.Drawing.Point(150,110)
    Size = New-Object System.Drawing.Size(135,25)
    Font = New-Object System.Drawing.Font('Tahoma',10)
}
$Form.Controls.Add($RdpAccountOutput)

# LAPS expiration timestamp
$ExpirationLabel = [System.Windows.Forms.Label]@{
    Text = "Expiration:"
    Location = New-Object System.Drawing.Point(20,145)
    Size = New-Object System.Drawing.Size(130,25)
    Font = New-Object System.Drawing.Font('Tahoma',11)
}
$Form.Controls.Add($ExpirationLabel)

$ExpirationOutput = [System.Windows.Forms.TextBox]@{
    ReadOnly = $true
    Location = New-Object System.Drawing.Point(150,140)
    Size = New-Object System.Drawing.Size(135,25)
    Font = New-Object System.Drawing.Font('Tahoma',10)
}
$Form.Controls.Add($ExpirationOutput)

$InputTextbox.Add_TextChanged({
    [void](Invoke-OwnedClipboardCleanup -retryOnFailure $true)
    Clear-RetrievedCredential
})

# Drive redirection checkbox
$EnableDriveCheckbox = [System.Windows.Forms.CheckBox]@{
    Text = "Redirect C:\"
    Checked = $true
    Location = New-Object System.Drawing.Point(23, 177)
    Size = New-Object System.Drawing.Size(120, 20)
    Font = New-Object System.Drawing.Font('Tahoma',11)
}
$Form.Controls.Add($EnableDriveCheckbox)

# Connect button
$ConnectButton = [System.Windows.Forms.Button]@{
    Text = "Connect!"
    Location = New-Object System.Drawing.Point(150,174)
    Size = New-Object System.Drawing.Size(135,26)
    Font = New-Object System.Drawing.Font('Tahoma',12)
}
$ConnectButton.Add_Click({
    $hostname = $InputTextbox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($hostname)) {
        [void][System.Windows.Forms.MessageBox]::Show(
            "Enter a hostname before connecting.",
            "Invalid Hostname",
            "OK",
            "Error")
        return
    }

    if (-not (Test-RetrievedCredentialMatchesHostname `
        -RetrievedState $script:RetrievedCredential `
        -Hostname $hostname)) {
        $retrievedCredential = $null
        try {
            Clear-RetrievedCredential
            $retrievedCredential = Get-LapsCredential -hostname $hostname
            Set-RetrievedCredential -RetrievedState $retrievedCredential
        } catch {
            Clear-RetrievedCredential
            Show-LapsRetrievalError -errorRecord $_
            return
        } finally {
            $retrievedCredential = $null
        }
    }

    if (-not (Test-Hostname -hostname $hostname)) {
        [void][System.Windows.Forms.MessageBox]::Show(
            "Hostname '$hostname' not found in DNS.",
            "Connection Error",
            "OK",
            "Error")
        return
    }

    $connectionCredential = $script:RetrievedCredential
    try {
        [void](Invoke-OwnedClipboardCleanup -retryOnFailure $true)
        Connect-RDP `
            -hostname $hostname `
            -account $connectionCredential.RdpAccount `
            -password $connectionCredential.LapsResult.Password `
            -redirectDrive $EnableDriveCheckbox.Checked
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show(
            "Unable to start or complete the RDP connection safely.",
            "Connection Error",
            "OK",
            "Error")
    } finally {
        $connectionCredential = $null
        Clear-RetrievedCredential
    }
})
$Form.Controls.Add($ConnectButton)

# Activate and show form
$Form.Add_Shown({$Form.Activate()})
$Form.Add_FormClosing({
    $clipboardCleanupSucceeded = Invoke-OwnedClipboardCleanup
    if (-not $clipboardCleanupSucceeded) {
        [void][System.Windows.Forms.MessageBox]::Show(
            "The copied password could not be verified or cleared from the clipboard.",
            "Clipboard Warning",
            "OK",
            "Warning")
    }

    if ($null -ne $script:ClipboardCleanupTimer) {
        $script:ClipboardCleanupTimer.Stop()
        $script:ClipboardCleanupTimer.Dispose()
        $script:ClipboardCleanupTimer = $null
    }

    Clear-RetrievedCredential
})
[void] $Form.ShowDialog()
