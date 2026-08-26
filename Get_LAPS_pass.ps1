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
$configPath = Join-Path -Path $applicationDirectory -ChildPath 'config.json'

try {
    if (-not (Test-Path -LiteralPath $configPath -PathType Leaf)) {
        throw (New-Object System.IO.FileNotFoundException)
    }

    $config = Get-Content -Raw -LiteralPath $configPath -ErrorAction Stop |
        ConvertFrom-Json -ErrorAction Stop

    $configPropertyNames = @($config.PSObject.Properties.Name)
    if (
        $null -eq $config -or
        $configPropertyNames -notcontains 'SearchTemplate' -or
        $configPropertyNames -notcontains 'UserForConnect' -or
        -not ($config.SearchTemplate -is [string]) -or
        -not ($config.UserForConnect -is [string]) -or
        [string]::IsNullOrWhiteSpace($config.UserForConnect)
    ) {
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

# ──────── FUNCTIONS ────────

function Get-LapsErrorMessage {
    param ([string]$category)

    switch ($category) {
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
    param ([string]$category)

    $exception = New-Object System.InvalidOperationException(
        (Get-LapsErrorMessage -category $category))
    $exception.Data['LapsErrorCategory'] = $category
    return $exception
}

function Get-LapsExceptionCategory {
    param ([System.Management.Automation.ErrorRecord]$errorRecord)

    $exception = $errorRecord.Exception
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

    $errorCategory = [string]$errorRecord.CategoryInfo.Category
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

    if ($errorRecord.FullyQualifiedErrorId -eq 'CommandNotFoundException') {
        return 'ModuleUnavailable'
    }

    return 'Unexpected'
}

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

    if ($lapsResults.Count -eq 0) {
        throw (New-LapsRetrievalException -category 'ComputerNotFound')
    }
    if ($lapsResults.Count -ne 1) {
        throw (New-LapsRetrievalException -category 'Unexpected')
    }

    $lapsResult = $lapsResults[0]
    $decryptionStatus = [string]$lapsResult.DecryptionStatus
    if ($decryptionStatus -eq 'Unauthorized') {
        throw (New-LapsRetrievalException -category 'AccessDenied')
    }
    if ($decryptionStatus -notin @('Success', 'NotApplicable')) {
        throw (New-LapsRetrievalException -category 'PasswordUnavailable')
    }
    if ([string]::IsNullOrEmpty([string]$lapsResult.Password)) {
        throw (New-LapsRetrievalException -category 'PasswordUnavailable')
    }

    $lapsAccount = [string]$lapsResult.Account
    if ([string]::IsNullOrWhiteSpace($lapsAccount)) {
        $rdpAccount = $config.UserForConnect.Trim()
        $accountSource = 'ConfigFallback'
    } else {
        $rdpAccount = $lapsAccount.Trim()
        $accountSource = 'LAPS'
    }

    return [pscustomobject]@{
        RequestedHostname = $hostname
        LapsResult = $lapsResult
        RdpAccount = $rdpAccount
        AccountSource = $accountSource
        ExpirationTimestamp = $lapsResult.ExpirationTimestamp
    }
}

function Clear-RetrievedCredential {
    $script:RetrievedCredential = $null

    if ($null -ne $PassOutput) {
        $PassOutput.Clear()
    }
}

function Set-RetrievedCredential {
    param ([psobject]$credential)

    $script:RetrievedCredential = $credential
    $PassOutput.Text = [string]$credential.LapsResult.Password
}

function Validate-Hostname {
    param ([string]$hostname)
    try {
        Resolve-DnsName -Name $hostname -ErrorAction Stop | Out-Null
        return $true
    } catch {
        return $false
    }
}

function Connect-RDP {
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

        $driveStoreRedirect = if ($redirectDrive) { 'C:\;' } else { '' }
        $rdpFileName = '{0}.rdp' -f [guid]::NewGuid().ToString('N')
        $rdpFile = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath $rdpFileName
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
full address:s:$hostname
username:s:$user
prompt for credentials:i:0
authentication level:i:2
enablecredsspsupport:i:1
"@
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

# ──────── FORM DESIGN ────────

$script:RetrievedCredential = $null

$Form = New-Object System.Windows.Forms.Form
$Form.ClientSize = '305,150'
$Form.Text = "Get LAPS pass"
$Form.FormBorderStyle = 'FixedSingle'
$Form.StartPosition = "CenterScreen"
$Form.KeyPreview = $true
$Form.TopMost = $false

# Enter/Escape key support
$Form.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") { $StartButton.PerformClick() }
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

# Start button
$StartButton = [System.Windows.Forms.Button]@{
    Text = "Start"
    Location = New-Object System.Drawing.Point(20,45)
    Size = New-Object System.Drawing.Size(120,26)
    Font = New-Object System.Drawing.Font('Tahoma',12)
}
$StartButton.Add_Click({
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
        Set-RetrievedCredential -credential $retrievedCredential
    } catch {
        Clear-RetrievedCredential
        Show-LapsRetrievalError -errorRecord $_
    } finally {
        $retrievedCredential = $null
    }
})
$Form.Controls.Add($StartButton)

# Copy button
$copyButton = [System.Windows.Forms.Button]@{
    Text = "Copy"
    Location = New-Object System.Drawing.Point(150,45)
    Size = New-Object System.Drawing.Size(135,26)
    Font = New-Object System.Drawing.Font('Tahoma',12)
}
$copyButton.Add_Click({ $PassOutput.Text | clip })
$Form.Controls.Add($copyButton)

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
    Location = New-Object System.Drawing.Point(150,80)
    Size = New-Object System.Drawing.Size(135,25)
    Font = New-Object System.Drawing.Font('Tahoma',11)
}
$Form.Controls.Add($PassOutput)

$InputTextbox.Add_TextChanged({
    if ($null -ne $script:RetrievedCredential) {
        Clear-RetrievedCredential
    }
})

# Drive redirection checkbox
$EnableDriveCheckbox = [System.Windows.Forms.CheckBox]@{
    Text = "Redirect C:\"
    Checked = $true
    Location = New-Object System.Drawing.Point(23, 115)
    Size = New-Object System.Drawing.Size(120, 20)
    Font = New-Object System.Drawing.Font('Tahoma',11)
}
$Form.Controls.Add($EnableDriveCheckbox)

# Connect button
$ConnectButton = [System.Windows.Forms.Button]@{
    Text = "Connect!"
    Location = New-Object System.Drawing.Point(150,112)
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

    if (
        $null -eq $script:RetrievedCredential -or
        $script:RetrievedCredential.RequestedHostname -cne $hostname
    ) {
        $retrievedCredential = $null
        try {
            Clear-RetrievedCredential
            $retrievedCredential = Get-LapsCredential -hostname $hostname
            Set-RetrievedCredential -credential $retrievedCredential
        } catch {
            Clear-RetrievedCredential
            Show-LapsRetrievalError -errorRecord $_
            return
        } finally {
            $retrievedCredential = $null
        }
    }

    if (-not (Validate-Hostname -hostname $hostname)) {
        [void][System.Windows.Forms.MessageBox]::Show(
            "Hostname '$hostname' not found in DNS.",
            "Connection Error",
            "OK",
            "Error")
        return
    }

    $connectionCredential = $script:RetrievedCredential
    try {
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
$Form.Add_FormClosed({ Clear-RetrievedCredential })
[void] $Form.ShowDialog()
