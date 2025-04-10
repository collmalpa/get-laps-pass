# Load configuration
$config = Get-Content -Raw -Path "config.json" | ConvertFrom-Json

# Load required .NET assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ──────── FUNCTIONS ────────

function Get-LapsPassword {
    param ([string]$hostname)
    try {
        return (Get-LapsADPassword $hostname -AsPlainText).Password
    } catch {
        return $null
    }
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
        [string]$password,
        [bool]$redirectDrive
    )

    $user = "$hostname\$($config.UserForConnect)"
    cmdkey /generic:"$hostname" /user:"$user" /pass:"$password"

    if ($redirectDrive) {
        $rdpFile = "$env:TEMP\$hostname.rdp"
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
redirectsmartcards:i:1
drivestoredirect:s:C:\
full address:s:$hostname
username:s:$user
prompt for credentials:i:0
authentication level:i:2
enablecredsspsupport:i:1
"@
        $rdpContent | Set-Content -Path $rdpFile -Encoding Unicode
        Start-Process -FilePath "mstsc.exe" -ArgumentList $rdpFile -Wait
    } else {
        Start-Process -FilePath "mstsc.exe" -ArgumentList "/v:$hostname", "/f" -Wait
    }

    cmdkey /delete:$hostname
}

# ──────── FORM DESIGN ────────

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
    $hostname = $InputTextbox.Text
    $password = Get-LapsPassword -hostname $hostname
    $PassOutput.Text = if ($password) { $password } else { "Error" }
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
    $hostname = $InputTextbox.Text

    if (-not $PassOutput.Text -or $PassOutput.Text -eq "Error") {
        $PassOutput.Text = Get-LapsPassword -hostname $hostname
    }

    if ($PassOutput.Text -eq "Error" -or -not $PassOutput.Text) {
        [System.Windows.Forms.MessageBox]::Show("Unable to retrieve password.","Error","OK","Error")
        return
    }

    if (-not (Validate-Hostname -hostname $hostname)) {
        [System.Windows.Forms.MessageBox]::Show("Hostname '$hostname' not found in DNS.","Connection Error","OK","Error")
        return
    }

    Connect-RDP -hostname $hostname -password $PassOutput.Text -redirectDrive $EnableDriveCheckbox.Checked
})
$Form.Controls.Add($ConnectButton)

# Activate and show form
$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()
