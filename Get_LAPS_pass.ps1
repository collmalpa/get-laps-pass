# Load the configuration file
$config = Get-Content -Raw -Path "config.json" | ConvertFrom-Json

# Import modules for working with Windows Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Define the form
$Form = New-Object System.Windows.Forms.Form
$Form.ClientSize = '305,150'
$Form.Text = "Get LAPS pass"
$Form.TopMost = $false
$Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$Form.KeyPreview = $True
$Form.StartPosition = "CenterScreen"

$Form.Add_KeyDown({if ($PSItem.KeyCode -eq "Enter") 
    {
        $StartButton.PerformClick()
    }
})
$Form.Add_KeyDown({if ($PSItem.KeyCode -eq "Escape") 
    {
        $Form.Close()
    }
})

# Label for hostname input
$Label = New-Object System.Windows.Forms.Label
$Label.Text = "Enter hostname:"
$Label.Location = New-Object System.Drawing.Point(20, 13)
$Form.Controls.Add($Label)
$Label.Size = New-Object System.Drawing.Size(130, 20)
$Label.Font = New-Object System.Drawing.Font('Tahoma',12,[System.Drawing.FontStyle]::Regular)

# Textbox for hostname input
$InputTextbox = New-Object System.Windows.Forms.TextBox
$InputTextbox.Location = New-Object System.Drawing.Point(150,10) 
$InputTextbox.Size = New-Object System.Drawing.Size(135,25)
$InputTextbox.Font = New-Object System.Drawing.Font('Tahoma',12,[System.Drawing.FontStyle]::Regular)
$InputTextbox.Multiline = $false
$InputTextbox.AcceptsReturn = $false
$InputTextbox.AcceptsTab = $false
$InputTextbox.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Left
$InputTextbox.Text = $config.SearchTemplate
$InputTextbox.SelectionStart = $InputTextbox.Text.Length
$Form.Controls.Add($InputTextbox)

# Button to start password retrieval
$StartButton = New-Object System.Windows.Forms.Button
$StartButton.Location = New-Object System.Drawing.Point(20,45)
$StartButton.Size = New-Object System.Drawing.Size(120,26)
$StartButton.Font = New-Object System.Drawing.Font('Tahoma',12,[System.Drawing.FontStyle]::Regular)
$StartButton.Text = "Start"
$StartButton.Add_Click({
    $hostname = $InputTextbox.Text
    try {
        $PassOutput.Text = (Get-LapsADPassword $hostname -AsPlainText).Password
    } catch {
        $PassOutput.Text = "Error"
    }
})
$Form.Controls.Add($StartButton)

# Textbox for displaying the password
$PassOutput = New-Object System.Windows.Forms.TextBox
$PassOutput.Location = New-Object System.Drawing.Point(150,80)
$PassOutput.Size = New-Object System.Drawing.Size(135,25)
$PassOutput.ReadOnly = $true
$PassOutput.Font = New-Object System.Drawing.Font('Tahoma',11,[System.Drawing.FontStyle]::Regular)
$Form.Controls.Add($PassOutput)

# Label for the password field (updated to a new name)
$PassLabel = New-Object System.Windows.Forms.Label
$PassLabel.Text = "Admin Password:"
$PassLabel.Location = New-Object System.Drawing.Point(20,85)
$PassLabel.Size = New-Object System.Drawing.Size(130, 25)
$PassLabel.Font = New-Object System.Drawing.Font('Tahoma',11,[System.Drawing.FontStyle]::Regular)
$Form.Controls.Add($PassLabel)

# Button to copy the password
$copyButton = New-Object System.Windows.Forms.Button
$copyButton.Location = New-Object System.Drawing.Point(150,45)
$copyButton.Size = New-Object System.Drawing.Size(135,26)
$copyButton.Text = "Copy"
$copyButton.Font = New-Object System.Drawing.Font('Tahoma',12,[System.Drawing.FontStyle]::Regular)
$copyButton.Add_Click({
    $PassOutput.Text | clip
})
$Form.Controls.Add($copyButton)

# Button to connect
$ConnectButton = New-Object System.Windows.Forms.Button
$ConnectButton.Location = New-Object System.Drawing.Point(20,112)
$ConnectButton.Size = New-Object System.Drawing.Size(265,26)
$ConnectButton.Font = New-Object System.Drawing.Font('Tahoma',14,[System.Drawing.FontStyle]::Regular)
$ConnectButton.Text = "Connect!"
$ConnectButton.Add_Click({
    $hostname = $InputTextbox.Text

    # If the password has not been retrieved yet, retrieve it first
    if ([string]::IsNullOrEmpty($PassOutput.Text) -or $PassOutput.Text -eq "Error") {
        try {
            $PassOutput.Text = (Get-LapsADPassword $hostname -AsPlainText).Password
        } catch {
            $PassOutput.Text = "Error"
        }
    }

    if ($PassOutput.Text -ne "Error") {
        $userforconnect = "$hostname\$($config.UserForConnect)"
        $passforconnect = $PassOutput.Text
        cmdkey /generic:"$hostname" /user:"$userforconnect" /pass:"$passforconnect"
        Start-Process -FilePath "mstsc.exe" -ArgumentList "/v:$hostname","/f" -Wait
        cmdkey /delete:$hostname
    } else {
        [System.Windows.Forms.MessageBox]::Show("Unable to retrieve password.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
$Form.Controls.Add($ConnectButton)

$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()
