# Get-LAPS-pass

**Get-LAPS-pass** is a PowerShell script designed to streamline the process of retrieving and utilizing Local Administrator Password Solution (LAPS) passwords. The script provides a user-friendly graphical interface (GUI) for inputting hostnames, retrieving passwords, and connecting to remote systems using Remote Desktop Protocol (RDP).

## Features

- **Retrieve LAPS Passwords** by entering a hostname.
- **Display the password** in a read-only field within the GUI.
- **Copy password to clipboard** with a single click.
- **Connect via RDP** using the retrieved credentials.
- **Redirect C:\\ drive** during RDP (enabled by default).
- **Error handling** for password lookup and connectivity issues.
- **DNS name resolution check** before launching the RDP session.
- **Keyboard shortcuts**: Press Enter to start, Escape to exit.

## Prerequisites

1. **Windows environment** - this script is intended for Windows.
2. **PowerShell 5.1 or later**
3. **LAPS Module**: The script uses the `Get-LapsADPassword` cmdlet, which requires the LAPS PowerShell module to be installed.
4. **Permission to read LAPS passwords** in Active Directory.
5. A configuration file named `config.json` in the same directory:
   ```json
   {
       "SearchTemplate": "template-hostname",
       "UserForConnect": "administrator"
   }
   ```
   
## Usage

Run the script:
```powershell
.\Get-LAPS-pass.ps1
```

Or launch the precompiled executable (`Get-LAPS-pass.exe`) for easier use without opening PowerShell.

### In the GUI:

1. Enter the hostname of the target system.
2. Click **Start** to retrieve the LAPS password.
3. Click **Copy** to copy the password to the clipboard.
4. Optionally, use the **Redirect C:\\** checkbox to control drive redirection.
5. Click **Connect!** to start an RDP session using the retrieved credentials.

## GUI Overview

- **Enter hostname** — Input field for the target system name.
- **Start** — Button to retrieve the password.
- **Admin Password** — Displays the retrieved password.
- **Copy** — Copies the password to clipboard.
- **Redirect C:\\** — Enables/disables local drive redirection in the RDP session.
- **Connect!** — Launches RDP with saved credentials.

### Screenshot

![Get-LAPS-pass GUI](GUI.png)

## Error Handling

- If the hostname cannot be resolved via DNS, an error message will appear.
- If password retrieval fails, the password field will display `Error`.

## Configuration

Customize behavior by editing `config.json`:
- `SearchTemplate`: Default value shown in the hostname field.
- `UserForConnect`: Default username for RDP (`hostname\username` format).

## License

This project is licensed under the [MIT License](LICENSE).

## Contributing

Contributions are welcome! Feel free to open issues or submit pull requests.

## Acknowledgments

[Microsoft LAPS](https://www.microsoft.com/en-us/download/details.aspx?id=46899) for the password management solution.

## Disclaimer

Use this tool responsibly and in accordance with your organization's security policies.
