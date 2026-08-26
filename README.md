# Get-LAPS-pass

Get-LAPS-pass is a small Windows PowerShell 5.1 WinForms utility for Service Desk workflows. It retrieves a Windows LAPS credential for a computer, displays the effective local account and password expiration, optionally shows or copies the password, and can start a Remote Desktop connection with that credential.

> [!IMPORTANT]
> `Get_LAPS_pass.ps1` is the authoritative current implementation. The tracked `Get_LAPS_pass.exe` is a legacy artifact that predates the current security and user-interface changes. Do not treat the EXE as the current release; it will remain legacy until a reproducible build process is available.

## Requirements

- Windows 10 or Windows 11
- Windows PowerShell 5.1
- Windows LAPS PowerShell cmdlets, including `Get-LapsADPassword`
- Permission to retrieve and decrypt the target computer's Windows LAPS password
- Active Directory and DNS connectivity appropriate for the environment
- The built-in Remote Desktop Connection client (`mstsc.exe`)

The application does not install or import missing modules automatically.

## Installation and source-first usage

1. Clone the repository:

   ```powershell
   git clone https://github.com/collmalpa/laps-password-rdp-tool.git
   cd laps-password-rdp-tool
   ```

2. Create the local configuration:

   ```powershell
   Copy-Item .\config.example.json .\config.json
   ```

3. Edit `config.json` for the local environment.

4. Launch the authoritative script with Windows PowerShell 5.1:

   ```powershell
   powershell.exe -NoProfile -File .\Get_LAPS_pass.ps1
   ```

The bundled EXE is intentionally not advertised as a current build.

## Configuration

`config.json` is intentionally ignored by Git and must exist beside the script or application executable. The application resolves it relative to that application directory, not the process's current working directory.

Start with `config.example.json`:

```json
{
  "SearchTemplate": "",
  "UserForConnect": "Administrator"
}
```

| Setting | Behavior |
| --- | --- |
| `SearchTemplate` | Prepopulates the hostname input. It may be an empty string. |
| `UserForConnect` | Supplies only the local RDP account name when Windows LAPS does not return a usable `Account` value. |

`UserForConnect` is never a fallback for a missing, unavailable, unauthorized, or undecryptable password. A missing or malformed `config.json` produces a controlled startup error and the application does not continue.

## Current UI behavior

- **Get Password** retrieves the credential for the trimmed hostname.
- The password is read-only and masked by default.
- **Show/Hide** changes only password visibility.
- **Copy** is enabled only while a valid credential is loaded.
- **RDP account** displays the effective `hostname\account` value.
- **Expiration** displays the LAPS expiration timestamp using the current Windows culture, or `Not available` when a successful result has no timestamp.
- **Redirect C:\** controls whether the generated RDP configuration requests local `C:\` drive redirection.
- **Connect** remains available without first pressing **Get Password**; it retrieves a valid credential automatically when required.

Changing the hostname invalidates and clears the displayed credential. Credential-related UI state is also cleared after the RDP lifecycle and when the form closes.

## Windows LAPS behavior

The application requests exactly one result with:

```powershell
Get-LapsADPassword -Identity <hostname> -AsPlainText -ErrorAction Stop
```

The complete returned LAPS result is retained in application memory for the active credential lifecycle:

- A non-empty Windows LAPS `Account` value is used as the RDP local account.
- `config.UserForConnect` is used only when `Account` is unavailable or whitespace.
- `ExpirationTimestamp` is retained for display.
- An unauthorized decryption status or a null/empty password is treated as failure.
- Module, computer-not-found, access-denied, password-unavailable, and unexpected failures produce fixed, categorized messages instead of raw exception details.

The application does not display complete LAPS result objects or sensitive directory metadata in errors.

## RDP credential lifecycle

For a connection, the application uses the built-in Windows Credential Manager API directly:

- It works with a temporary `CRED_TYPE_GENERIC` credential.
- The exact target is `TERMSRV/<hostname>`.
- The credential uses session persistence.
- A unique, non-secret ownership marker identifies the credential created for that connection.
- Cleanup reads the current credential metadata and deletes it only when the marker still matches.
- If an exact Generic credential already exists, Connect stops before launching RDP and leaves the saved credential untouched.
- If another process changes the credential, the application leaves it untouched and shows a safe cleanup warning.

The application attempts to remove its temporary credential during its normal `finally` cleanup path after MSTSC exits or an error occurs. Cleanup is not guaranteed after abrupt process termination, a crash, forced process kill, machine shutdown, or loss of the Windows logon session.

## RDP file and redirection behavior

Every connection uses a unique temporary `.rdp` file. Its relevant settings request:

- a full-screen connection;
- clipboard redirection;
- no smart-card redirection;
- only `C:\` drive redirection when **Redirect C:\** is checked;
- no drive redirection when **Redirect C:\** is unchecked.

The generated file contains connection settings and the effective username, but not the LAPS password. The application attempts to remove the file during its normal cleanup path.

### Remote Desktop Connection security dialog

Current Windows versions may show a Remote Desktop Connection security dialog when the unsigned temporary `.rdp` file is opened, including an unknown-publisher warning and approval choices for requested local-resource redirection. This is expected. Get-LAPS-pass intentionally does not suppress the dialog or bypass Windows security controls.

The user must review and approve the requested **Clipboard** and, when enabled, **Drives** redirection. Approval cannot expand the request beyond the generated settings: drive redirection remains limited to `C:\`, or disabled, according to the checkbox.

## Clipboard behavior

**Copy** uses the WinForms clipboard API. For cleanup ownership, the application retains only an in-memory SHA-256 fingerprint of the exact copied text; it does not keep a second long-lived plaintext clipboard-password value.

After 30 seconds, the application reads the current clipboard and clears it only when its fingerprint still matches the value copied by Get-LAPS-pass. If the user or another application has copied something else, that content is left untouched.

Ownership-verified cleanup is also attempted when the hostname changes, immediately before the blocking RDP launch, and while the form is closing. Clipboard access can be temporarily unavailable, so cleanup is best-effort, uses bounded short timer retries where applicable, and never blindly clears unverified clipboard content.

## Limitations

- Unsigned temporary RDP files may show the expected Windows security warning.
- MSTSC is launched synchronously, so the application UI remains blocked during the RDP lifecycle.
- An existing exact Generic `TERMSRV/<hostname>` credential blocks the connection and must be handled manually.
- In-process credential, RDP-file, and clipboard cleanup cannot be guaranteed after abrupt termination.
- Clipboard cleanup is best-effort when another process temporarily owns the Windows clipboard.
- The tracked `Get_LAPS_pass.exe` is legacy and does not represent the current script until reproducible build automation is established.

## Repository structure

```text
get-laps-pass/
├── .gitignore
├── AGENTS.md
├── config.example.json
├── Get_LAPS_pass.ps1
├── Get_LAPS_pass.exe       # legacy artifact
├── GUI.png                 # obsolete screenshot; replacement pending
├── LICENSE
└── README.md
```

A local `config.json` must be created beside the script/application. It is intentionally ignored and is not part of the tracked repository layout.

## Screenshot status

`GUI.png` is retained for now but shows an obsolete interface and is not presented as the current UI. A sanitized screenshot of the current interface will be supplied manually in a later update.

## License

This project is licensed under the MIT License. See `LICENSE`.
