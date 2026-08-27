# Get-LAPS-pass

Get-LAPS-pass is a compact Windows PowerShell 5.1 WinForms utility for Service Desk workflows. It retrieves a Windows LAPS credential, displays the effective local account and password expiration, supports controlled password copying, and can launch Remote Desktop with a temporary credential.

![Get-LAPS-pass credential interface](GUI.png)

## Requirements

Runtime requirements:

- Windows 10 or Windows 11;
- Windows LAPS PowerShell cmdlets, including `Get-LapsADPassword`;
- permission to read and decrypt the target computer's Windows LAPS password;
- Active Directory and DNS connectivity appropriate for the environment;
- the built-in Remote Desktop Connection client (`mstsc.exe`).

Running the source directly requires Windows PowerShell 5.1. The application does not install or import missing modules automatically.

## Run from source

Keep these files together in one directory:

- `Get_LAPS_pass.ps1`;
- `GetLapsPass.Core.psm1`;
- a local `config.json`.

Create the local configuration from the tracked template, edit it for the environment, and launch the script with Windows PowerShell 5.1:

```powershell
Copy-Item .\config.example.json .\config.json
powershell.exe -NoProfile -File .\Get_LAPS_pass.ps1
```

## Use a packaged release

Release binaries are distributed through versioned release packages rather than tracked in Git. `Get_LAPS_pass.exe` is produced by `build.ps1`.

1. Download `Get-LAPS-pass-2.0.0.zip` and its sibling `Get-LAPS-pass-2.0.0.zip.sha256` file.
2. Verify the ZIP's SHA-256 hash against the `.sha256` file.
3. Extract the ZIP.
4. Keep `GetLapsPass.Core.psm1` beside `Get_LAPS_pass.exe`.
5. Copy `config.example.json` to `config.json` in that same application directory and edit the local values.
6. Run `Get_LAPS_pass.exe`.

The package's `SHA256SUMS.txt` contains SHA-256 hashes for the release payload files. `config.json` is intentionally excluded from both Git and release packages.

## Configuration

`config.example.json` is the safe tracked template. The application requires a valid `config.json` beside the script or executable and resolves it from the application directory, not the process's current working directory.

| Setting | Behavior |
| --- | --- |
| `SearchTemplate` | Prepopulates the hostname input and may be empty. |
| `UserForConnect` | Supplies only the local RDP account name when Windows LAPS does not return a usable `Account` value. |

`UserForConnect` never substitutes for a missing, unauthorized, or undecryptable password. A missing or malformed configuration produces a controlled startup error.

## Credential and UI behavior

- **Get Password** retrieves exactly one Windows LAPS result for the trimmed hostname.
- A non-empty Windows LAPS `Account` is used for RDP; `UserForConnect` is account-name fallback only.
- The password is read-only and masked by default; **Show/Hide** changes only its visibility.
- **Copy** is enabled only while a valid credential is loaded.
- **RDP account** shows the effective `hostname\account` value.
- **Expiration** shows `ExpirationTimestamp`, or `Not available` when the successful result has no timestamp.
- **Connect** can retrieve a credential automatically when one is not already loaded for the current hostname.

Changing the hostname clears the loaded credential and displayed fields. Credential state is also cleared after the RDP lifecycle and when the form closes. Passwords and complete LAPS results are not written to logs or configuration files.

## Clipboard behavior

After **Copy**, the application retains an in-memory SHA-256 fingerprint of the copied value. After 30 seconds it clears the clipboard only if the current clipboard text still matches that fingerprint. Content copied later by the user or another application is left untouched.

Ownership-verified cleanup is also attempted when the hostname changes, before RDP starts, and when the form closes. Clipboard access can be temporarily unavailable, so cleanup is best-effort and never blindly clears unverified content.

## RDP and temporary credential lifecycle

Each connection uses:

- the exact Generic Credential Manager target `TERMSRV/<hostname>`;
- a unique non-secret ownership marker on the temporary session credential;
- a unique temporary `.rdp` file containing connection settings and username, but not the password;
- full-screen RDP with clipboard redirection enabled and smart-card redirection disabled;
- `C:\` drive redirection only when **Redirect C:\** is checked.

If a Generic credential already exists for the exact target, the application does not modify it or launch RDP. During normal `finally` cleanup, the temporary credential is deleted only if its ownership marker still matches, and the temporary `.rdp` file is removed. A credential changed by another process is left untouched.

MSTSC is launched synchronously. Normal cleanup runs after it exits or when an error occurs, but in-process cleanup cannot be guaranteed after a crash, forced termination, shutdown, or loss of the Windows logon session.

### Remote Desktop security dialog

Starting with the April 2026 Windows security update, opening the unsigned temporary `.rdp` file can show the Remote Desktop Connection security dialog with **Unknown publisher** and approval choices for requested local-resource redirections. This is expected. Get-LAPS-pass does not suppress the dialog or bypass Windows security controls.

The user must review the requested clipboard and, when enabled, drive redirection before connecting.

## Reproducible build

The release build requires an x64 Windows PowerShell 5.1 process and these exact module versions:

- PS2EXE `1.0.18`;
- Pester `3.4.0`;
- PSScriptAnalyzer `1.25.0`.

The build does not install or update modules. From the repository root, run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\build.ps1
```

The fail-closed workflow verifies source inputs, runs the offline Pester suite, runs PSScriptAnalyzer with the project settings, builds the x64 no-console executable, validates the package, and generates SHA-256 manifests.

The workflow pins these tool versions and reproduces the same build, validation, packaging, and hashing process. PS2EXE/CodeDOM does not guarantee byte-for-byte identical executable output across separate builds, and ZIP files likewise must not be assumed to have identical SHA-256 hashes. `SHA256SUMS.txt` and the ZIP `.sha256` file authenticate the artifacts produced by that specific release build.

For v2.0.0 the output is:

```text
dist/
├── Get-LAPS-pass-2.0.0/
│   ├── Get_LAPS_pass.exe
│   ├── GetLapsPass.Core.psm1
│   ├── config.example.json
│   ├── README.md
│   ├── LICENSE
│   └── SHA256SUMS.txt
├── Get-LAPS-pass-2.0.0.zip
└── Get-LAPS-pass-2.0.0.zip.sha256
```

The build never copies a local `config.json`. The ignored `dist/` directory contains generated release artifacts and should not be committed.

## Repository layout

```text
get-laps-pass/
├── .gitignore
├── AGENTS.md
├── build.ps1
├── config.example.json
├── Get_LAPS_pass.ps1
├── GetLapsPass.Core.psm1
├── GUI.png
├── LICENSE
├── PSScriptAnalyzerSettings.psd1
├── README.md
└── tests/
    └── GetLapsPass.Core.Tests.ps1
```

The repository contains source and build workflow only. Generated executables belong in versioned release packages.

## License

This project is licensed under the MIT License. See `LICENSE`.
