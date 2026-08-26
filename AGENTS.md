# Get-LAPS-pass

## Project purpose

Get-LAPS-pass is a small Windows PowerShell WinForms utility used to:
- retrieve Windows LAPS credentials for a computer;
- display/copy the password;
- launch an RDP session using the retrieved local administrator credential.

## Workspace

This repository is expected to live under C:\ScriptsUtils.
Do not access or modify sibling directories outside this repository.

## Compatibility requirements

- Maintain Windows PowerShell 5.1 compatibility.
- Maintain Windows 10/11 compatibility.
- Keep WinForms.
- Do not migrate the project to WPF, C#, WinUI, PowerShell 7, or another framework unless explicitly requested.
- Keep the application small and dependency-light.

## Security requirements

- Never log a LAPS password.
- Never persist a LAPS password to disk.
- Never add passwords to test fixtures, sample files, Git commits, or documentation.
- Minimize the lifetime of plaintext credentials.
- Temporary RDP credentials and files must always be cleaned up.
- Do not contact real Active Directory, retrieve real LAPS passwords, create Windows credentials, or launch real RDP sessions unless explicitly requested.
- Never use real production hostnames in tests.

## Repository rules

- Do not modify or rebuild Get_LAPS_pass.exe unless explicitly requested.
- Do not commit or push changes unless explicitly requested.
- Do not switch branches unless explicitly requested.
- Prefer small focused changes over broad rewrites.
- Do not add external dependencies without approval.
- Preserve existing user-visible behavior unless the task explicitly changes it.

## Development workflow

Before making a significant change:
1. Inspect the relevant existing code.
2. Explain the intended change.
3. Keep the scope limited to the requested task.

After changing PowerShell code:
1. Check PowerShell syntax.
2. Review the Git diff.
3. Run available non-destructive tests.
4. Report what was tested and what still requires manual Windows/AD validation.

## Code quality

- Use meaningful PowerShell function and variable names.
- Prefer terminating errors with -ErrorAction Stop when exceptions need to be handled.
- Do not silently swallow useful errors.
- Keep UI code separate from reusable logic where practical.
- Prefer try/finally for cleanup of credentials and temporary files.