# Windows App Crash Doctor

`Windows App Crash Doctor` is a PowerShell-first troubleshooting tool for the annoying Windows pattern where many apps seem to fail at once and the crash logs keep pointing at shared runtime DLLs such as `KERNELBASE.dll`, `MSVCP140.dll`, `VCRUNTIME140.dll`, or `ucrtbase.dll`.

Instead of treating every crash as an isolated app bug, it looks for shared suspects like:

- enterprise VPN and network hook stacks
- OEM background services in crash loops
- third-party IME hook layers
- common runtime and system-health warning signs

## What it does

- reads recent `Application Error (Event ID 1000)` crash events
- summarizes the most frequently crashing apps and modules
- checks installed software, services, processes, startup items, Winsock providers, and driver catalog entries
- scores suspect profiles such as `Sangfor / EasyConnect`, `ROG Live Service`, and third-party Chinese IMEs
- generates human-readable and JSON reports
- can run a limited `apply-fix` workflow for high-confidence targets that have a known repair plan

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- Windows event log access
- Administrator privileges for some repair actions

## Quick start

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\AppCrashDoctor.ps1 -Mode scan
```

Generate a richer report:

```powershell
.\AppCrashDoctor.ps1 `
  -Mode suggest-fix `
  -Days 14 `
  -ReportPath .\reports\last-scan.md `
  -JsonPath .\reports\last-scan.json
```

Preview an automatic fix plan without changing the machine:

```powershell
.\AppCrashDoctor.ps1 `
  -Mode apply-fix `
  -Target Sangfor `
  -ResetWinsock `
  -WhatIf
```

Run the real repair plan for a detected target:

```powershell
Start-Process powershell -Verb RunAs -ArgumentList @(
  '-ExecutionPolicy', 'Bypass',
  '-File', '.\AppCrashDoctor.ps1',
  '-Mode', 'apply-fix',
  '-Target', 'Sangfor',
  '-ResetWinsock'
)
```

## Modes

### `scan`

Read-only system scan with compact console output plus optional saved reports.

### `suggest-fix`

Read-only scan with more guidance printed to the console.

### `apply-fix`

Runs a repair plan for specific targets.

Current implemented target:

- `Sangfor`

Current suggestion-only targets:

- `ROGLiveService`
- `ChineseImeStack`

## How scoring works

Each suspect profile gains confidence from a mix of:

- installed program matches
- service matches
- running process matches
- startup item matches
- recent crash evidence
- Winsock provider matches
- driver catalog matches

The tool labels each suspect as `low`, `medium`, or `high` confidence.

## What `apply-fix` for `Sangfor` does

The built-in plan tries to:

- stop known Sangfor helper processes
- stop known Sangfor services if possible
- run vendor uninstallers found in registry and common install paths
- remove matching drivers with `pnputil`
- delete the `SangforVnic` service
- optionally reset Winsock
- remove the leftover installation directory

## Safety notes

- Always start with `scan` or `suggest-fix`.
- Use `-WhatIf` before `apply-fix`.
- Treat `apply-fix` as a guided cleanup tool, not a universal magic button.
- Corporate VPN, EDR, remote support, or OEM control software may be required in your environment. Confirm before removal.

## Publish to GitHub

```powershell
git init
git add .
git commit -m "Initial commit"
gh repo create windows-app-crash-doctor --source . --push --public
```

Switch `--public` to `--private` if you do not want the repository to be public.

## Known limitations

- Event log parsing depends on English-style `Application Error` message fields.
- The repair logic is intentionally conservative and currently has one automatic target.
- System-file repair steps like `DISM` and `sfc` are recommended, not executed automatically.
- Some driver or service cleanup steps may need elevation or a reboot.
