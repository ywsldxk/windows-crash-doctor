# Windows Crash Doctor

[English](README.md) | [简体中文](README.zh-CN.md)

Diagnose and fix shared-cause Windows app crashes by checking Event Viewer, services, processes, startup items, Winsock providers, driver packages, and common hook-heavy software such as VPN clients, OEM background services, and third-party IMEs.

This project is designed for the frustrating Windows pattern where many apps fail around the same time and crash logs keep pointing to shared DLLs like:

- `KERNELBASE.dll`
- `MSVCP140.dll`
- `VCRUNTIME140.dll`
- `ucrtbase.dll`

Those DLLs are often where the crash surfaces, not where the real problem starts.

## Features

- Scans recent `Application Error` events from Event Viewer.
- Summarizes the most frequently crashing apps and modules.
- Detects likely shared suspects from installed programs, services, processes, startup items, Winsock providers, and driver packages.
- Scores suspect profiles and labels them as `low`, `medium`, or `high` confidence.
- Exports both Markdown and JSON reports.
- Supports a limited `apply-fix` workflow for known targets.

## Current targets

Automatic repair:

- `Sangfor`

Detection and recommendations:

- `ROGLiveService`
- `ChineseImeStack`

## Files

- `WindowsCrashDoctor.ps1`: main CLI entry point
- `AppCrashDoctor.ps1`: compatibility wrapper for the old script name
- `README.zh-CN.md`: Simplified Chinese documentation

## Quick start

Run a read-only scan:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\WindowsCrashDoctor.ps1 -Mode scan
```

Generate Markdown and JSON reports:

```powershell
.\WindowsCrashDoctor.ps1 `
  -Mode suggest-fix `
  -Days 14 `
  -ReportPath .\reports\last-scan.md `
  -JsonPath .\reports\last-scan.json
```

Preview a repair plan without changing anything:

```powershell
.\WindowsCrashDoctor.ps1 `
  -Mode apply-fix `
  -Target Sangfor `
  -ResetWinsock `
  -WhatIf
```

Run a real repair flow from an elevated PowerShell session:

```powershell
Start-Process powershell -Verb RunAs -ArgumentList @(
  '-ExecutionPolicy', 'Bypass',
  '-File', '.\WindowsCrashDoctor.ps1',
  '-Mode', 'apply-fix',
  '-Target', 'Sangfor',
  '-ResetWinsock'
)
```

## Modes

### `scan`

Read-only detection with compact console output.

### `suggest-fix`

Read-only detection plus recommended next steps.

### `apply-fix`

Runs the repair workflow for a selected target.

## What `apply-fix` currently does for `Sangfor`

- Stops known Sangfor helper processes.
- Tries to stop known Sangfor services.
- Runs vendor uninstallers found in common install paths.
- Tries to remove matching driver packages with `pnputil`.
- Tries to delete `SangforVnic`.
- Optionally runs `netsh winsock reset`.
- Tries to delete the leftover install directory.

## Safety notes

- Start with `scan` or `suggest-fix`.
- Use `-WhatIf` before `apply-fix`.
- Some repairs require administrator rights.
- Some driver removals still require a reboot.
- Corporate VPN, EDR, remote support, or OEM software may be required in your environment. Confirm before removal.

## Example use cases

- Many apps start crashing after a VPN client was installed.
- Event Viewer keeps blaming `KERNELBASE.dll`, but the real issue looks shared.
- An OEM helper service is crashing in a loop.
- You want a structured Windows crash triage report instead of random guesswork.

## License

MIT
