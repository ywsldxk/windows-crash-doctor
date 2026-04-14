[CmdletBinding()]
param(
    [ValidateSet('scan', 'suggest-fix', 'apply-fix')]
    [string]$Mode = 'scan',

    [int]$Days = 7,

    [int]$MaxEvents = 200,

    [string[]]$Target = @(),

    [string]$ReportPath,

    [string]$JsonPath,

    [switch]$ResetWinsock,

    [switch]$WhatIf,

    [switch]$PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-RegexValue {
    param(
        [string]$InputString,
        [string]$Pattern
    )

    if ([string]::IsNullOrWhiteSpace($InputString)) {
        return $null
    }

    $match = [regex]::Match($InputString, $Pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($match.Success) {
        return $match.Groups[1].Value.Trim()
    }

    return $null
}

function Test-AnyMatch {
    param(
        [string[]]$Patterns,
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $false
    }

    foreach ($pattern in $Patterns) {
        if ($Value -match $pattern) {
            return $true
        }
    }

    return $false
}

function Resolve-CommandLine {
    param(
        [string]$CommandLine
    )

    if ([string]::IsNullOrWhiteSpace($CommandLine)) {
        return $null
    }

    $trimmed = $CommandLine.Trim()

    if ($trimmed -match '^\s*"([^"]+)"\s*(.*)$') {
        return [pscustomobject]@{
            FilePath     = $matches[1]
            ArgumentList = $matches[2].Trim()
        }
    }

    if ($trimmed -match '^\s*([^\s]+\.exe)\s*(.*)$') {
        return [pscustomobject]@{
            FilePath     = $matches[1]
            ArgumentList = $matches[2].Trim()
        }
    }

    return [pscustomobject]@{
        FilePath     = $trimmed
        ArgumentList = ''
    }
}

function Get-SystemSummary {
    $computerInfo = $null
    try {
        $computerInfo = Get-ComputerInfo
    }
    catch {
        $computerInfo = $null
    }

    return [pscustomobject]@{
        ComputerName     = $env:COMPUTERNAME
        UserName         = $env:USERNAME
        PowerShell       = $PSVersionTable.PSVersion.ToString()
        IsAdministrator  = Test-IsAdministrator
        WindowsProduct   = if ($computerInfo) { $computerInfo.WindowsProductName } else { $null }
        WindowsVersion   = if ($computerInfo) { $computerInfo.WindowsVersion } else { $null }
        OsBuildNumber    = if ($computerInfo) { $computerInfo.OsBuildNumber } else { $null }
        OsName           = if ($computerInfo) { $computerInfo.OsName } else { $null }
        CollectedAtLocal = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    }
}

function Get-InstalledPrograms {
    $registryPaths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    $items = foreach ($path in $registryPaths) {
        Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
    }

    $items |
        Where-Object {
            $displayNameProperty = $_.PSObject.Properties['DisplayName']
            $displayName = if ($displayNameProperty) { [string]$displayNameProperty.Value } else { '' }
            -not [string]::IsNullOrWhiteSpace($displayName)
        } |
        Select-Object `
            @{ Name = 'DisplayName'; Expression = { if ($_.PSObject.Properties['DisplayName']) { $_.PSObject.Properties['DisplayName'].Value } else { $null } } },
            @{ Name = 'DisplayVersion'; Expression = { if ($_.PSObject.Properties['DisplayVersion']) { $_.PSObject.Properties['DisplayVersion'].Value } else { $null } } },
            @{ Name = 'Publisher'; Expression = { if ($_.PSObject.Properties['Publisher']) { $_.PSObject.Properties['Publisher'].Value } else { $null } } },
            @{ Name = 'InstallDate'; Expression = { if ($_.PSObject.Properties['InstallDate']) { $_.PSObject.Properties['InstallDate'].Value } else { $null } } },
            @{ Name = 'UninstallString'; Expression = { if ($_.PSObject.Properties['UninstallString']) { $_.PSObject.Properties['UninstallString'].Value } else { $null } } },
            @{ Name = 'QuietUninstallString'; Expression = { if ($_.PSObject.Properties['QuietUninstallString']) { $_.PSObject.Properties['QuietUninstallString'].Value } else { $null } } } |
        Sort-Object DisplayName -Unique
}

function Get-ServiceSnapshot {
    Get-Service |
        Select-Object Name, DisplayName, Status
}

function Get-ProcessSnapshot {
    Get-Process -ErrorAction SilentlyContinue |
        Select-Object `
            ProcessName,
            Id,
            @{ Name = 'Path'; Expression = { $_.Path } }
}

function Get-StartupSnapshot {
    try {
        return Get-CimInstance Win32_StartupCommand |
            Select-Object Name, Command, Location
    }
    catch {
        return @()
    }
}

function Get-RecentCrashEvents {
    param(
        [int]$RecentDays,
        [int]$Limit
    )

    $startTime = (Get-Date).AddDays(-1 * [math]::Abs($RecentDays))
    $events = @()

    try {
        $events = Get-WinEvent -FilterHashtable @{
            LogName   = 'Application'
            Id        = 1000
            StartTime = $startTime
        } -MaxEvents $Limit
    }
    catch {
        return @()
    }

    foreach ($event in $events) {
        $message = $event.Message
        [pscustomobject]@{
            TimeCreated         = $event.TimeCreated
            AppName             = Get-RegexValue -InputString $message -Pattern 'Faulting application name:\s*([^,]+),'
            AppPath             = Get-RegexValue -InputString $message -Pattern 'Faulting application path:\s*(.+)'
            FaultingModule      = Get-RegexValue -InputString $message -Pattern 'Faulting module name:\s*([^,]+),'
            FaultingModulePath  = Get-RegexValue -InputString $message -Pattern 'Faulting module path:\s*(.+)'
            ExceptionCode       = Get-RegexValue -InputString $message -Pattern 'Exception code:\s*(0x[0-9A-Fa-f]+)'
            RawMessage          = $message
        }
    }
}

function Get-CrashAppSummary {
    param(
        [object[]]$CrashEvents
    )

    $CrashEvents |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_.AppName) } |
        Group-Object AppName |
        Sort-Object Count -Descending |
        Select-Object `
            @{ Name = 'AppName'; Expression = { $_.Name } },
            Count,
            @{ Name = 'LatestSeen'; Expression = { ($_.Group | Sort-Object TimeCreated -Descending | Select-Object -First 1).TimeCreated } }
}

function Get-CrashModuleSummary {
    param(
        [object[]]$CrashEvents
    )

    $CrashEvents |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_.FaultingModule) } |
        Group-Object FaultingModule |
        Sort-Object Count -Descending |
        Select-Object `
            @{ Name = 'FaultingModule'; Expression = { $_.Name } },
            Count
}

function Get-WinsockCatalogText {
    try {
        return (netsh winsock show catalog | Out-String)
    }
    catch {
        return ''
    }
}

function Get-DriverCatalogText {
    try {
        return (pnputil /enum-drivers | Out-String)
    }
    catch {
        return ''
    }
}

function Get-RepairProfiles {
    return @(
        [pscustomobject]@{
            Id                  = 'Sangfor'
            Name                = 'Sangfor / EasyConnect'
            Description         = 'Enterprise VPN and remote access stack that can install Winsock providers, drivers, and helper services.'
            ProgramPatterns     = @('Sangfor', '深信服', 'EasyConnect')
            ServicePatterns     = @('Sangfor', '深信服')
            ProcessPatterns     = @('Sangfor', 'EasyConnect', 'ECAgent', 'Svpn')
            StartupPatterns     = @('Sangfor', '深信服', 'EasyConnect')
            CrashPatterns       = @('Sangfor', 'EasyConnect')
            WinsockPatterns     = @('Sangfor', '深信服')
            DriverPatterns      = @('Sangfor', '深信服', 'sangforvnic')
            SupportsApplyFix    = $true
            RiskLevel           = 'medium'
            SuggestedActions    = @(
                'Stop Sangfor helper processes and uninstall vendor components.',
                'If network hooks were present, reset Winsock and reboot.',
                'If a virtual NIC driver remains, remove it with pnputil from an elevated console.'
            )
        },
        [pscustomobject]@{
            Id                  = 'ROGLiveService'
            Name                = 'ROG Live Service / Armoury Crate'
            Description         = 'ASUS background services that can enter crash loops and destabilize startup.'
            ProgramPatterns     = @('ROG', 'Armoury Crate', 'ASUS')
            ServicePatterns     = @('ROG Live Service', 'AsusROGLSLService')
            ProcessPatterns     = @('ROGLiveService', 'ArmouryCrate')
            StartupPatterns     = @('ROG', 'Armoury', 'ASUS')
            CrashPatterns       = @('ROGLiveService', 'Armoury')
            WinsockPatterns     = @()
            DriverPatterns      = @()
            SupportsApplyFix    = $false
            RiskLevel           = 'low'
            SuggestedActions    = @(
                'Update Armoury Crate and ASUS support components.',
                'If crashes continue, disable or uninstall ROG Live Service.'
            )
        },
        [pscustomobject]@{
            Id                  = 'ChineseImeStack'
            Name                = 'Third-party Chinese IME stack'
            Description         = 'Input methods can inject UI and text hooks into many desktop apps.'
            ProgramPatterns     = @('Sogou', '搜狗', 'iFly', '讯飞')
            ServicePatterns     = @()
            ProcessPatterns     = @('Sogou', 'iFly')
            StartupPatterns     = @('Sogou', '搜狗', 'iFly', '讯飞')
            CrashPatterns       = @('sogou', 'ifly')
            WinsockPatterns     = @()
            DriverPatterns      = @()
            SupportsApplyFix    = $false
            RiskLevel           = 'low'
            SuggestedActions    = @(
                'Temporarily keep only Microsoft IME and retest the crashing apps.',
                'Reinstall the IME that you actually use after stability is confirmed.'
            )
        }
    )
}

function Get-MatchedPrograms {
    param(
        [object[]]$Programs,
        [string[]]$Patterns
    )

    return @(
        $Programs | Where-Object {
            Test-AnyMatch -Patterns $Patterns -Value ($_.DisplayName + ' ' + $_.Publisher)
        }
    )
}

function Get-MatchedServices {
    param(
        [object[]]$Services,
        [string[]]$Patterns
    )

    return @(
        $Services | Where-Object {
            Test-AnyMatch -Patterns $Patterns -Value ($_.Name + ' ' + $_.DisplayName)
        }
    )
}

function Get-MatchedProcesses {
    param(
        [object[]]$Processes,
        [string[]]$Patterns
    )

    return @(
        $Processes | Where-Object {
            Test-AnyMatch -Patterns $Patterns -Value ($_.ProcessName + ' ' + $_.Path)
        }
    )
}

function Get-MatchedStartupItems {
    param(
        [object[]]$StartupItems,
        [string[]]$Patterns
    )

    return @(
        $StartupItems | Where-Object {
            Test-AnyMatch -Patterns $Patterns -Value ($_.Name + ' ' + $_.Command + ' ' + $_.Location)
        }
    )
}

function Get-MatchedCrashEvents {
    param(
        [object[]]$CrashEvents,
        [string[]]$Patterns
    )

    return @(
        $CrashEvents | Where-Object {
            Test-AnyMatch -Patterns $Patterns -Value ($_.AppName + ' ' + $_.AppPath + ' ' + $_.FaultingModule + ' ' + $_.FaultingModulePath)
        }
    )
}

function Convert-ScoreToConfidence {
    param(
        [int]$Score
    )

    if ($Score -ge 8) {
        return 'high'
    }

    if ($Score -ge 4) {
        return 'medium'
    }

    return 'low'
}

function Get-Suspects {
    param(
        [object[]]$Profiles,
        [object[]]$Programs,
        [object[]]$Services,
        [object[]]$Processes,
        [object[]]$StartupItems,
        [object[]]$CrashEvents,
        [string]$WinsockCatalog,
        [string]$DriverCatalog
    )

    $suspects = foreach ($profile in $Profiles) {
        $matchedPrograms = Get-MatchedPrograms -Programs $Programs -Patterns $profile.ProgramPatterns
        $matchedServices = Get-MatchedServices -Services $Services -Patterns $profile.ServicePatterns
        $matchedProcesses = Get-MatchedProcesses -Processes $Processes -Patterns $profile.ProcessPatterns
        $matchedStartup = Get-MatchedStartupItems -StartupItems $StartupItems -Patterns $profile.StartupPatterns
        $matchedCrashes = Get-MatchedCrashEvents -CrashEvents $CrashEvents -Patterns $profile.CrashPatterns
        $winsockHit = if (@($profile.WinsockPatterns).Count -gt 0) { Test-AnyMatch -Patterns $profile.WinsockPatterns -Value $WinsockCatalog } else { $false }
        $driverHit = if (@($profile.DriverPatterns).Count -gt 0) { Test-AnyMatch -Patterns $profile.DriverPatterns -Value $DriverCatalog } else { $false }

        $score = 0
        if (@($matchedPrograms).Count -gt 0) { $score += 2 }
        if (@($matchedServices).Count -gt 0) { $score += 2 }
        if (@($matchedProcesses).Count -gt 0) { $score += 1 }
        if (@($matchedStartup).Count -gt 0) { $score += 1 }
        if ($winsockHit) { $score += 4 }
        if ($driverHit) { $score += 3 }
        if (@($matchedCrashes).Count -ge 3) {
            $score += 3
        }
        elseif (@($matchedCrashes).Count -gt 0) {
            $score += 1
        }

        if ($score -le 0) {
            continue
        }

        $evidence = @()
        foreach ($item in $matchedPrograms | Select-Object -First 5) {
            $evidence += "Installed program: $($item.DisplayName) $($item.DisplayVersion)"
        }
        foreach ($item in $matchedServices | Select-Object -First 5) {
            $evidence += "Service: $($item.Name) [$($item.Status)]"
        }
        foreach ($item in $matchedProcesses | Select-Object -First 5) {
            $evidence += "Process: $($item.ProcessName)"
        }
        foreach ($item in $matchedStartup | Select-Object -First 5) {
            $evidence += "Startup item: $($item.Name)"
        }
        foreach ($item in $matchedCrashes | Select-Object -First 5) {
            $evidence += "Crash evidence: $($item.AppName) -> $($item.FaultingModule)"
        }
        if ($winsockHit) {
            $evidence += 'Winsock catalog contains matching provider entries.'
        }
        if ($driverHit) {
            $evidence += 'Driver catalog contains matching driver entries.'
        }

        [pscustomobject]@{
            Id               = $profile.Id
            Name             = $profile.Name
            Description      = $profile.Description
            Score            = $score
            Confidence       = Convert-ScoreToConfidence -Score $score
            RiskLevel        = $profile.RiskLevel
            SupportsApplyFix = $profile.SupportsApplyFix
            Evidence         = $evidence
            SuggestedActions = $profile.SuggestedActions
        }
    }

    return @($suspects | Sort-Object Score -Descending)
}

function Get-SystemAdvice {
    param(
        [object[]]$CrashAppSummary,
        [object[]]$CrashModuleSummary
    )

    $advice = New-Object System.Collections.Generic.List[string]

    if (@($CrashAppSummary).Count -gt 1) {
        $advice.Add('Multiple apps have crashed recently. Treat this as a shared environment issue first, not as a single broken app.')
    }

    $runtimeModules = @('MSVCP140.dll', 'VCRUNTIME140.dll', 'ucrtbase.dll', 'KERNELBASE.dll')
    $runtimeHits = @(
        $CrashModuleSummary | Where-Object { $runtimeModules -contains $_.FaultingModule }
    )

    if (@($runtimeHits).Count -gt 0) {
        $advice.Add('Common runtime modules are showing up in crash logs. Reinstall the VC++ 2015-2022 redistributables and run DISM/SFC from an elevated console.')
    }

    if (@($CrashAppSummary).Count -eq 0) {
        $advice.Add('No Application Error 1000 events were found in the selected window. Increase -Days if you want a longer lookback.')
    }

    $advice.Add('If the likely culprit is a VPN, security stack, input method, overlay, or OEM background service, remove one suspect at a time and reboot between changes.')
    $advice.Add('If the system still crashes after the suspect components are removed, run a clean boot and then retest.')

    return @($advice)
}

function Convert-ToMarkdownReport {
    param(
        [pscustomobject]$Report
    )

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('# App Crash Doctor Report')
    $lines.Add('')
    $lines.Add("Generated: $($Report.System.CollectedAtLocal)")
    $lines.Add("Computer: $($Report.System.ComputerName)")
    $lines.Add("Windows: $($Report.System.WindowsProduct) ($($Report.System.WindowsVersion), build $($Report.System.OsBuildNumber))")
    $lines.Add('')

    $lines.Add('## Recent Crash Summary')
    if (@($Report.CrashAppSummary).Count -eq 0) {
        $lines.Add('- No crash events found in the selected lookback window.')
    }
    else {
        foreach ($item in $Report.CrashAppSummary | Select-Object -First 10) {
            $lines.Add("- $($item.AppName): $($item.Count) event(s), latest $($item.LatestSeen)")
        }
    }
    $lines.Add('')

    $lines.Add('## Common Faulting Modules')
    if (@($Report.CrashModuleSummary).Count -eq 0) {
        $lines.Add('- No faulting modules found.')
    }
    else {
        foreach ($item in $Report.CrashModuleSummary | Select-Object -First 10) {
            $lines.Add("- $($item.FaultingModule): $($item.Count)")
        }
    }
    $lines.Add('')

    $lines.Add('## Suspects')
    if (@($Report.Suspects).Count -eq 0) {
        $lines.Add('- No suspect profile crossed the scoring threshold.')
    }
    else {
        foreach ($suspect in $Report.Suspects) {
            $lines.Add("- $($suspect.Name) [$($suspect.Confidence)] score=$($suspect.Score)")
            foreach ($evidenceLine in $suspect.Evidence | Select-Object -First 5) {
                $lines.Add("  - $evidenceLine")
            }
        }
    }
    $lines.Add('')

    $lines.Add('## Suggested Actions')
    foreach ($action in $Report.SystemAdvice) {
        $lines.Add("- $action")
    }
    foreach ($suspect in $Report.Suspects) {
        foreach ($action in $suspect.SuggestedActions) {
            $lines.Add("- [$($suspect.Id)] $action")
        }
    }
    $lines.Add('')

    if (@($Report.RepairResults).Count -gt 0) {
        $lines.Add('## Repair Results')
        foreach ($result in $Report.RepairResults) {
            $lines.Add("- $($result.Step): $($result.Result)")
        }
        $lines.Add('')
    }

    return ($lines -join [Environment]::NewLine)
}

function Save-ReportFiles {
    param(
        [pscustomobject]$Report,
        [string]$MarkdownPath,
        [string]$StructuredJsonPath
    )

    if (-not [string]::IsNullOrWhiteSpace($MarkdownPath)) {
        $folder = Split-Path -Path $MarkdownPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($folder)) {
            New-Item -ItemType Directory -Path $folder -Force | Out-Null
        }
        Convert-ToMarkdownReport -Report $Report | Set-Content -Path $MarkdownPath -Encoding UTF8
    }

    if (-not [string]::IsNullOrWhiteSpace($StructuredJsonPath)) {
        $folder = Split-Path -Path $StructuredJsonPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($folder)) {
            New-Item -ItemType Directory -Path $folder -Force | Out-Null
        }
        $Report | ConvertTo-Json -Depth 8 | Set-Content -Path $StructuredJsonPath -Encoding UTF8
    }
}

function Invoke-ProcessCommand {
    param(
        [string]$FilePath,
        [string]$ArgumentList,
        [string]$Label,
        [switch]$Simulate
    )

    if ($Simulate) {
        return [pscustomobject]@{
            Step   = $Label
            Result = "WhatIf: `"$FilePath`" $ArgumentList"
        }
    }

    if (-not (Test-Path -LiteralPath $FilePath)) {
        return [pscustomobject]@{
            Step   = $Label
            Result = 'Skipped: file not found'
        }
    }

    try {
        $process = Start-Process -FilePath $FilePath -ArgumentList $ArgumentList -PassThru -Wait -WindowStyle Hidden
        return [pscustomobject]@{
            Step   = $Label
            Result = "ExitCode=$($process.ExitCode)"
        }
    }
    catch {
        return [pscustomobject]@{
            Step   = $Label
            Result = $_.Exception.Message
        }
    }
}

function Invoke-SangforFix {
    param(
        [object[]]$Programs,
        [switch]$DoResetWinsock,
        [switch]$Simulate
    )

    $results = New-Object System.Collections.Generic.List[object]
    $processNames = @('ECAgent', 'EasyConnect', 'SangforPromoteService', 'SvpnJobber')

    foreach ($name in $processNames) {
        $matches = @(Get-Process -Name $name -ErrorAction SilentlyContinue)
        foreach ($match in $matches) {
            if ($Simulate) {
                $results.Add([pscustomobject]@{ Step = "Stop process $($match.ProcessName)"; Result = 'WhatIf' })
                continue
            }

            try {
                Stop-Process -Id $match.Id -Force -ErrorAction Stop
                $results.Add([pscustomobject]@{ Step = "Stop process $($match.ProcessName)"; Result = 'OK' })
            }
            catch {
                $results.Add([pscustomobject]@{ Step = "Stop process $($match.ProcessName)"; Result = $_.Exception.Message })
            }
        }
    }

    $serviceNames = @('SangforSP', 'SangforVnic')
    foreach ($serviceName in $serviceNames) {
        $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
        if (-not $service) {
            continue
        }

        if ($Simulate) {
            $results.Add([pscustomobject]@{ Step = "Stop service $serviceName"; Result = 'WhatIf' })
            continue
        }

        try {
            Stop-Service -Name $serviceName -Force -ErrorAction Stop
            $results.Add([pscustomobject]@{ Step = "Stop service $serviceName"; Result = 'OK' })
        }
        catch {
            $results.Add([pscustomobject]@{ Step = "Stop service $serviceName"; Result = $_.Exception.Message })
        }
    }

    $commands = New-Object System.Collections.Generic.List[object]
    foreach ($program in $Programs | Where-Object { Test-AnyMatch -Patterns @('Sangfor', '深信服', 'EasyConnect') -Value ($_.DisplayName + ' ' + $_.Publisher) }) {
        $rawCommand = if (-not [string]::IsNullOrWhiteSpace($program.QuietUninstallString)) { $program.QuietUninstallString } else { $program.UninstallString }
        $resolved = Resolve-CommandLine -CommandLine $rawCommand
        if (-not $resolved) {
            continue
        }

        $argumentList = $resolved.ArgumentList
        if ($argumentList -notmatch '(^| )/(S|silent|quiet)\b' -and $argumentList -notmatch '(^| )-(s|silent|quiet)\b') {
            $argumentList = ($argumentList + ' /S').Trim()
        }

        $commands.Add([pscustomobject]@{
                Label        = "Uninstall $($program.DisplayName)"
                FilePath     = $resolved.FilePath
                ArgumentList = $argumentList
            })
    }

    $extraPaths = @(
        'C:\Program Files (x86)\Sangfor\SSL\ClientComponent\Uninstall.exe',
        'C:\Program Files (x86)\Sangfor\SSL\ClientComponent2\Uninstall.exe',
        'C:\Program Files (x86)\Sangfor\SSL\EasyConnect\EasyConnectUninstaller.exe',
        'C:\Program Files (x86)\Sangfor\SSL\DnsDriver\UnDnsDriverInstaller.exe',
        'C:\Program Files (x86)\Sangfor\SSL\DnsDriver1\UnDnsDriverInstaller.exe',
        'C:\Program Files (x86)\Sangfor\SSL\TcpDriver\TcpDriverUnInstaller.exe',
        'C:\Program Files (x86)\Sangfor\SSL\TcpDriver1\TcpDriverUnInstaller.exe',
        'C:\Program Files (x86)\Sangfor\SSL\Promote\PromoteServiceUninstall.exe',
        'C:\Program Files (x86)\Sangfor\SSL\Promote\PromoteUninstall.exe',
        'C:\Program Files (x86)\Sangfor\SSL\RemoteAppClient\Uninstaller.exe',
        'C:\Program Files (x86)\Sangfor\SSL\SangforServiceClient\SangforServiceClientUninstaller.exe',
        'C:\Program Files (x86)\Sangfor\SSL\SangforCSClient\SangforCSClientUninstaller.exe',
        'C:\Program Files (x86)\Sangfor\SSL\CSClient\VNIC\uninstall.exe',
        'C:\Program Files (x86)\Sangfor\SSL\UBDllS\DataParserUninstaller.exe'
    )

    foreach ($path in $extraPaths) {
        if (-not (Test-Path -LiteralPath $path)) {
            continue
        }

        $existingCommandPaths = @($commands | ForEach-Object { $_.FilePath })
        if ($existingCommandPaths -contains $path) {
            continue
        }

        $commands.Add([pscustomobject]@{
                Label        = "Run vendor uninstaller $path"
                FilePath     = $path
                ArgumentList = '/S'
            })
    }

    $seen = @{}
    foreach ($command in $commands) {
        $key = "$($command.FilePath)|$($command.ArgumentList)"
        if ($seen.ContainsKey($key)) {
            continue
        }
        $seen[$key] = $true
        $results.Add((Invoke-ProcessCommand -FilePath $command.FilePath -ArgumentList $command.ArgumentList -Label $command.Label -Simulate:$Simulate))
    }

    $driverCatalog = Get-DriverCatalogText
    $driverMatches = [regex]::Matches($driverCatalog, 'Published Name\s*:\s*(oem\d+\.inf)[\s\S]*?(Original Name|Provider Name)\s*:\s*([^\r\n]+)', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    foreach ($driverMatch in $driverMatches) {
        $publishedName = $driverMatch.Groups[1].Value
        $descriptor = $driverMatch.Groups[3].Value
        if ($descriptor -notmatch 'Sangfor' -and $descriptor -notmatch 'sangforvnic') {
            continue
        }

        $results.Add((Invoke-ProcessCommand -FilePath 'pnputil.exe' -ArgumentList "/delete-driver $publishedName /uninstall /force" -Label "Remove driver $publishedName" -Simulate:$Simulate))
    }

    $results.Add((Invoke-ProcessCommand -FilePath 'sc.exe' -ArgumentList 'delete SangforVnic' -Label 'Delete service SangforVnic' -Simulate:$Simulate))

    if ($DoResetWinsock) {
        $results.Add((Invoke-ProcessCommand -FilePath 'netsh.exe' -ArgumentList 'winsock reset' -Label 'Reset Winsock' -Simulate:$Simulate))
    }
    else {
        $results.Add([pscustomobject]@{
                Step   = 'Reset Winsock'
                Result = 'Skipped: pass -ResetWinsock to include this step'
            })
    }

    $rootPath = 'C:\Program Files (x86)\Sangfor'
    if ($Simulate) {
        $results.Add([pscustomobject]@{
                Step   = "Remove folder $rootPath"
                Result = 'WhatIf'
            })
    }
    elseif (Test-Path -LiteralPath $rootPath) {
        try {
            Remove-Item -LiteralPath $rootPath -Recurse -Force -ErrorAction Stop
            $results.Add([pscustomobject]@{
                    Step   = "Remove folder $rootPath"
                    Result = 'OK'
                })
        }
        catch {
            $results.Add([pscustomobject]@{
                    Step   = "Remove folder $rootPath"
                    Result = $_.Exception.Message
                })
        }
    }

    return @($results.ToArray())
}

function Invoke-RepairActions {
    param(
        [string[]]$RequestedTargets,
        [object[]]$Programs,
        [switch]$DoResetWinsock,
        [switch]$Simulate
    )

    $results = New-Object System.Collections.Generic.List[object]

    foreach ($requestedTarget in $RequestedTargets) {
        switch ($requestedTarget) {
            'Sangfor' {
                foreach ($result in Invoke-SangforFix -Programs $Programs -DoResetWinsock:$DoResetWinsock -Simulate:$Simulate) {
                    $results.Add($result)
                }
            }
            default {
                $results.Add([pscustomobject]@{
                        Step   = "Target $requestedTarget"
                        Result = 'No apply-fix implementation exists for this target.'
                    })
            }
        }
    }

    return @($results.ToArray())
}

$system = Get-SystemSummary
$programs = @(Get-InstalledPrograms)
$services = @(Get-ServiceSnapshot)
$processes = @(Get-ProcessSnapshot)
$startupItems = @(Get-StartupSnapshot)
$crashEvents = @(Get-RecentCrashEvents -RecentDays $Days -Limit $MaxEvents)
$crashAppSummary = @(Get-CrashAppSummary -CrashEvents $crashEvents)
$crashModuleSummary = @(Get-CrashModuleSummary -CrashEvents $crashEvents)
$winsockCatalog = Get-WinsockCatalogText
$driverCatalog = Get-DriverCatalogText
$profiles = @(Get-RepairProfiles)
$suspects = @(Get-Suspects `
    -Profiles $profiles `
    -Programs $programs `
    -Services $services `
    -Processes $processes `
    -StartupItems $startupItems `
    -CrashEvents $crashEvents `
    -WinsockCatalog $winsockCatalog `
    -DriverCatalog $driverCatalog)
$systemAdvice = @(Get-SystemAdvice -CrashAppSummary $crashAppSummary -CrashModuleSummary $crashModuleSummary)

$repairResults = @()
if ($Mode -eq 'apply-fix') {
    if (@($Target).Count -eq 0) {
        throw 'apply-fix mode requires at least one -Target value, for example -Target Sangfor.'
    }

    $applyableTargets = @(
        $suspects |
            Where-Object { $_.SupportsApplyFix } |
            ForEach-Object { $_.Id }
    )

    foreach ($requestedTarget in $Target) {
        if ($applyableTargets -notcontains $requestedTarget) {
            Write-Warning "Target '$requestedTarget' is not currently detected as an active high-confidence apply-fix suspect."
        }
    }

    $repairResults = Invoke-RepairActions -RequestedTargets $Target -Programs $programs -DoResetWinsock:$ResetWinsock -Simulate:$WhatIf
}

$report = [pscustomobject]@{
    System             = $system
    Mode               = $Mode
    LookbackDays       = $Days
    CrashEventCount    = @($crashEvents).Count
    CrashAppSummary    = @($crashAppSummary)
    CrashModuleSummary = @($crashModuleSummary)
    Suspects           = @($suspects)
    SystemAdvice       = @($systemAdvice)
    RepairResults      = @($repairResults)
}

Save-ReportFiles -Report $report -MarkdownPath $ReportPath -StructuredJsonPath $JsonPath

switch ($Mode) {
    'scan' {
        Write-Host ''
        Write-Host '== App Crash Doctor Scan =='
        Write-Host "Machine: $($system.ComputerName)  Windows: $($system.WindowsProduct) build $($system.OsBuildNumber)"
        Write-Host "Crash events found: $($report.CrashEventCount)"
        Write-Host ''
        Write-Host 'Top crashed apps:'
        if (@($crashAppSummary).Count -eq 0) {
            Write-Host '  none'
        }
        else {
            foreach ($item in $crashAppSummary | Select-Object -First 8) {
                Write-Host "  $($item.AppName) -> $($item.Count)"
            }
        }
        Write-Host ''
        Write-Host 'Suspects:'
        if (@($suspects).Count -eq 0) {
            Write-Host '  none'
        }
        else {
            foreach ($suspect in $suspects) {
                Write-Host "  $($suspect.Id) [$($suspect.Confidence)] score=$($suspect.Score)"
            }
        }
    }
    'suggest-fix' {
        Write-Host ''
        Write-Host '== App Crash Doctor Suggestions =='
        foreach ($line in $systemAdvice) {
            Write-Host "  - $line"
        }
        foreach ($suspect in $suspects) {
            Write-Host ''
            Write-Host "[$($suspect.Id)] $($suspect.Name) [$($suspect.Confidence)]"
            foreach ($line in $suspect.Evidence | Select-Object -First 5) {
                Write-Host "  evidence: $line"
            }
            foreach ($line in $suspect.SuggestedActions) {
                Write-Host "  action: $line"
            }
        }
    }
    'apply-fix' {
        Write-Host ''
        Write-Host '== App Crash Doctor Apply Fix =='
        foreach ($item in $repairResults) {
            Write-Host "  $($item.Step): $($item.Result)"
        }
    }
}

if ($PassThru) {
    $report
}
