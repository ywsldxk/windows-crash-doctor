[CmdletBinding(DefaultParameterSetName = 'Forward')]
param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$RemainingArgs
)

$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$newScriptPath = Join-Path $scriptDirectory 'WindowsCrashDoctor.ps1'
$engine = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh' } else { 'powershell' }

& $engine -ExecutionPolicy Bypass -File $newScriptPath @RemainingArgs
exit $LASTEXITCODE
