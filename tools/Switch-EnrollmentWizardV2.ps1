[CmdletBinding()]
param(
    [string]$ProjectRoot
)

if ([string]::IsNullOrWhiteSpace($ProjectRoot)) { $ProjectRoot = Join-Path $PSScriptRoot '..' }

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-SwitchLog {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('DEBUG', 'INFO', 'WARN', 'ERROR')][string]$Level,
        [Parameter(Mandatory = $true)][string]$Message,
        [hashtable]$Context = @{}
    )
    $payload = [ordered]@{
        timestamp = (Get-Date).ToString('o')
        level = $Level
        operation = 'Switch-EnrollmentWizardV2'
        message = $Message
    }
    foreach ($key in $Context.Keys) { $payload[$key] = $Context[$key] }
    $line = $payload | ConvertTo-Json -Compress -Depth 5
    if ($Level -eq 'DEBUG') { Write-Verbose $line }
    elseif ($Level -eq 'WARN') { Write-Warning $line }
    elseif ($Level -eq 'ERROR') { Write-Error $line }
    else { Write-Host $line }
}

function Replace-ExactText {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][Text.Encoding]$Encoding,
        [Parameter(Mandatory = $true)][int]$MinimumMatches
    )
    $bytes = [IO.File]::ReadAllBytes($Path)
    $text = $Encoding.GetString($bytes)
    $pattern = 'frmEnrollmentWizard(?!V2)'
    $matches = ([regex]::Matches($text, $pattern)).Count
    if ($matches -eq 0 -and $text.Contains('frmEnrollmentWizardV2')) {
        Write-SwitchLog INFO 'Target already routes to V2; no rewrite needed.' @{ path = $Path }
        return
    }
    if ($matches -lt $MinimumMatches) { throw "Expected at least $MinimumMatches V1 references in $Path; found $matches." }
    $updated = [regex]::Replace($text, $pattern, 'frmEnrollmentWizardV2')
    [IO.File]::WriteAllText($Path, $updated, $Encoding)
    Write-SwitchLog INFO 'Switched active enrollment-form references to V2.' @{ path = $Path; replacements = $matches }
}

$resolvedRoot = [IO.Path]::GetFullPath($ProjectRoot)
$windows1251 = [Text.Encoding]::GetEncoding(1251)
$utf8NoBom = [Text.UTF8Encoding]::new($false)
$targets = @(
    @{ path = (Join-Path $resolvedRoot 'CreateOrder.xlsm.modules\mdlEnrollmentWorkflow.bas'); encoding = $windows1251; minimum = 5 },
    @{ path = (Join-Path $resolvedRoot 'CreateOrder.xlsm.modules\mdlRibbonHandlers.bas'); encoding = $windows1251; minimum = 3 },
    @{ path = (Join-Path $resolvedRoot 'Test-PaymentsEnrollmentAcceptance.ps1'); encoding = $utf8NoBom; minimum = 10 }
)

foreach ($target in $targets) {
    if (-not (Test-Path -LiteralPath $target.path)) { throw "Missing switch target: $($target.path)" }
    Replace-ExactText -Path $target.path -Encoding $target.encoding -MinimumMatches $target.minimum
}

Write-SwitchLog INFO 'Enrollment Wizard V2 routing switch completed.' @{ targets = $targets.Count }
