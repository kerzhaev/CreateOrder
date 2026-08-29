[CmdletBinding()]
param(
    [string]$WorkbookPath,
    [string]$SourceDirectory,
    [string]$TargetComponentName = 'frmPersonnelActionWizardV2'
)

if ([string]::IsNullOrWhiteSpace($WorkbookPath)) { $WorkbookPath = Join-Path $PSScriptRoot '..\CreateOrder.xlsm' }
if ([string]::IsNullOrWhiteSpace($SourceDirectory)) { $SourceDirectory = Join-Path $PSScriptRoot '..\CreateOrder.xlsm.modules' }

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Verbose ('Exporting owner-edited personnel action designer form: {0}' -f $TargetComponentName)
& (Join-Path $PSScriptRoot 'Export-EnrollmentWizardV2Designer.ps1') `
    -WorkbookPath $WorkbookPath `
    -SourceDirectory $SourceDirectory `
    -TargetComponentName $TargetComponentName `
    -OperationName 'Export-PersonnelActionWizardV2Designer' `
    -BackupPrefix 'personnel-action-v2-owner-layout' `
    -Verbose:($VerbosePreference -eq 'Continue')
