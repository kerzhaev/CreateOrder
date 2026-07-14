param(
    [string]$WorkbookPath = (Join-Path $PSScriptRoot "..\CreateOrder.xlsm")
)

$ErrorActionPreference = "Stop"
$workspace = Split-Path -Parent $PSScriptRoot
$testDirectory = Join-Path $workspace "_tmp_enrollment_fizo_reference_test"
$testWorkbookPath = Join-Path $testDirectory "CreateOrder_enrollment_fizo_reference_test.xlsm"
$modulePath = Join-Path $workspace "CreateOrder.xlsm.modules\mdlEnrollmentWorkflow.bas"

function Import-CodeModuleText($Workbook, [string]$ModuleName, [string]$Path) {
    $code = [IO.File]::ReadAllText($Path, [Text.Encoding]::GetEncoding(1251))
    $code = [regex]::Replace($code, '^Attribute VB_Name\s*=\s*"[^"]+"\r?\n', '', 1)
    $module = $Workbook.VBProject.VBComponents.Item($ModuleName).CodeModule
    if ($module.CountOfLines -gt 0) { $module.DeleteLines(1, $module.CountOfLines) }
    $module.AddFromString($code)
}

New-Item -ItemType Directory -Path $testDirectory -Force | Out-Null
Copy-Item -LiteralPath $WorkbookPath -Destination $testWorkbookPath -Force
$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try { $excel.AutomationSecurity = 1 } catch {}
    $workbook = $excel.Workbooks.Open($testWorkbookPath, 0, $false)
    Import-CodeModuleText $workbook "mdlEnrollmentWorkflow" $modulePath

    try { $workbook.VBProject.VBComponents.Remove($workbook.VBProject.VBComponents.Item("enrollment_fizo_reference_probe")) } catch {}
    $probe = $workbook.VBProject.VBComponents.Add(1)
    $probe.Name = "enrollment_fizo_reference_probe"
    $probe.CodeModule.AddFromString(@"
Option Explicit
Public Function ProbeEnrollmentFizoReference() As String
    Dim values As Collection
    mdlEnrollmentWorkflow.EnsureEnrollmentReferenceData
    Set values = mdlEnrollmentWorkflow.GetEnrollmentReferenceValues("FIZO")
    If values.Count <> 1 Then
        ProbeEnrollmentFizoReference = "FAILED: FIZO reference must contain the configured level"
        Exit Function
    End If
    If mdlEnrollmentWorkflow.GetEnrollmentReferenceAmount("FIZO", CStr(values(1))) <> "15" Then
        ProbeEnrollmentFizoReference = "FAILED: FIZO level 2 must resolve to 15%"
        Exit Function
    End If
    ProbeEnrollmentFizoReference = "OK"
End Function
"@)
    $result = $excel.Run("'$($workbook.Name)'!ProbeEnrollmentFizoReference")
    if ($result -ne "OK") { throw $result }
    $workbook.Close($false)
    $workbook = $null
    $excel.Quit()
    $excel = $null
    Write-Output "Enrollment FIZO reference safe acceptance passed."
}
finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
