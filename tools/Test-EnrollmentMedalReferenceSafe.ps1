param(
    [string]$WorkbookPath = (Join-Path $PSScriptRoot "..\CreateOrder.xlsm")
)

$ErrorActionPreference = "Stop"
$workspace = Split-Path -Parent $PSScriptRoot
$testDirectory = Join-Path $workspace "_tmp_enrollment_medal_reference_test"
$testWorkbookPath = Join-Path $testDirectory "CreateOrder_enrollment_medal_reference_test.xlsm"

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
    Import-CodeModuleText $workbook "mdlEnrollmentWorkflow" (Join-Path $workspace "CreateOrder.xlsm.modules\mdlEnrollmentWorkflow.bas")
    $formPath = Join-Path $workspace "CreateOrder.xlsm.modules\frmEnrollmentWizard.frm"
    $components = $workbook.VBProject.VBComponents
    try { $components.Remove($components.Item("frmEnrollmentWizard")) } catch {}
    $components.Import($formPath) | Out-Null

    try { $workbook.VBProject.VBComponents.Remove($workbook.VBProject.VBComponents.Item("enr_medal_probe")) } catch {}
    $probe = $workbook.VBProject.VBComponents.Add(1)
    $probe.Name = "enr_medal_probe"
    $probe.CodeModule.AddFromString(@"
Option Explicit
Public Function ProbeEnrollmentMedalReferences() As String
    Dim values As Collection
    Dim i As Long
    Dim totalAmount As Long
    Load frmEnrollmentWizard
    Unload frmEnrollmentWizard
    mdlEnrollmentWorkflow.EnsureEnrollmentReferenceData
    Set values = mdlEnrollmentWorkflow.GetEnrollmentReferenceValues("ACHIEVEMENT")
    If values.Count <> 4 Then
        ProbeEnrollmentMedalReferences = "FAILED: expected four medal references"
        Exit Function
    End If
    For i = 1 To values.Count
        totalAmount = totalAmount + CLng(mdlEnrollmentWorkflow.GetEnrollmentReferenceAmount("ACHIEVEMENT", CStr(values(i))))
        If mdlEnrollmentWorkflow.GetEnrollmentReferenceCode("ACHIEVEMENT", CStr(values(i))) = "" Then
            ProbeEnrollmentMedalReferences = "FAILED: medal reference code is missing"
            Exit Function
        End If
    Next i
    If totalAmount <> 80 Then
        ProbeEnrollmentMedalReferences = "FAILED: medal percentages must be 30, 20, 20 and 10"
        Exit Function
    End If
    ProbeEnrollmentMedalReferences = "OK"
End Function
"@)
    $result = $excel.Run("'$($workbook.Name)'!ProbeEnrollmentMedalReferences")
    if ($result -ne "OK") { throw $result }
    $workbook.Close($false)
    $workbook = $null
    $excel.Quit()
    $excel = $null
    Write-Output "Enrollment medal reference safe acceptance passed."
}
finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
