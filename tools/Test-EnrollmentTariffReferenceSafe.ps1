param(
    [string]$WorkbookPath = (Join-Path $PSScriptRoot "..\CreateOrder.xlsm")
)

$ErrorActionPreference = "Stop"
$workspace = Split-Path -Parent $PSScriptRoot
$testDirectory = Join-Path $workspace "_tmp_enrollment_tariff_reference_test"
$testWorkbookPath = Join-Path $testDirectory "CreateOrder_enrollment_tariff_reference_test.xlsm"

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
    try { $workbook.VBProject.VBComponents.Remove($workbook.VBProject.VBComponents.Item("enr_tariff_probe")) } catch {}
    $probe = $workbook.VBProject.VBComponents.Add(1)
    $probe.Name = "enr_tariff_probe"
    $probe.CodeModule.AddFromString(@"
Option Explicit
Public Function ProbeEnrollmentTariffReference() As String
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim displayValue As String
    Dim rankCode As String
    Dim rankDisplayValue As String
    mdlEnrollmentWorkflow.EnsureEnrollmentReferenceData
    Set ws = ThisWorkbook.Worksheets("EnrollmentReferenceData")
    For rowNum = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(rowNum, 1).Value) = "TARIFF_RANK" And CStr(ws.Cells(rowNum, 2).Value) = "1" Then
            displayValue = CStr(ws.Cells(rowNum, 3).Value)
            ws.Cells(rowNum, 4).Value = "12345"
            Exit For
        End If
    Next rowNum
    If displayValue = "" Then
        ProbeEnrollmentTariffReference = "FAILED: tariff reference is missing"
    ElseIf mdlEnrollmentWorkflow.GetTariffRankReferenceAmount(displayValue) <> "12345" Then
        ProbeEnrollmentTariffReference = "FAILED: display name did not resolve the tariff salary"
    ElseIf mdlEnrollmentWorkflow.GetTariffRankReferenceAmount("1") <> "12345" Then
        ProbeEnrollmentTariffReference = "FAILED: tariff code did not resolve the tariff salary"
    Else
        For rowNum = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If CStr(ws.Cells(rowNum, 1).Value) = "RANK" Then
                rankCode = CStr(ws.Cells(rowNum, 2).Value)
                rankDisplayValue = CStr(ws.Cells(rowNum, 3).Value)
                ws.Cells(rowNum, 4).Value = "5000"
                Exit For
            End If
        Next rowNum
        If rankCode = "" Or rankDisplayValue = rankCode Then
            ProbeEnrollmentTariffReference = "FAILED: rank display was not shortened"
        ElseIf mdlEnrollmentWorkflow.GetEnrollmentReferenceCodeOrDisplay("RANK", rankDisplayValue) <> rankCode Then
            ProbeEnrollmentTariffReference = "FAILED: rank display did not return the full rank code"
        ElseIf mdlEnrollmentWorkflow.GetEnrollmentReferenceDisplayNameOrCode("RANK", rankCode) <> rankDisplayValue Then
            ProbeEnrollmentTariffReference = "FAILED: full rank code did not return the short display value"
        ElseIf mdlEnrollmentWorkflow.GetRankReferenceAmount(rankCode) <> "5000" Then
            ProbeEnrollmentTariffReference = "FAILED: full rank code did not resolve the rank salary"
        Else
            ProbeEnrollmentTariffReference = "OK"
        End If
    End If
End Function
"@)
    $result = $excel.Run("'$($workbook.Name)'!ProbeEnrollmentTariffReference")
    if ($result -ne "OK") { throw $result }
    $workbook.Close($false)
    $workbook = $null
    $excel.Quit()
    $excel = $null
    Write-Output "Enrollment tariff reference safe acceptance passed."
}
finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
