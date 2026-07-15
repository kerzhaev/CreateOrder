param(
    [string]$WorkbookPath = (Join-Path $PSScriptRoot "..\CreateOrder.xlsm")
)

$ErrorActionPreference = "Stop"
$workspace = Split-Path -Parent $PSScriptRoot
$testDirectory = Join-Path $workspace "_tmp_enrollment_compact_ui_test"
$testWorkbookPath = Join-Path $testDirectory "CreateOrder_enrollment_compact_ui_test.xlsm"

function Read-VbaText([string]$Path) {
    [IO.File]::ReadAllText($Path, [Text.Encoding]::GetEncoding(1251))
}

function Import-CodeModuleText($Workbook, [string]$ModuleName, [string]$ModulePath) {
    $code = Read-VbaText $ModulePath
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
    Import-CodeModuleText $workbook "mdlEnrollmentOrderExport" (Join-Path $workspace "CreateOrder.xlsm.modules\mdlEnrollmentOrderExport.bas")
    try { $workbook.VBProject.VBComponents.Remove($workbook.VBProject.VBComponents.Item("frmEnrollmentWizard")) } catch {}
    $form = $workbook.VBProject.VBComponents.Import((Join-Path $workspace "CreateOrder.xlsm.modules\frmEnrollmentWizard.frm"))
    if ($form.Type -ne 3) { throw "Enrollment form was imported as component type $($form.Type), expected 3." }

    try { $workbook.VBProject.VBComponents.Remove($workbook.VBProject.VBComponents.Item("enrollment_compact_ui_probe")) } catch {}
    $probe = $workbook.VBProject.VBComponents.Add(1)
    $probe.Name = "enrollment_compact_ui_probe"
    $probe.CodeModule.AddFromString(@"
Option Explicit
Public Function ProbeEnrollmentCompactUi() As String
    Dim pageIndex As Long
    Dim wsReferences As Worksheet
    Dim frameHost As Object
    Dim exportResult As String
    Dim referenceRows As Long
    Dim startedAt As Double
    Dim rowNum As Long
    Dim referenceType As String
    Dim controlItem As Object
    Dim hasPersonalToggle As Boolean
    Dim hasBankToggle As Boolean
    Dim tariffDefinition As Object
    Dim secrecyDefinition As Object
    On Error GoTo Failed
    startedAt = Timer
    mdlEnrollmentWorkflow.EnsureEnrollmentInfrastructure
    mdlEnrollmentWorkflow.EnsureEnrollmentInfrastructure
    Set wsReferences = ThisWorkbook.Worksheets("EnrollmentReferenceData")
    referenceRows = wsReferences.Cells(wsReferences.Rows.Count, 1).End(xlUp).Row - 1
    If referenceRows > 400 Then Err.Raise 803, , "Enrollment reference data was not deduplicated"
    For rowNum = 2 To wsReferences.Cells(wsReferences.Rows.Count, 1).End(xlUp).Row
        referenceType = UCase`$(CStr(wsReferences.Cells(rowNum, 1).Value))
        If referenceType = "POSITION" Or referenceType = "VUS" Or referenceType = "SECTION" Or referenceType = "MILITARY_UNIT" Then _
            Err.Raise 807, , "Staff data remained in EnrollmentReferenceData: " & referenceType
    Next rowNum
    Set tariffDefinition = mdlEnrollmentWorkflow.GetEnrollmentPaymentDefinition("std_tariff")
    Set secrecyDefinition = mdlEnrollmentWorkflow.GetEnrollmentPaymentDefinition("secrecy")
    If tariffDefinition("word_block_target") <> "Section1MonthlyPersonal" Then Err.Raise 810, , "1-4 tariff allowance is not assigned to order 430"
    If secrecyDefinition("word_block_target") <> "Section1MonthlyStandard" Then Err.Raise 811, , "Secrecy allowance is not assigned to order 727"
    exportResult = mdlEnrollmentOrderExport.ExportEnrollmentOrderByDraftId("", 0)
    If Left`$(exportResult, 6) <> "ERROR:" Then Err.Raise 806, , "Empty export did not return the expected safe error"
    Load frmEnrollmentWizard
    If frmEnrollmentWizard.Controls("mpWizard").Pages.Count <> 7 Then Err.Raise 801, , "Unexpected page count"
    Set frameHost = frmEnrollmentWizard.Controls("mpWizard").Pages(2).Controls("fraOrder727")
    If frameHost Is Nothing Then Err.Raise 804, , "Order 727 frame is missing"
    Set frameHost = frmEnrollmentWizard.Controls("mpWizard").Pages(2).Controls("fraOrder430")
    If frameHost Is Nothing Then Err.Raise 805, , "Order 430 frame is missing"
    Set frameHost = frmEnrollmentWizard.Controls("mpWizard").Pages(1).Controls("fraDocsArrival")
    If frameHost Is Nothing Then Err.Raise 812, , "Arrival-details frame is missing"
    Set frameHost = frmEnrollmentWizard.Controls("mpWizard").Pages(1).Controls("fraDocsReport")
    If frameHost Is Nothing Then Err.Raise 813, , "Report-details frame is missing"
    For Each controlItem In frmEnrollmentWizard.Controls("mpWizard").Pages(3).Controls
        If controlItem.Top = 138 Then
            If controlItem.Left = 12 Then hasPersonalToggle = True
            If controlItem.Left = 550 Then hasBankToggle = True
        End If
    Next controlItem
    If Not hasPersonalToggle Or Not hasBankToggle Then Err.Raise 808, , "Optional personal or bank data toggles are missing"
    For pageIndex = 0 To frmEnrollmentWizard.Controls("mpWizard").Pages.Count - 1
        If frmEnrollmentWizard.Controls("mpWizard").Pages(pageIndex).ScrollBars <> 0 Then Err.Raise 802, , "Vertical scrolling remains enabled"
    Next pageIndex
    Unload frmEnrollmentWizard
    ProbeEnrollmentCompactUi = "OK|" & Format`$(Timer - startedAt, "0.000")
    Exit Function
Failed:
    ProbeEnrollmentCompactUi = "FAILED: " & Err.Description
End Function
"@)
    $result = $excel.Run("'$($workbook.Name)'!ProbeEnrollmentCompactUi")
    if (-not $result.StartsWith("OK|")) { throw $result }
    $workbook.Close($false); $workbook = $null
    $excel.Quit(); $excel = $null
    Write-Output "Enrollment compact UI safe acceptance passed. VBA preparation and form load: $($result.Split('|')[1]) sec."
}
finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
