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
    Import-CodeModuleText $workbook "mdlEnrollmentOrderExport" (Join-Path $workspace "CreateOrder.xlsm.modules\mdlEnrollmentOrderExport.bas")
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
    Dim record As Object
    Dim evaluation As Object
    Dim recordKey As Variant
    Dim tariffDefinition As Object
    Dim stageName As String
    On Error GoTo Failed
    mdlEnrollmentWorkflow.EnsureEnrollmentInfrastructure
    Set ws = ThisWorkbook.Worksheets("EnrollmentReferenceData")
    For rowNum = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(rowNum, 1).Value) = "TARIFF_RANK" And CStr(ws.Cells(rowNum, 2).Value) = "1" Then
            displayValue = CStr(ws.Cells(rowNum, 3).Value)
            If CStr(ws.Cells(rowNum, 4).Value) <> "14330" Then
                ProbeEnrollmentTariffReference = "FAILED: tariff rank 1 salary was not seeded"
                Exit Function
            End If
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
                If CStr(ws.Cells(rowNum, 4).Value) <> "7166" Then
                    ProbeEnrollmentTariffReference = "FAILED: rank salary was not seeded"
                    Exit Function
                End If
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
            Set tariffDefinition = mdlEnrollmentWorkflow.GetEnrollmentPaymentDefinition("std_tariff")
            If tariffDefinition Is Nothing Then
                ProbeEnrollmentTariffReference = "FAILED: tariff allowance definition was not created"
                Exit Function
            End If
            Set record = mdlEnrollmentWorkflow.GetBackendRecord
            record("fio") = "Тестовый военнослужащий"
            record("personal_number") = "TEST-001"
            record("rank") = rankCode
            record("position") = "Тестовая должность"
            record("section") = "Тестовое подразделение"
            record("position_salary") = "12345"
            record("rank_salary") = "5000"
            record("tariff_rank") = "1"
            record("order_number") = "1"
            record("order_date") = "15.07.2026"
            record("accept_date") = "15.07.2026"
            record("enroll_date") = "15.07.2026"
            record("duty_start_date") = "15.07.2026"
            record("basis_section1") = "Тестовое основание"
            record("std_duty_enabled") = "YES"
            record("std_duty_percent") = "100"
            record("std_tariff_enabled") = "YES"
            record("std_tariff_percent") = "50"
            record("fizo_enabled") = "YES"
            record("fizo_param") = "SECOND"
            record("fizo_percent") = "80"
            record("fizo_basis") = "Тестовое основание ФИЗО"
            record("personal_details_enabled") = "NO"
            record("bank_details_enabled") = "NO"
            record("arrival_details_enabled") = "NO"
            record("arrival_source") = "СКРЫТЬ-ПРИБЫТИЕ"
            record("prescription_number") = "СКРЫТЬ-ПРЕДПИСАНИЕ"
            record("report_details_enabled") = "NO"
            record("report_info") = "СКРЫТЬ-РАПОРТ"
            stageName = "base evaluation"
            Set evaluation = mdlEnrollmentWorkflow.EvaluateEnrollmentRecord(record)
            If evaluation("word_ready") <> "YES" Then
                ProbeEnrollmentTariffReference = "FAILED: optional empty personal or bank fields blocked the order"
                Exit Function
            End If
            If InStr(1, evaluation("preview_section1"), "СКРЫТЬ-", vbTextCompare) > 0 Then
                ProbeEnrollmentTariffReference = "FAILED: disabled arrival or report details leaked into Word preview"
                Exit Function
            End If
            If InStr(1, CStr(evaluation("preview_standard")) & CStr(evaluation("preview_personal")) & CStr(evaluation("preview_section1")), "Тестовое основание ФИЗО", vbTextCompare) = 0 Then
                ProbeEnrollmentTariffReference = "FAILED: enabled payment basis was not included in Word preview; personal=" & CStr(evaluation("preview_personal"))
                Exit Function
            End If
            If InStr(1, CStr(evaluation("preview_section1")), "727", vbTextCompare) = 0 Then
                ProbeEnrollmentTariffReference = "FAILED: order 727 preamble was not included in the Word preview"
                Exit Function
            End If
            If InStr(1, CStr(evaluation("preview_section1")), "430", vbTextCompare) = 0 Then
                ProbeEnrollmentTariffReference = "FAILED: order 430 preamble was not included in the Word preview"
                Exit Function
            End If
            If CStr(tariffDefinition("word_legal_group")) <> "MO_430" Or CStr(tariffDefinition("word_legal_clause")) = "" Then
                ProbeEnrollmentTariffReference = "FAILED: automatic tariff allowance does not have its legal group and rule clause"
                Exit Function
            End If
            record("extra_monthly1_enabled") = "YES"
            record("extra_monthly1_name") = "Тестовая иная выплата"
            record("extra_monthly1_param") = "TEST-ACT-7"
            record("extra_monthly1_amount") = "10%"
            record("extra_monthly1_start") = "15.07.2026"
            record("extra_monthly1_basis") = "Тестовый подтверждающий документ"
            stageName = "other legal-act evaluation"
            Set evaluation = mdlEnrollmentWorkflow.EvaluateEnrollmentRecord(record)
            If InStr(1, CStr(evaluation("preview_section1")), "TEST-ACT-7", vbTextCompare) = 0 Then
                ProbeEnrollmentTariffReference = "FAILED: other legal-act preamble was not included in the Word preview"
                Exit Function
            End If
            For Each recordKey In record.Keys
                mdlEnrollmentWorkflow.SetBackendValue CStr(recordKey), record(recordKey)
            Next recordKey
            mdlEnrollmentWorkflow.RefreshEnrollmentForm
            If CStr(mdlEnrollmentWorkflow.GetBackendValue("std_tariff_enabled")) <> "YES" Then
                ProbeEnrollmentTariffReference = "FAILED: tariff 1 did not enable the automatic 1-4 allowance"
                Exit Function
            End If
            mdlEnrollmentWorkflow.SetBackendValue "tariff_rank", "5"
            mdlEnrollmentWorkflow.RefreshEnrollmentForm
            If CStr(mdlEnrollmentWorkflow.GetBackendValue("std_tariff_enabled")) <> "NO" Then
                ProbeEnrollmentTariffReference = "FAILED: tariff 5 did not disable the automatic 1-4 allowance"
                Exit Function
            End If
            mdlEnrollmentWorkflow.SetBackendValue "tariff_rank", "1"
            mdlEnrollmentWorkflow.RefreshEnrollmentForm
            Set record = mdlEnrollmentWorkflow.GetBackendRecord
            record("personal_details_enabled") = "YES"
            Set evaluation = mdlEnrollmentWorkflow.EvaluateEnrollmentRecord(record)
            If evaluation("word_ready") <> "NO" Then
                ProbeEnrollmentTariffReference = "FAILED: enabled personal-data block did not enforce required fields"
                Exit Function
            End If
            ProbeEnrollmentTariffReference = "OK"
        End If
    End If
    Exit Function
Failed:
    ProbeEnrollmentTariffReference = "FAILED: runtime " & CStr(Err.Number) & " at " & stageName & " - " & Err.Description
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
