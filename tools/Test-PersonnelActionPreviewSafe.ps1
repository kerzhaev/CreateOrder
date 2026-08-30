[CmdletBinding()]
param(
    [string]$WorkbookPath,
    [string]$SourceDirectory
)

if ([string]::IsNullOrWhiteSpace($WorkbookPath)) { $WorkbookPath = Join-Path $PSScriptRoot '..\CreateOrder.xlsm' }
if ([string]::IsNullOrWhiteSpace($SourceDirectory)) { $SourceDirectory = Join-Path $PSScriptRoot '..\CreateOrder.xlsm.modules' }

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Read-VbaText {
    param([Parameter(Mandatory = $true)][string]$Path)
    return [IO.File]::ReadAllText($Path, [Text.Encoding]::GetEncoding(1251))
}

function Release-ComObject {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return }
    try {
        if ([Runtime.InteropServices.Marshal]::IsComObject($Value)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
        }
    } catch {}
}

function Import-CodeModuleText {
    param(
        [Parameter(Mandatory = $true)][object]$Workbook,
        [Parameter(Mandatory = $true)][string]$ModuleName,
        [Parameter(Mandatory = $true)][string]$ModulePath
    )
    $code = Read-VbaText -Path $ModulePath
    $code = [regex]::Replace($code, '^Attribute VB_Name\s*=\s*"[^"]+"\r?\n', '', 1)
    $component = $null
    $module = $null
    try {
        try { $component = $Workbook.VBProject.VBComponents.Item($ModuleName) } catch {}
        if ($null -eq $component) {
            $component = $Workbook.VBProject.VBComponents.Add(1)
            $component.Name = $ModuleName
        }
        $module = $component.CodeModule
        if ($module.CountOfLines -gt 0) { $module.DeleteLines(1, $module.CountOfLines) }
        $module.AddFromString($code)
    }
    finally {
        Release-ComObject $module
        Release-ComObject $component
    }
}

if (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue) {
    throw 'Excel or Word is running. Close Office applications before the personnel preview test.'
}

$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedSource = [IO.Path]::GetFullPath($SourceDirectory)
$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$testDirectory = Join-Path $projectRoot ("Trash\personnel-action-preview-$((Get-Date).ToString('yyyyMMdd-HHmmss'))")
New-Item -ItemType Directory -Path $testDirectory -Force | Out-Null
$testWorkbookPath = Join-Path $testDirectory 'CreateOrder.personnel-action-preview.xlsm'
Copy-Item -LiteralPath $resolvedWorkbook -Destination $testWorkbookPath -Force

$probeCode = @'
Option Explicit

Private Function LastDataRow(ByVal ws As Worksheet) As Long
    LastDataRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If LastDataRow < 1 Then LastDataRow = 1
End Function

Private Function DataRowCount(ByVal sheetName As String) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    DataRowCount = LastDataRow(ws) - 1
    If DataRowCount < 0 Then DataRowCount = 0
End Function

Private Function RegistrySignature() As String
    Dim names As Variant
    Dim item As Variant
    names = Array("Employees", "EmployeeCurrentState", "PersonnelEvents", "PersonnelStateSnapshots", "PaymentAssignments", "DocumentRegistry")
    For Each item In names
        RegistrySignature = RegistrySignature & CStr(item) & "=" & CStr(DataRowCount(CStr(item))) & ";"
    Next item
End Function

Private Sub AssertTrue(ByVal condition As Boolean, ByVal messageText As String)
    If Not condition Then Err.Raise 970, , messageText
End Sub

Private Function FindField(ByVal preview As Object, ByVal fieldKey As String) As Object
    Dim item As Object
    For Each item In preview("changed_fields")
        If StrComp(CStr(item("key")), fieldKey, vbTextCompare) = 0 Then
            Set FindField = item
            Exit Function
        End If
    Next item
End Function

Private Function FindPayment(ByVal preview As Object, ByVal paymentCode As String) As Object
    Dim item As Object
    For Each item In preview("payment_changes")
        If StrComp(CStr(item("payment_code")), paymentCode, vbTextCompare) = 0 Then
            Set FindPayment = item
            Exit Function
        End If
    Next item
End Function

Private Function WarningSummary(ByVal preview As Object) As String
    Dim warning As Object
    For Each warning In preview("warnings")
        If WarningSummary <> "" Then WarningSummary = WarningSummary & ","
        WarningSummary = WarningSummary & CStr(warning("code"))
    Next warning
End Function

Private Sub AddEmployeeFixture(ByVal employeeID As String)
    Dim ws As Worksheet
    Dim rowNum As Long
    Set ws = ThisWorkbook.Worksheets("Employees")
    rowNum = LastDataRow(ws) + 1
    ws.Cells(rowNum, 1).Value = employeeID
    ws.Cells(rowNum, 2).Value = "Preview Fixture Employee"
    ws.Cells(rowNum, 3).Value = "PREVIEW-PERSONAL"
    ws.Cells(rowNum, 4).Value = "PREVIEW-TABLE"
    ws.Cells(rowNum, 5).Value = "MANUAL"
    ws.Cells(rowNum, 6).Value = "MANUAL_ONLY"
    ws.Cells(rowNum, 8).Value = DateSerial(2026, 8, 1)
    ws.Cells(rowNum, 9).Value = DateSerial(2026, 8, 1)
    ws.Cells(rowNum, 10).Value = "YES"
End Sub

Private Sub AddStateFixture(ByVal employeeID As String)
    Dim ws As Worksheet
    Dim rowNum As Long
    Set ws = ThisWorkbook.Worksheets("EmployeeCurrentState")
    rowNum = LastDataRow(ws) + 1
    ws.Cells(rowNum, 1).Value = employeeID
    ws.Cells(rowNum, 2).Value = "Private"
    ws.Cells(rowNum, 3).Value = DateSerial(2026, 7, 1)
    ws.Cells(rowNum, 4).Value = "Original position"
    ws.Cells(rowNum, 5).Value = "Original section"
    ws.Cells(rowNum, 6).Value = "Original unit"
    ws.Cells(rowNum, 7).Value = "100100"
    ws.Cells(rowNum, 8).Value = "5"
    ws.Cells(rowNum, 9).Value = "25000"
    ws.Cells(rowNum, 10).Value = "10000"
    ws.Cells(rowNum, 11).Value = "REGULAR"
    ws.Cells(rowNum, 17).Value = "SECOND"
    ws.Cells(rowNum, 20).Value = "NO"
    ws.Cells(rowNum, 21).Value = "NO"
    ws.Cells(rowNum, 14).Value = DateSerial(2026, 7, 1)
End Sub

Private Sub AddAssignmentFixture(ByVal employeeID As String, ByVal paymentCode As String, ByVal amountValue As String)
    Dim ws As Worksheet
    Dim rowNum As Long
    Set ws = ThisWorkbook.Worksheets("PaymentAssignments")
    rowNum = LastDataRow(ws) + 1
    ws.Cells(rowNum, 1).Value = "PREVIEW-ASSIGN-" & paymentCode
    ws.Cells(rowNum, 2).Value = employeeID
    ws.Cells(rowNum, 3).Value = "PREVIEW-EVENT"
    ws.Cells(rowNum, 4).Value = "Preview"
    ws.Cells(rowNum, 5).Value = paymentCode
    ws.Cells(rowNum, 6).Value = "PERCENT"
    ws.Cells(rowNum, 7).Value = amountValue
    ws.Cells(rowNum, 11).Value = "ACTIVE"
    ws.Cells(rowNum, 19).Value = "SPECIAL_ACHIEVEMENTS_P2"
    ws.Cells(rowNum, 20).Value = "Preview fixture assignment"
End Sub

Public Function RunPersonnelActionPreview() As String
    Dim employeeID As String
    Dim draft As Object
    Dim exclusionDraft As Object
    Dim invalidDraft As Object
    Dim previewOne As Object
    Dim previewTwo As Object
    Dim exclusionPreview As Object
    Dim invalidPreview As Object
    Dim positionChange As Object
    Dim continuePayment As Object
    Dim stoppedPayment As Object
    Dim beforeSignature As String
    Dim afterSignature As String

    On Error GoTo Failed
    mdlPersonnelEvents.EnsurePersonnelEventInfrastructure
    employeeID = "PREVIEW-FIXTURE-20260829"
    AddEmployeeFixture employeeID
    AddStateFixture employeeID
    AddAssignmentFixture employeeID, "FIZO_SECOND", "80"
    AddAssignmentFixture employeeID, "OLD_PAYMENT", "10"
    beforeSignature = RegistrySignature()

    Set draft = CreateObject("Scripting.Dictionary")
    draft.Add "event_type", "TRANSFER"
    draft.Add "employee_id", employeeID
    draft.Add "event_date", DateSerial(2026, 8, 2)
    draft.Add "effective_date", DateSerial(2026, 8, 3)
    draft.Add "order_reference", "PREVIEW-TRANSFER-001"
    draft.Add "basis_text", "Preview transfer fixture"
    draft.Add "new_position", "Preview new position"
    draft.Add "new_section", "Preview new section"
    draft.Add "new_military_unit", "Preview new unit"
    draft.Add "new_vus", "200200"
    draft.Add "handover_date", DateSerial(2026, 8, 2)
    draft.Add "acceptance_date", DateSerial(2026, 8, 3)
    draft.Add "duty_start_date", DateSerial(2026, 8, 4)
    draft.Add "destination_unit", "Preview destination"
    draft.Add "destination_location", "Preview city"

    Set previewOne = mdlPersonnelActionPreview.BuildPersonnelActionPreview(draft)
    Set previewTwo = mdlPersonnelActionPreview.BuildPersonnelActionPreview(draft)
    afterSignature = RegistrySignature()
    AssertTrue beforeSignature = afterSignature, "TRANSFER preview mutated personnel registries: " & beforeSignature & " -> " & afterSignature
    AssertTrue previewOne("can_confirm") = True, "Valid TRANSFER preview cannot be confirmed; warnings=" & WarningSummary(previewOne)
    AssertTrue previewTwo("can_confirm") = True, "Second TRANSFER preview cannot be confirmed"
    AssertTrue CLng(previewOne("counts")("changed_fields")) = CLng(previewTwo("counts")("changed_fields")), "Repeated preview changed its field count"
    Set positionChange = FindField(previewOne, "new_position")
    AssertTrue Not positionChange Is Nothing, "TRANSFER preview did not include new_position"
    AssertTrue CStr(positionChange("change_kind")) = "CHANGED", "TRANSFER position diff was not CHANGED"
    AssertTrue CStr(positionChange("after")) = "Preview new position", "TRANSFER after position is incorrect"
    Set continuePayment = FindPayment(previewOne, "FIZO_SECOND")
    AssertTrue Not continuePayment Is Nothing, "TRANSFER preview did not include continuing FIZO payment"
    AssertTrue CStr(continuePayment("change_kind")) = "CONTINUE", "Existing FIZO payment was not marked CONTINUE"
    Set stoppedPayment = FindPayment(previewOne, "OLD_PAYMENT")
    AssertTrue Not stoppedPayment Is Nothing, "TRANSFER preview did not include stopped obsolete payment"
    AssertTrue CStr(stoppedPayment("change_kind")) = "STOP", "Obsolete payment was not marked STOP"

    Set exclusionDraft = CreateObject("Scripting.Dictionary")
    exclusionDraft.Add "event_type", "EXCLUSION"
    exclusionDraft.Add "employee_id", employeeID
    exclusionDraft.Add "event_date", DateSerial(2026, 8, 5)
    exclusionDraft.Add "effective_date", DateSerial(2026, 8, 6)
    exclusionDraft.Add "order_reference", "PREVIEW-EXCLUSION-001"
    exclusionDraft.Add "basis_text", "Preview exclusion fixture"
    exclusionDraft.Add "handover_date", DateSerial(2026, 8, 5)
    exclusionDraft.Add "destination_unit", "Preview archive"
    exclusionDraft.Add "destination_location", "Preview city"
    Set exclusionPreview = mdlPersonnelActionPreview.BuildPersonnelActionPreview(exclusionDraft)
    AssertTrue RegistrySignature() = beforeSignature, "EXCLUSION preview mutated personnel registries"
    AssertTrue exclusionPreview("can_confirm") = True, "Valid EXCLUSION preview cannot be confirmed"
    AssertTrue CStr(exclusionPreview("after")("is_active")) = "NO", "EXCLUSION preview did not project IsActive=NO"
    Set stoppedPayment = FindPayment(exclusionPreview, "FIZO_SECOND")
    AssertTrue Not stoppedPayment Is Nothing, "EXCLUSION preview omitted active FIZO stop"
    AssertTrue CStr(stoppedPayment("change_kind")) = "STOP", "EXCLUSION FIZO payment was not marked STOP"

    Set invalidDraft = CreateObject("Scripting.Dictionary")
    invalidDraft.Add "event_type", "TRANSFER"
    invalidDraft.Add "employee_id", employeeID
    invalidDraft.Add "effective_date", DateSerial(2026, 8, 3)
    Set invalidPreview = mdlPersonnelActionPreview.BuildPersonnelActionPreview(invalidDraft)
    AssertTrue invalidPreview("can_confirm") = False, "Invalid preview was confirmable"
    AssertTrue RegistrySignature() = beforeSignature, "Invalid preview mutated personnel registries"

    RunPersonnelActionPreview = "PERSONNEL_ACTION_PREVIEW_OK"
    Exit Function
Failed:
    RunPersonnelActionPreview = "FAILED: " & Err.Description
End Function
'@

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try { $excel.AutomationSecurity = 1 } catch {}
    $workbook = $excel.Workbooks.Open($testWorkbookPath, 0, $false)

    Import-CodeModuleText -Workbook $workbook -ModuleName 'ModuleLocalization' -ModulePath (Join-Path $resolvedSource 'ModuleLocalization.bas')
    Import-CodeModuleText -Workbook $workbook -ModuleName 'mdlPersonnelOrderText' -ModulePath (Join-Path $resolvedSource 'mdlPersonnelOrderText.bas')
    Import-CodeModuleText -Workbook $workbook -ModuleName 'mdlPersonnelAllowanceRules' -ModulePath (Join-Path $resolvedSource 'mdlPersonnelAllowanceRules.bas')
    Import-CodeModuleText -Workbook $workbook -ModuleName 'mdlPersonnelEvents' -ModulePath (Join-Path $resolvedSource 'mdlPersonnelEvents.bas')
    Import-CodeModuleText -Workbook $workbook -ModuleName 'mdlPersonnelActionPreview' -ModulePath (Join-Path $resolvedSource 'mdlPersonnelActionPreview.bas')

    try { $workbook.VBProject.VBComponents.Remove($workbook.VBProject.VBComponents.Item('personnel_preview_probe')) } catch {}
    $probe = $workbook.VBProject.VBComponents.Add(1)
    $probe.Name = 'personnel_preview_probe'
    $probe.CodeModule.AddFromString($probeCode)

    $result = [string]$excel.Run("'$($workbook.Name)'!personnel_preview_probe.RunPersonnelActionPreview")
    if ($result -ne 'PERSONNEL_ACTION_PREVIEW_OK') { throw $result }

    $workbook.Close($false)
    Release-ComObject $workbook
    $workbook = $null
    $excel.Quit()
    Release-ComObject $excel
    $excel = $null
    Write-Output "PERSONNEL_ACTION_PREVIEW_OK|$testWorkbookPath"
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
        Release-ComObject $workbook
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

for ($attempt = 1; $attempt -le 40 -and (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue); $attempt++) {
    Start-Sleep -Milliseconds 250
}
if (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue) {
    throw 'Office process remained running after the personnel preview test.'
}
