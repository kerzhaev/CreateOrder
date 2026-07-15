$ErrorActionPreference = "Stop"

function Assert-NoOfficeProcesses {
    $officeProcesses = @(Get-Process Excel,WINWORD -ErrorAction SilentlyContinue)
    if ($officeProcesses.Count -gt 0) {
        $names = ($officeProcesses | ForEach-Object { "$($_.ProcessName) (PID $($_.Id))" }) -join ", "
        throw "Acceptance test was not started because Office is already open: $names. Close only the disposable Office sessions and run the test again; this script never terminates user sessions."
    }
}

function Remove-IfExists {
    param([string]$Path)
    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Force -Recurse
    }
}

function Assert-True {
    param(
        [bool]$Condition,
        [string]$Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

function Assert-Eq {
    param(
        $Actual,
        $Expected,
        [string]$Message
    )

    if ($Actual -ne $Expected) {
        throw "$Message`nExpected: $Expected`nActual: $Actual"
    }
}

function Read-VbaText {
    param([string]$Path)

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $utf8 = New-Object System.Text.UTF8Encoding($false, $true)

    try {
        return $utf8.GetString($bytes)
    }
    catch [System.Text.DecoderFallbackException] {
        return [System.Text.Encoding]::GetEncoding(1251).GetString($bytes)
    }
}

function Assert-RibbonHandlersUseLocalization {
    param([string]$ModulePath)

    $source = Read-VbaText -Path $ModulePath
    Assert-True (-not [regex]::IsMatch($source, 'MsgBox\s+"')) "mdlRibbonHandlers contains direct MsgBox string literals; use t/tf localization wrappers."
    Assert-True (-not [regex]::IsMatch($source, 'Application\.StatusBar\s*=\s+"')) "mdlRibbonHandlers contains direct StatusBar string literals; use t/tf localization wrappers."
}

function Assert-PaymentModulesUseLocalization {
    param([string[]]$ModulePaths)

    foreach ($modulePath in $ModulePaths) {
        $source = Read-VbaText -Path $modulePath
        $moduleName = [IO.Path]::GetFileNameWithoutExtension($modulePath)
        Assert-True (-not [regex]::IsMatch($source, 'MsgBox\s+"')) "$moduleName contains direct MsgBox string literals; use t/tf localization wrappers."
        Assert-True (-not [regex]::IsMatch($source, 'Application\.StatusBar\s*=\s+"')) "$moduleName contains direct StatusBar string literals; use t/tf localization wrappers."
        Assert-True (-not [regex]::IsMatch($source, 'Err\.Raise[^\r\n]*"[А-Яа-яЁё]')) "$moduleName contains direct Russian Err.Raise text; use t/tf localization wrappers."
    }
}

function Assert-CustomUiUsesLocalizationCallbacks {
    param([string]$CustomUiPath)

    $source = Get-Content -LiteralPath $CustomUiPath -Raw -Encoding UTF8
    Assert-True (-not [regex]::IsMatch($source, '\s(label|screentip|supertip)\s*=')) "customUI14.xml contains static ribbon text; use localization callbacks instead."
    Assert-True ($source.Contains('getLabel="GetRibbonLabel"')) "customUI14.xml does not use GetRibbonLabel callback."
    Assert-True ($source.Contains('getScreentip="GetRibbonScreentip"')) "customUI14.xml does not use GetRibbonScreentip callback for ribbon tips."
    Assert-True ($source.Contains('getSupertip="GetRibbonSupertip"')) "customUI14.xml does not use GetRibbonSupertip callback for ribbon tips."
}

function Assert-EnrollmentLocalizationKeysSeeded {
    param(
        [string[]]$ModulePaths,
        [string]$LocalizationModulePath
    )

    $keyPattern = '(?:\b(?:t|tf|L|ET|LocalizeDirect)\s*\(|AddPage(?:TextBox|CheckBox|ComboBox|Memo|TextBoxT)\s*\([^,]+,\s*)"([a-zA-Z0-9_.-]+)"'
    $keys = New-Object "System.Collections.Generic.HashSet[string]"

    foreach ($modulePath in $ModulePaths) {
        $source = Read-VbaText -Path $modulePath
        foreach ($match in [regex]::Matches($source, $keyPattern)) {
            $key = $match.Groups[1].Value
            if ([string]::IsNullOrWhiteSpace($key)) { continue }
            if ($key.EndsWith(".")) { continue }
            [void]$keys.Add($key)
        }
    }

    $localizationSource = Read-VbaText -Path $LocalizationModulePath
    $seededKeys = New-Object "System.Collections.Generic.HashSet[string]"
    foreach ($match in [regex]::Matches($localizationSource, 'AddSafe\s+"([^"]+)"')) {
        [void]$seededKeys.Add($match.Groups[1].Value)
    }

    $missing = @($keys | Where-Object { -not $seededKeys.Contains($_) } | Sort-Object)
    Assert-True ($missing.Count -eq 0) ("Enrollment localization keys are not seeded safely: " + ($missing -join ", "))
}

function Assert-LocalizationKeysSeededSafely {
    param(
        [string[]]$Keys,
        [string]$LocalizationModulePath
    )

    $localizationSource = Read-VbaText -Path $LocalizationModulePath
    $seededKeys = New-Object "System.Collections.Generic.HashSet[string]"
    foreach ($match in [regex]::Matches($localizationSource, 'AddSafe\s+"([^"]+)"')) {
        [void]$seededKeys.Add($match.Groups[1].Value)
    }

    $missing = @($Keys | Where-Object { -not $seededKeys.Contains($_) } | Sort-Object)
    Assert-True ($missing.Count -eq 0) ("Core localization keys are not seeded through AddSafe: " + ($missing -join ", "))
}

function Import-CodeModuleText {
    param(
        $Workbook,
        [string]$ModuleName,
        [string]$ModulePath
    )

    $code = Read-VbaText -Path $ModulePath
    $code = [regex]::Replace($code, '^Attribute VB_Name\s*=\s*"[^"]+"\r?\n', '', 1)

    $component = $Workbook.VBProject.VBComponents.Item($ModuleName)
    $codeModule = $component.CodeModule
    if ($codeModule.CountOfLines -gt 0) {
        $codeModule.DeleteLines(1, $codeModule.CountOfLines)
    }
    $codeModule.AddFromString($code)
}

function Import-DocumentModuleText {
    param(
        $Workbook,
        [string]$ComponentName,
        [string]$ModulePath
    )

    $code = Read-VbaText -Path $ModulePath

    $component = $Workbook.VBProject.VBComponents.Item($ComponentName)
    $codeModule = $component.CodeModule
    if ($codeModule.CountOfLines -gt 0) {
        $codeModule.DeleteLines(1, $codeModule.CountOfLines)
    }
    $codeModule.AddFromString($code)
}

function Import-UserFormComponent {
    param(
        $Workbook,
        [string]$FormName,
        [string]$FormPath
    )

    try {
        $Workbook.VBProject.VBComponents.Remove($Workbook.VBProject.VBComponents.Item($FormName))
    }
    catch {
    }

    $null = $Workbook.VBProject.VBComponents.Import($FormPath)
}

function Add-TestProbeModule {
    param($Workbook)

    try {
        $Workbook.VBProject.VBComponents.Remove($Workbook.VBProject.VBComponents.Item("codex_acceptance_probe"))
    }
    catch {
    }

    $component = $Workbook.VBProject.VBComponents.Add(1)
    $component.Name = "codex_acceptance_probe"
    $component.CodeModule.AddFromString(@"
Option Explicit

Public Function ProbeEnrollmentConflict(ByVal orderDraftId As String) As String
    On Error GoTo ErrorHandler
    ProbeEnrollmentConflict = mdlEnrollmentOrderExport.ExportEnrollmentOrderByDraftId(orderDraftId)
    Exit Function
ErrorHandler:
    ProbeEnrollmentConflict = "ERROR: " & Err.Description
End Function

Public Function ProbeEnrollmentConflictRaw(ByVal orderDraftId As String) As String
    ProbeEnrollmentConflictRaw = mdlEnrollmentOrderExport.ExportEnrollmentOrderByDraftId(orderDraftId)
End Function

Public Function ProbeEnrollmentExportRow(ByVal rowNum As Long) As String
    On Error GoTo ErrorHandler
    ProbeEnrollmentExportRow = mdlEnrollmentOrderExport.ExportEnrollmentOrderByRow(rowNum)
    Exit Function
ErrorHandler:
    ProbeEnrollmentExportRow = "ERROR: " & Err.Description
End Function

Public Function ProbeEnrollmentExportRowRaw(ByVal rowNum As Long) As String
    ProbeEnrollmentExportRowRaw = mdlEnrollmentOrderExport.ExportEnrollmentOrderByRow(rowNum)
End Function

Public Function ProbeEnrollmentExportBlockingIssues(ByVal orderDraftId As String, ByVal fallbackRow As Long) As String
    On Error GoTo ErrorHandler
    ProbeEnrollmentExportBlockingIssues = mdlEnrollmentOrderExport.GetEnrollmentExportBlockingIssues(orderDraftId, fallbackRow)
    Exit Function
ErrorHandler:
    ProbeEnrollmentExportBlockingIssues = "ERROR: " & Err.Description
End Function

Public Function ProbeEnrollmentDefinitionBlockCount(ByVal wordBlockTarget As String) As Long
    ProbeEnrollmentDefinitionBlockCount = mdlEnrollmentWorkflow.GetEnrollmentPaymentDefinitionsByBlock(wordBlockTarget).Count
End Function

Public Function ProbeContinuePackage(ByVal rowNum As Long) As String
    On Error GoTo ErrorHandler
    Dim draftId As String

    draftId = mdlEnrollmentWorkflow.PrepareNextEnrollmentInPackage(rowNum)
    ProbeContinuePackage = draftId & "|" & CStr(mdlEnrollmentWorkflow.GetBackendValue("fio")) & "|" & CStr(mdlEnrollmentWorkflow.GetBackendValue("order_number")) & "|" & CStr(mdlEnrollmentWorkflow.GetBackendValue("std_duty_enabled"))
    Exit Function
ErrorHandler:
    ProbeContinuePackage = "ERROR: " & Err.Description
End Function

Public Function ProbeSaveAndContinuePackage() As String
    On Error GoTo ErrorHandler
    Dim draftId As String

    draftId = mdlEnrollmentWorkflow.SaveEnrollmentFormAndContinuePackage()
    ProbeSaveAndContinuePackage = draftId & "|" & CStr(mdlEnrollmentWorkflow.GetBackendValue("fio")) & "|" & CStr(mdlEnrollmentWorkflow.GetBackendValue("order_number")) & "|" & CStr(mdlEnrollmentWorkflow.GetBackendValue("std_duty_enabled"))
    Exit Function
ErrorHandler:
    ProbeSaveAndContinuePackage = "ERROR: " & Err.Description
End Function

Private Sub SafeUnloadEnrollmentWizard()
    On Error Resume Next
    Dim currentForm As Object

    For Each currentForm In VBA.UserForms
        If TypeName(currentForm) = "frmEnrollmentWizard" Then
            Unload currentForm
            Exit For
        End If
    Next currentForm
End Sub

Public Function ProbeEnrollmentWizardStaffLoad(ByVal personalNumber As String) As String
    On Error GoTo ErrorHandler

    Load frmEnrollmentWizard
    frmEnrollmentWizard.LoadEmployeeFromStaffNumber personalNumber
    ProbeEnrollmentWizardStaffLoad = frmEnrollmentWizard.GetEmployeeSnapshot()
    SafeUnloadEnrollmentWizard
    Exit Function
ErrorHandler:
    SafeUnloadEnrollmentWizard
    ProbeEnrollmentWizardStaffLoad = "ERROR: " & Err.Description
End Function

Public Function ProbeEnrollmentWizardInlineSearch(ByVal queryText As String) As String
    On Error GoTo ErrorHandler

    Load frmEnrollmentWizard
    ProbeEnrollmentWizardInlineSearch = frmEnrollmentWizard.ProbeInlineSearch(queryText)
    SafeUnloadEnrollmentWizard
    Exit Function
ErrorHandler:
    SafeUnloadEnrollmentWizard
    ProbeEnrollmentWizardInlineSearch = "ERROR: " & Err.Description
End Function

Public Function ProbeEnrollmentWizardInlineSearchUiText() As String
    On Error GoTo ErrorHandler

    Load frmEnrollmentWizard
    ProbeEnrollmentWizardInlineSearchUiText = CStr(frmEnrollmentWizard.Controls("txtSearch").ControlTipText) & "|" & CStr(frmEnrollmentWizard.Controls("btnLoadFromInlineSearchDynamic").Caption)
    SafeUnloadEnrollmentWizard
    Exit Function
ErrorHandler:
    SafeUnloadEnrollmentWizard
    ProbeEnrollmentWizardInlineSearchUiText = "ERROR: " & Err.Description
End Function

Public Function ProbeEnrollmentWizardLayout() As String
    On Error GoTo ErrorHandler

    Load frmEnrollmentWizard
    ProbeEnrollmentWizardLayout = frmEnrollmentWizard.ProbeLayoutSnapshot()
    SafeUnloadEnrollmentWizard
    Exit Function
ErrorHandler:
    SafeUnloadEnrollmentWizard
    ProbeEnrollmentWizardLayout = "ERROR: " & Err.Description
End Function

Public Function ProbeEnrollmentWizardSaveGenerate() As String
    On Error GoTo ErrorHandler

    Load frmEnrollmentWizard
    ProbeEnrollmentWizardSaveGenerate = frmEnrollmentWizard.RunSaveGenerateAction()
    SafeUnloadEnrollmentWizard
    Exit Function
ErrorHandler:
    SafeUnloadEnrollmentWizard
    ProbeEnrollmentWizardSaveGenerate = "ERROR: " & Err.Description
End Function

Public Function ProbeEnrollmentWizardSaveContinue() As String
    On Error GoTo ErrorHandler

    Load frmEnrollmentWizard
    ProbeEnrollmentWizardSaveContinue = frmEnrollmentWizard.RunSaveContinuePackageAction()
    SafeUnloadEnrollmentWizard
    Exit Function
ErrorHandler:
    SafeUnloadEnrollmentWizard
    ProbeEnrollmentWizardSaveContinue = "ERROR: " & Err.Description
End Function

Public Function ProbeEnrollmentWizardFullRoundTrip(ByVal rowNum As Long) As String
    On Error GoTo ErrorHandler
    Dim beforeSnapshot As String
    Dim saveResult As String
    Dim afterSnapshot As String

    mdlEnrollmentWorkflow.LoadEnrollmentRowToBackend rowNum
    Load frmEnrollmentWizard
    beforeSnapshot = frmEnrollmentWizard.ProbeFullCardSnapshot()
    saveResult = frmEnrollmentWizard.RunSaveCardAction()
    afterSnapshot = frmEnrollmentWizard.ProbeFullCardSnapshot()
    ProbeEnrollmentWizardFullRoundTrip = beforeSnapshot & "||" & saveResult & "||" & afterSnapshot
    SafeUnloadEnrollmentWizard
    Exit Function
ErrorHandler:
    SafeUnloadEnrollmentWizard
    ProbeEnrollmentWizardFullRoundTrip = "ERROR: " & Err.Description
End Function

Public Function ProbeEnrollmentWizardCheck() As String
    On Error GoTo ErrorHandler

    Load frmEnrollmentWizard
    ProbeEnrollmentWizardCheck = frmEnrollmentWizard.RunCheckAction()
    SafeUnloadEnrollmentWizard
    Exit Function
ErrorHandler:
    SafeUnloadEnrollmentWizard
    ProbeEnrollmentWizardCheck = "ERROR: " & Err.Description
End Function

Public Function ProbeEnrollmentWizardExport() As String
    On Error GoTo ErrorHandler

    Load frmEnrollmentWizard
    ProbeEnrollmentWizardExport = frmEnrollmentWizard.RunExportAction()
    SafeUnloadEnrollmentWizard
    Exit Function
ErrorHandler:
    SafeUnloadEnrollmentWizard
    ProbeEnrollmentWizardExport = "ERROR: " & Err.Description
End Function

Public Function ProbeRibbonLocalization() As String
    On Error GoTo ErrorHandler

    ModuleLocalization.ResetLocalizationCache
    ProbeRibbonLocalization = ModuleLocalization.t("ribbon.error.main_export", "") & "|" & _
        ModuleLocalization.t("ribbon.smart_validation.title", "") & "|" & _
        ModuleLocalization.t("ribbon.error.open_workbook_folder", "")
    Exit Function
ErrorHandler:
    ProbeRibbonLocalization = "ERROR: " & Err.Description
End Function

Public Function ProbeRibbonUiLocalization() As String
    On Error GoTo ErrorHandler

    ModuleLocalization.ResetLocalizationCache
    ProbeRibbonUiLocalization = mdlRibbonHandlers.GetRibbonUiTextById("openEnrollmentForm", "label") & "|" & _
        mdlRibbonHandlers.GetRibbonUiTextById("importWordRaport", "screentip") & "|" & _
        mdlRibbonHandlers.GetRibbonUiTextById("validateZP12", "supertip") & "|" & _
        mdlRibbonHandlers.GetRibbonUiTextById("exportEnrollmentPackageById", "label") & "|" & _
        ModuleLocalization.t("enrollment.ribbon.prompt.order_draft_id", "") & "|" & _
        ModuleLocalization.t("enrollment.ribbon.message.order_draft_id_empty", "") & "|" & _
        ModuleLocalization.t("enrollment.ribbon.exported", "") & "|" & _
        ModuleLocalization.t("enrollment.ribbon.error.export_by_id", "")
    Exit Function
ErrorHandler:
    ProbeRibbonUiLocalization = "ERROR: " & Err.Description
End Function

Public Function ProbeSettingsDiagnosticText() As String
    On Error GoTo ErrorHandler

    ModuleLocalization.ResetLocalizationCache
    ProbeSettingsDiagnosticText = mdlRibbonHandlers.BuildSettingsDiagnosticText()
    Exit Function
ErrorHandler:
    ProbeSettingsDiagnosticText = "ERROR: " & Err.Description
End Function

Public Function ProbeCoreLocalization() As String
    On Error GoTo ErrorHandler

    ModuleLocalization.ResetLocalizationCache
    ProbeCoreLocalization = ModuleLocalization.t("product.name", "fallback") & "|" & _
        ModuleLocalization.t("form.about.title", "fallback") & "|" & _
        ModuleLocalization.t("license.caption.service", "fallback") & "|" & _
        ModuleLocalization.t("form.license_actions.title", "fallback")
    Exit Function
ErrorHandler:
    ProbeCoreLocalization = "ERROR: " & Err.Description
End Function

Public Function ProbeLegacyUiLocalization() As String
    On Error GoTo ErrorHandler

    ModuleLocalization.ResetLocalizationCache
    ProbeLegacyUiLocalization = ModuleLocalization.t("form.search_fio.title", "fallback") & "|" & _
        ModuleLocalization.t("form.select_employee.title", "fallback") & "|" & _
        ModuleLocalization.t("form.search_fio.message.long_period_confirm", "fallback") & "|" & _
        ModuleLocalization.t("license.details.header", "fallback")
    Exit Function
ErrorHandler:
    ProbeLegacyUiLocalization = "ERROR: " & Err.Description
End Function
"@)
}

function Remove-TestProbeModule {
    param($Workbook)

    if ($null -eq $Workbook) {
        return
    }

    try {
        $Workbook.VBProject.VBComponents.Remove($Workbook.VBProject.VBComponents.Item("codex_acceptance_probe"))
    }
    catch {
    }
}

function Get-WorksheetByName {
    param(
        $Workbook,
        [string]$SheetName
    )

    foreach ($sheet in $Workbook.Worksheets) {
        if ($sheet.Name -eq $SheetName) {
            return $sheet
        }
    }

    return $null
}

function Get-ShapeText {
    param($Shape)

    return $Shape.TextFrame.Characters().Text
}

function Get-LastUsedRow {
    param(
        $Worksheet,
        [int]$ColumnNumber
    )

    return $Worksheet.Cells($Worksheet.Rows.Count, $ColumnNumber).End(-4162).Row
}

function Get-CellValue {
    param(
        $Worksheet,
        [int]$RowNumber,
        [int]$ColumnNumber
    )

    return $Worksheet.Cells($RowNumber, $ColumnNumber).Value2
}

function Get-CellText {
    param(
        $Worksheet,
        [int]$RowNumber,
        [int]$ColumnNumber
    )

    return [string]$Worksheet.Cells($RowNumber, $ColumnNumber).Text
}

function Get-BackendFieldRow {
    param(
        $Worksheet,
        [string]$FieldKey
    )

    $lastRow = $Worksheet.Cells($Worksheet.Rows.Count, 1).End(-4162).Row
    for ($row = 2; $row -le $lastRow; $row++) {
        if ([string]$Worksheet.Cells($row, 1).Value2 -eq $FieldKey) {
            return $row
        }
    }

    return 0
}

function Set-BackendValue {
    param(
        $Worksheet,
        [string]$FieldKey,
        $FieldValue
    )

    $row = Get-BackendFieldRow -Worksheet $Worksheet -FieldKey $FieldKey
    if ($row -le 0) {
        throw "Backend field '$FieldKey' was not found on the enrollment form sheet."
    }

    $Worksheet.Cells($row, 3).Value2 = $FieldValue
}

function Get-BackendValue {
    param(
        $Worksheet,
        [string]$FieldKey
    )

    $row = Get-BackendFieldRow -Worksheet $Worksheet -FieldKey $FieldKey
    if ($row -le 0) {
        throw "Backend field '$FieldKey' was not found on the enrollment form sheet."
    }

    return [string]$Worksheet.Cells($row, 3).Value2
}

function Get-SettingRow {
    param(
        $Worksheet,
        [string]$SettingKey
    )

    $lastRow = $Worksheet.Cells($Worksheet.Rows.Count, 1).End(-4162).Row
    for ($row = 1; $row -le $lastRow; $row++) {
        if ([string]$Worksheet.Cells($row, 1).Value2 -eq $SettingKey) {
            return $row
        }
    }

    return 0
}

function Set-EnrollmentSetting {
    param(
        $Worksheet,
        [string]$SettingKey,
        [string]$SettingValue
    )

    $row = Get-SettingRow -Worksheet $Worksheet -SettingKey $SettingKey
    if ($row -le 0) {
        throw "Enrollment setting '$SettingKey' was not found."
    }

    $Worksheet.Cells($row, 2).Value2 = $SettingValue
}

function Set-TextCell {
    param(
        $Worksheet,
        [int]$RowNumber,
        [int]$ColumnNumber,
        [string]$Value
    )

    $Worksheet.Cells($RowNumber, $ColumnNumber).NumberFormat = "@"
    $Worksheet.Cells($RowNumber, $ColumnNumber).Value2 = "'" + $Value
}

function Set-EnrollmentRequiredFields {
    param(
        $Worksheet,
        [int]$RowNumber,
        [string]$OrderIssuer = "начальника пункта отбора",
        [string]$BasisSection1 = "выписка из приказа, предписание, рапорт",
        [string]$PassportSeries = "1234",
        [string]$PassportNumber = "567890",
        [string]$PassportIssuer = "Тестовый орган",
        [string]$PassportIssueDate = "01.01.2025",
        [string]$PassportCode = "320-009",
        [string]$Inn = "123456789012",
        [string]$Snils = "123-456-789 00",
        [string]$ContractBasis = "контракт / нормативный блок",
        [string]$BankAccount = "40817810000000000001",
        [string]$BankName = "СБЕР БАНК"
    )

    $Worksheet.Cells($RowNumber, 31).Value = $ContractBasis
    $Worksheet.Cells($RowNumber, 35).Value = "18000"
    $Worksheet.Cells($RowNumber, 37).Value = $OrderIssuer
    $Worksheet.Cells($RowNumber, 46).Value = $BasisSection1
    Set-TextCell -Worksheet $Worksheet -RowNumber $RowNumber -ColumnNumber 51 -Value $Inn
    Set-TextCell -Worksheet $Worksheet -RowNumber $RowNumber -ColumnNumber 52 -Value $Snils
    Set-TextCell -Worksheet $Worksheet -RowNumber $RowNumber -ColumnNumber 53 -Value $PassportSeries
    Set-TextCell -Worksheet $Worksheet -RowNumber $RowNumber -ColumnNumber 54 -Value $PassportNumber
    $Worksheet.Cells($RowNumber, 55).Value = $PassportIssuer
    $Worksheet.Cells($RowNumber, 56).Value = $PassportIssueDate
    Set-TextCell -Worksheet $Worksheet -RowNumber $RowNumber -ColumnNumber 57 -Value $PassportCode
    Set-TextCell -Worksheet $Worksheet -RowNumber $RowNumber -ColumnNumber 58 -Value $BankAccount
    $Worksheet.Cells($RowNumber, 59).Value = $BankName
}

function New-Text {
    param([int[]]$CodePoints)

    return -join ($CodePoints | ForEach-Object { [char]$_ })
}

function Get-MatchCount {
    param(
        [string]$Text,
        [string]$Pattern
    )

    if ([string]::IsNullOrEmpty($Text) -or [string]::IsNullOrEmpty($Pattern)) {
        return 0
    }

    return ([regex]::Matches($Text, [regex]::Escape($Pattern))).Count
}

function Get-DocxText {
    param([string]$Path)

    $xml = Get-DocxXml -Path $Path
    if ([string]::IsNullOrEmpty($xml)) {
        return ""
    }

    return [regex]::Replace($xml, "<[^>]+>", "")
}

function Get-DocxXml {
    param([string]$Path)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $zip.GetEntry("word/document.xml")
        if ($null -eq $entry) {
            return ""
        }

        $reader = New-Object System.IO.StreamReader($entry.Open())
        try {
            $xml = $reader.ReadToEnd()
        }
        finally {
            $reader.Close()
        }

        return $xml
    }
    finally {
        $zip.Dispose()
    }
}

function New-ListTemplateDocx {
    param(
        [string]$Path,
        [string]$BodyText
    )

    $word = $null
    $doc = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $doc = $word.Documents.Add()
        $doc.Range().Text = $BodyText
        $doc.SaveAs([ref]$Path, [ref]16)
    }
    finally {
        if ($doc -ne $null) {
            try { $doc.Close($false) } catch {}
        }
        if ($word -ne $null) {
            try { $word.Quit() } catch {}
        }
    }
}

Assert-NoOfficeProcesses

$workspace = "C:\Users\Nachfin\Desktop\Projets\CreateOrder"
$testDir = Join-Path $workspace "_tmp_acceptance_test"
$workbookPath = Join-Path $testDir "CreateOrder.xlsm"

Assert-RibbonHandlersUseLocalization -ModulePath (Join-Path $workspace "CreateOrder.xlsm.modules\mdlRibbonHandlers.bas")
Assert-PaymentModulesUseLocalization -ModulePaths @(
    (Join-Path $workspace "CreateOrder.xlsm.modules\mdlUniversalPaymentExport.bas"),
    (Join-Path $workspace "CreateOrder.xlsm.modules\mdlPaymentValidation.bas")
)
Assert-CustomUiUsesLocalizationCallbacks -CustomUiPath (Join-Path $workspace "resources\customUI14.xml")
Assert-EnrollmentLocalizationKeysSeeded -ModulePaths @(
    (Join-Path $workspace "CreateOrder.xlsm.modules\frmEnrollmentWizard.frm"),
    (Join-Path $workspace "CreateOrder.xlsm.modules\mdlEnrollmentOrderExport.bas"),
    (Join-Path $workspace "CreateOrder.xlsm.modules\mdlEnrollmentWorkflow.bas")
) -LocalizationModulePath (Join-Path $workspace "CreateOrder.xlsm.modules\ModuleLocalization.bas")
Assert-EnrollmentLocalizationKeysSeeded -ModulePaths @(
    (Join-Path $workspace "CreateOrder.xlsm.modules\frmSearchFIO.frm"),
    (Join-Path $workspace "CreateOrder.xlsm.modules\frmSelectEmployee.frm"),
    (Join-Path $workspace "CreateOrder.xlsm.modules\frmAbout.frm"),
    (Join-Path $workspace "CreateOrder.xlsm.modules\frmLicenseActions.frm"),
    (Join-Path $workspace "CreateOrder.xlsm.modules\modActivation.bas")
) -LocalizationModulePath (Join-Path $workspace "CreateOrder.xlsm.modules\ModuleLocalization.bas")
Assert-EnrollmentLocalizationKeysSeeded -ModulePaths @(
    (Join-Path $workspace "CreateOrder.xlsm.modules\mdlUniversalPaymentExport.bas"),
    (Join-Path $workspace "CreateOrder.xlsm.modules\mdlPaymentValidation.bas")
) -LocalizationModulePath (Join-Path $workspace "CreateOrder.xlsm.modules\ModuleLocalization.bas")
Assert-LocalizationKeysSeededSafely -LocalizationModulePath (Join-Path $workspace "CreateOrder.xlsm.modules\ModuleLocalization.bas") -Keys @(
    "product.name",
    "product.author",
    "product.company",
    "product.activation_hint",
    "form.about.title",
    "form.about.button.export_request",
    "form.about.button.import_response",
    "form.about.button.license_status",
    "form.license_actions.title",
    "form.license_actions.export",
    "form.license_actions.import",
    "license.caption.service",
    "license.caption.state",
    "license.status.personal_active",
    "license.status.corporate_active",
    "license.status.trial_active"
)

$sheetDso = New-Text @(1044,1057,1054)
$sheetPayments = New-Text @(1042,1099,1087,1083,1072,1090,1099,95,1041,1077,1079,95,1055,1077,1088,1080,1086,1076,1086,1074)
$sheetEnrollment = New-Text @(1047,1072,1095,1080,1089,1083,1077,1085,1080,1077)
$sheetStaff = New-Text @(1064,1090,1072,1090)
$sheetVusCrew = New-Text @(1057,1087,1088,1072,1074,1086,1095,1085,1080,1082,95,1042,1059,1057,95,1069,1082,1080,1087,1072,1078)
$sheetPaymentTypes = New-Text @(1057,1087,1088,1072,1074,1086,1095,1085,1080,1082,95,1058,1080,1087,1099,95,1042,1099,1087,1083,1072,1090)
$paymentClassQualification = New-Text @(1050,1083,1072,1089,1089,1085,1072,1103,32,1082,1074,1072,1083,1080,1092,1080,1082,1072,1094,1080,1103)
$paymentCrew = New-Text @(1069,1082,1080,1087,1072,1078)
$paymentStdDuty = New-Text @(1053,1072,1076,1073,1072,1074,1082,1072,32,1087,1086,32,1074,1086,1080,1085,1089,1082,1086,1081,32,1076,1086,1083,1078,1085,1086,1089,1090,1080)
$paymentStdSpecial = New-Text @(1054,1089,1086,1073,1099,1077,32,1091,1089,1083,1086,1074,1080,1103,32,1089,1083,1091,1078,1073,1099)
$previewStdSpecial = New-Text @(1054,1089,1086,1073,1099,1077,32,1091,1089,1083,1086,1074,1080,1103)
$paymentFizo = New-Text @(1060,1048,1047,1054)
$paymentSecrecy = New-Text @(1057,1077,1082,1088,1077,1090,1085,1086,1089,1090,1100)
$paymentAchievement = New-Text @(1054,1089,1086,1073,1099,1077,32,1076,1086,1089,1090,1080,1078,1077,1085,1080,1103)
$paymentStdTariff = New-Text @(1053,1072,1076,1073,1072,1074,1082,1072,32,49,45,52,32,1090,1072,1088,1080,1092,1085,1099,1093,32,1088,1072,1079,1088,1103,1076,1086,1074)
$paymentStdContract430 = New-Text @(1053,1072,1076,1073,1072,1074,1082,1072,32,52,51,48,32,1087,1088,1080,1082,1072,1079)
$statusReady = New-Text @(1043,1086,1090,1086,1074,1086)
$paymentStatusOk = "OK"
$statusWarning = New-Text @(1055,1088,1077,1076,1091,1087,1088,1077,1078,1076,1077,1085,1080,1077)
$statusBlocked = New-Text @(1041,1083,1086,1082,1080,1088,1086,1074,1072,1085,1086)
$dateSourceManual = "MANUAL"
$previewSheetName = "Payment_Preview"
$groupExportYes = "YES"
$enrollmentTemplateName = New-Text @(1064,1072,1073,1083,1086,1085,95,1047,1072,1095,1080,1089,1083,1077,1085,1080,1077,46,100,111,99,120)
$enrollmentBodyMarker = "[ENROLLMENT_ORDER_BODY]"
$legacyEnrollmentTemplateTokens = @("[ЗВАНИЕ_СКЛОНЕННОЕ]", "[ФИО_ИМЕНИТЕЛЬНЫЙ]", "[РАЗМЕР]", "[ОСНОВАНИЕ]")

$rootEnrollmentTemplatePath = Join-Path $workspace $enrollmentTemplateName
Assert-True (Test-Path -LiteralPath $rootEnrollmentTemplatePath) "Enrollment template Шаблон_Зачисление.docx is missing from the project root."
$rootEnrollmentTemplateText = Get-DocxText -Path $rootEnrollmentTemplatePath
Assert-True ($rootEnrollmentTemplateText.Contains($enrollmentBodyMarker)) "Enrollment template does not contain the body marker $enrollmentBodyMarker."
foreach ($legacyToken in $legacyEnrollmentTemplateTokens) {
    Assert-True (-not $rootEnrollmentTemplateText.Contains($legacyToken)) "Enrollment template still contains legacy placeholder $legacyToken."
}

Remove-IfExists -Path $testDir
New-Item -ItemType Directory -Path $testDir | Out-Null
Copy-Item -LiteralPath (Join-Path $workspace "CreateOrder.xlsm") -Destination $workbookPath
Get-ChildItem -LiteralPath $workspace -Filter "*.docx" | ForEach-Object {
    Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $testDir $_.Name)
}

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    try {
        $excel.AutomationSecurity = 1
    } catch {
    }

    $workbook = $excel.Workbooks.Open($workbookPath, 0, $false)

    $moduleBase = Join-Path $workspace "CreateOrder.xlsm.modules"
    $documentModuleBase = Join-Path $workspace "workbook-modules"
    $modulesToImport = @(
        "mdlHelper",
        "mdlDataValidation",
        "ModuleLocalization",
        "mdlRibbonHandlers",
        "mdlReferenceData",
        "mdlPaymentTypes",
        "mdlPaymentValidation",
        "mdlPaymentPackageSupport",
        "mdlUniversalPaymentExport",
        "mdlEnrollmentOrderExport",
        "mdlEnrollmentWorkflow"
    )

    foreach ($moduleName in $modulesToImport) {
        Import-CodeModuleText -Workbook $workbook -ModuleName $moduleName -ModulePath (Join-Path $moduleBase ($moduleName + ".bas"))
    }
    Import-DocumentModuleText -Workbook $workbook -ComponentName "Лист1" -ModulePath (Join-Path $documentModuleBase "Лист1 (ДСО).bas")
    Import-UserFormComponent -Workbook $workbook -FormName "frmSearchFIO" -FormPath (Join-Path $moduleBase "frmSearchFIO.frm")
    Import-UserFormComponent -Workbook $workbook -FormName "frmSelectEmployee" -FormPath (Join-Path $moduleBase "frmSelectEmployee.frm")
    Import-UserFormComponent -Workbook $workbook -FormName "frmEnrollmentWizard" -FormPath (Join-Path $moduleBase "frmEnrollmentWizard.frm")
    Add-TestProbeModule -Workbook $workbook

    $excel.Run("'$($workbook.Name)'!mdlPaymentPackageSupport.EnsurePaymentsFeatureInfrastructure")
    $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.EnsureEnrollmentInfrastructure")
    Remove-IfExists -Path (Join-Path $testDir $enrollmentTemplateName)
    $excel.Run("'$($workbook.Name)'!mdlEnrollmentOrderExport.EnsureEnrollmentTemplateAvailable")
    Assert-True (Test-Path -LiteralPath (Join-Path $testDir $enrollmentTemplateName)) "Enrollment template fallback did not create the dedicated enrollment Word template."

    $dsoCode = $workbook.VBProject.VBComponents.Item("Лист1").CodeModule.Lines(1, $workbook.VBProject.VBComponents.Item("Лист1").CodeModule.CountOfLines)
    Assert-True ($dsoCode -match "Target\.CountLarge") "DSO worksheet module did not receive the CountLarge overflow fix."
    Assert-True ($dsoCode -notmatch "Target\.Count\s*>\s*50") "DSO worksheet module still contains the overflow-prone Target.Count check."

    Write-Output "1. Self-healing"
    foreach ($sheetName in @($sheetDso, $sheetPayments, $sheetEnrollment)) {
        $sheet = Get-WorksheetByName -Workbook $workbook -SheetName $sheetName
        if ($null -ne $sheet) {
            $sheet.Delete()
        }
    }
    $excel.Run("'$($workbook.Name)'!mdlDataValidation.HealWorkbookStructure", $true)
    $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.EnsureEnrollmentInfrastructure")

    $wsDso = Get-WorksheetByName -Workbook $workbook -SheetName $sheetDso
    $wsPayments = Get-WorksheetByName -Workbook $workbook -SheetName $sheetPayments
    $wsEnrollment = Get-WorksheetByName -Workbook $workbook -SheetName $sheetEnrollment
    $wsEnrollmentForm = Get-WorksheetByName -Workbook $workbook -SheetName "EnrollmentForm"
    if ($null -eq $wsEnrollmentForm) {
        $wsEnrollmentForm = Get-WorksheetByName -Workbook $workbook -SheetName "Мастер_Зачисления"
    }

    Assert-True ($null -ne $wsDso) "Self-healing did not recreate sheet DSO."
    Assert-True ($null -ne $wsPayments) "Self-healing did not recreate payments sheet."
    Assert-True ($null -ne $wsEnrollment) "Self-healing did not recreate enrollment sheet."
    Assert-True ([string](Get-CellValue -Worksheet $wsPayments -RowNumber 1 -ColumnNumber 14) -ne "") "Expanded payments column model was not recreated."

    $shapeNames = @{}
    foreach ($shape in $wsEnrollment.Shapes) {
        $shapeNames[$shape.Name] = Get-ShapeText -Shape $shape
    }
    Assert-True (-not $shapeNames.ContainsKey("btnValidateEnrollment")) "Enrollment validate button should not overlap the sheet after self-healing."
    Assert-True (-not $shapeNames.ContainsKey("btnSuggestEnrollment")) "Enrollment suggestions button should not overlap the sheet after self-healing."
    Assert-True (-not $shapeNames.ContainsKey("btnTransferEnrollment")) "Enrollment transfer button should not overlap the sheet after self-healing."

    $paymentShapeNames = @{}
    foreach ($shape in $wsPayments.Shapes) {
        $paymentShapeNames[$shape.Name] = Get-ShapeText -Shape $shape
    }
    foreach ($legacyPaymentButton in @(
        "btnCreatePaymentPackage",
        "btnSelectPaymentEmployee",
        "btnPastePaymentNumbers",
        "btnFillSharedPaymentFields",
        "btnRecalcPaymentRows",
        "btnPreviewPaymentPackage",
        "btnExportPaymentsDocx",
        "btnOpenWorkbookFolder",
        "btnOpenPaymentsRibbonHint"
    )) {
        Assert-True (-not $paymentShapeNames.ContainsKey($legacyPaymentButton)) "Legacy payments sheet button '$legacyPaymentButton' should not overlap the sheet after self-healing."
    }

    if ($null -ne $wsEnrollmentForm) {
        $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.ClearEnrollmentForm")
        Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "fio" -FieldValue "Иванов Алексей Александрович"
        Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "personal_number" -FieldValue "Ю-110111"
        Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "position" -FieldValue "123/ТЕСТ"
        Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "order_date" -FieldValue "05.07.2026"
        Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "class_param" -FieldValue "1 класс"
        Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "passport_series" -FieldValue "1234"
        Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "passport_number" -FieldValue "567890"
        Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "passport_issuer" -FieldValue "Тестовый орган"
        Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "passport_issue_date" -FieldValue "01.01.2025"
        Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "inn" -FieldValue "123456789012"
        Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "snils" -FieldValue "123-456-789 00"
        Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "bank_account" -FieldValue "40817810000000000001"
        Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "bank_name" -FieldValue "СБЕР БАНК"
        Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "basis_section1" -FieldValue "Тестовое основание"
        $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentForm")
        Assert-True ((Get-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "preview_status") -ne "") "Enrollment master form did not calculate status."
        Assert-True ((Get-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "class_percent") -ne "") "Enrollment master form did not calculate class allowance percent."
    }

    Write-Output "2. Smart validation routing"
    $routeDso = $excel.Run("'$($workbook.Name)'!mdlRibbonHandlers.RunSmartValidationBySheetName", $sheetDso, $true)
    $routePayments = $excel.Run("'$($workbook.Name)'!mdlRibbonHandlers.RunSmartValidationBySheetName", $sheetPayments, $true)
    $routeEnrollment = $excel.Run("'$($workbook.Name)'!mdlRibbonHandlers.RunSmartValidationBySheetName", $sheetEnrollment, $true)
    Assert-Eq $routeDso "DSO" "Smart validation did not route DSO correctly."
    Assert-Eq $routePayments "PAYMENTS" "Smart validation did not route payments correctly."
    Assert-Eq $routeEnrollment "ENROLLMENT" "Smart validation did not route enrollment correctly."
    $ribbonLocalization = [string]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeRibbonLocalization")
    Assert-True (-not $ribbonLocalization.StartsWith("ERROR:")) "Ribbon localization probe failed: $ribbonLocalization"
    Assert-True ($ribbonLocalization.Contains((New-Text @(1054,1096,1080,1073,1082,1072,32,1087,1088,1080,32,1089,1086,1079,1076,1072,1085,1080,1080,32,1086,1089,1085,1086,1074,1085,1086,1075,1086,32,1087,1088,1080,1082,1072,1079,1072)))) "Ribbon main export error text is not loaded from Unicode-safe localization."
    Assert-True ($ribbonLocalization.Contains((New-Text @(1059,1084,1085,1072,1103,32,1087,1088,1086,1074,1077,1088,1082,1072)))) "Ribbon smart-validation title is not loaded from Unicode-safe localization."
    Assert-True ($ribbonLocalization.Contains((New-Text @(1054,1096,1080,1073,1082,1072,32,1087,1088,1080,32,1086,1090,1082,1088,1099,1090,1080,1080,32,1087,1072,1087,1082,1080,32,1082,1085,1080,1075,1080)))) "Ribbon open-workbook-folder error text is not loaded from Unicode-safe localization."
    $ribbonUiLocalization = [string]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeRibbonUiLocalization")
    Assert-True (-not $ribbonUiLocalization.StartsWith("ERROR:")) "Ribbon UI localization probe failed: $ribbonUiLocalization"
    Assert-True ($ribbonUiLocalization.Contains((New-Text @(1052,1072,1089,1090,1077,1088,32,1079,1072,1095,1080,1089,1083,1077,1085,1080,1103)))) "Ribbon enrollment button label is not loaded through localization callback."
    Assert-True ($ribbonUiLocalization.Contains((New-Text @(1048,1084,1087,1086,1088,1090,32,1080,1079,32,87,111,114,100)))) "Ribbon Word import screentip is not loaded through localization callback."
    Assert-True ($ribbonUiLocalization.Contains((New-Text @(1055,1088,1086,1074,1077,1088,1103,1077,1090,32,1044,56,57)))) "Ribbon D89 supertip is not loaded through localization callback."
    Assert-True ($ribbonUiLocalization.Contains((New-Text @(1069,1082,1089,1087,1086,1088,1090,32,1087,1086,32,79,114,100,101,114,68,114,97,102,116,73,100)))) "Ribbon export-by-OrderDraftId label is not loaded through localization callback."
    Assert-True ($ribbonUiLocalization.Contains((New-Text @(1042,1074,1077,1076,1080,1090,1077,32,79,114,100,101,114,68,114,97,102,116,73,100,32,1087,1088,1086,1077,1082,1090,1072,32,1087,1088,1080,1082,1072,1079,1072,58)))) "Ribbon export-by-OrderDraftId prompt is not loaded through localization."
    Assert-True ($ribbonUiLocalization.Contains((New-Text @(79,114,100,101,114,68,114,97,102,116,73,100,32,1085,1077,32,1091,1082,1072,1079,1072,1085,46,32,1069,1082,1089,1087,1086,1088,1090,32,1086,1090,1084,1077,1085,1105,1085,46)))) "Ribbon export-by-OrderDraftId empty-id message is not loaded through localization."
    Assert-True ($ribbonUiLocalization.Contains((New-Text @(1055,1072,1082,1077,1090,32,1087,1088,1080,1082,1072,1079,1072,32,1089,1092,1086,1088,1084,1080,1088,1086,1074,1072,1085)))) "Ribbon export-by-OrderDraftId success message is not loaded through localization."
    Assert-True ($ribbonUiLocalization.Contains((New-Text @(1054,1096,1080,1073,1082,1072,32,1101,1082,1089,1087,1086,1088,1090,1072,32,1087,1072,1082,1077,1090,1072)))) "Ribbon export-by-OrderDraftId error message is not loaded through localization."
    Assert-True ($ribbonUiLocalization -notlike "*Package exported*" -and $ribbonUiLocalization -notlike "*Export cancelled*" -and $ribbonUiLocalization -notlike "*Enrollment package export*") "Ribbon export-by-OrderDraftId leaked English fallback text."
    $settingsLocalization = [string]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeSettingsDiagnosticText")
    Assert-True (-not $settingsLocalization.StartsWith("ERROR:")) "Settings diagnostic localization probe failed: $settingsLocalization"
    Assert-True ($settingsLocalization.Contains((New-Text @(1053,1040,1057,1058,1056,1054,1049,1050,1048,32,1052,1040,1050,1056,1054,1057,1054,1042)))) "Settings diagnostic header is not loaded from Unicode-safe localization."
    Assert-True ($settingsLocalization.Contains((New-Text @(1055,1056,1054,1042,1045,1056,1050,1040)))) "Settings diagnostic template-check block is not loaded from Unicode-safe localization."
    Assert-True ($settingsLocalization.Contains((New-Text @(1057,1058,1040,1058,1059,1057,32,1040,1050,1058,1048,1042,1040,1062,1048,1048)))) "Settings diagnostic activation status label is not loaded from Unicode-safe localization."
    $coreLocalization = [string]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeCoreLocalization")
    Assert-Eq $coreLocalization ((New-Text @(1060,1086,1088,1084,1080,1088,1086,1074,1072,1090,1077,1083,1100,32,1087,1088,1080,1082,1072,1079,1086,1074)) + "|" + (New-Text @(1054,32,1087,1088,1086,1075,1088,1072,1084,1084,1077)) + "|" + (New-Text @(1057,1077,1088,1074,1080,1089,32,1083,1080,1094,1077,1085,1079,1080,1080)) + "|" + (New-Text @(1060,1072,1081,1083,1099,32,1083,1080,1094,1077,1085,1079,1080,1080))) "Core localization is not protected against mojibake."
    $legacyUiLocalization = [string]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeLegacyUiLocalization")
    Assert-Eq $legacyUiLocalization ((New-Text @(1055,1086,1080,1089,1082,32,1080,32,1087,1077,1088,1080,1086,1076,1099,32,1057,1042,1054)) + "|" + (New-Text @(1042,1099,1073,1086,1088,32,1089,1086,1090,1088,1091,1076,1085,1080,1082,1072)) + "|" + (New-Text @(1053,1077,1087,1088,1077,1088,1099,1074,1085,1099,1081,32,1087,1077,1088,1080,1086,1076,32,1089,1086,1089,1090,1072,1074,1083,1103,1077,1090,32,123,100,97,121,115,125,32,1076,1085,1077,1081,46,32,1055,1088,1086,1074,1077,1088,1100,1090,1077,32,1084,1077,1089,1103,1094,32,1080,32,1076,1072,1090,1099,46,32,1055,1088,1086,1076,1086,1083,1078,1080,1090,1100,32,1089,1086,1093,1088,1072,1085,1077,1085,1080,1077,63)) + "|" + (New-Text @(1057,1086,1089,1090,1086,1103,1085,1080,1077,32,1083,1080,1094,1077,1085,1079,1080,1080,58,32,123,115,116,97,116,117,115,125))) "Legacy UI localization is not protected against mojibake."

    Write-Output "3. Direct class qualification on payments sheet"
    $wsStaff = Get-WorksheetByName -Workbook $workbook -SheetName $sheetStaff
    $fio1 = [string](Get-CellValue -Worksheet $wsStaff -RowNumber 2 -ColumnNumber 4)
    $ln1 = [string](Get-CellValue -Worksheet $wsStaff -RowNumber 2 -ColumnNumber 2)
    $fio2 = [string](Get-CellValue -Worksheet $wsStaff -RowNumber 3 -ColumnNumber 4)
    $ln2 = [string](Get-CellValue -Worksheet $wsStaff -RowNumber 3 -ColumnNumber 2)
    $birthDate1 = [string](Get-CellText -Worksheet $wsStaff -RowNumber 2 -ColumnNumber 5)
    $citizenship1 = [string](Get-CellValue -Worksheet $wsStaff -RowNumber 2 -ColumnNumber 11)
    $serviceCategory1 = [string](Get-CellValue -Worksheet $wsStaff -RowNumber 2 -ColumnNumber 13)
    $contractKind1 = [string](Get-CellValue -Worksheet $wsStaff -RowNumber 2 -ColumnNumber 14)
    $vus1 = [string](Get-CellValue -Worksheet $wsStaff -RowNumber 2 -ColumnNumber 31)
    $tariff1 = [string](Get-CellText -Worksheet $wsStaff -RowNumber 2 -ColumnNumber 32)

    Write-Output "2a. Enrollment wizard staff selector flow"
    $wizardStaffLoad = [string]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeEnrollmentWizardStaffLoad", $ln1.Trim())
    Assert-True ($wizardStaffLoad -like ("staff|" + $fio1.Trim() + "|" + $ln1.Trim() + "|*")) "Enrollment wizard did not load employee data from the shared staff selector."
    Assert-True ($wizardStaffLoad -like ("*|" + $serviceCategory1.Trim() + "|" + $vus1.Trim() + "|" + $tariff1.Trim() + "|" + $birthDate1.Trim() + "|" + $citizenship1.Trim() + "|*")) "Enrollment wizard did not hydrate extended staff fields into the enrollment form. Snapshot: $wizardStaffLoad"
    $wizardInlineUiText = [string]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeEnrollmentWizardInlineSearchUiText")
    Assert-Eq $wizardInlineUiText "Введите ФИО, личный или табельный номер.|Загрузить из поиска" "Enrollment wizard inline-search UI is not fully localized."
    $wizardLayout = [string]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeEnrollmentWizardLayout")
    $wizardLayoutParts = $wizardLayout.Split('|')
    Assert-True ($wizardLayoutParts.Count -ge 25) "Enrollment wizard layout probe returned an invalid snapshot: $wizardLayout"
    $wizardHeight = [int]$wizardLayoutParts[0]
    $wizardWidth = [int]$wizardLayoutParts[1]
    $wizardPageHeight = [int]$wizardLayoutParts[2]
    $wizardPageWidth = [int]$wizardLayoutParts[3]
    $wizardFirstButtonTop = [int]$wizardLayoutParts[4]
    $wizardSecondButtonTop = [int]$wizardLayoutParts[5]
    $wizardCloseButtonBottom = [int]$wizardLayoutParts[6]
    $wizardSelectButtonRight = [int]$wizardLayoutParts[7]
    $wizardRequisitesNoteRight = [int]$wizardLayoutParts[10]
    $wizardRequisitesNoteBottom = [int]$wizardLayoutParts[11]
    $wizardExtraMonthly1Bottom = [int]$wizardLayoutParts[13]
    $wizardExtraMonthly2Top = [int]$wizardLayoutParts[14]
    $wizardExtraMonthly4Bottom = [int]$wizardLayoutParts[15]
    $wizardExtraOneTime1Top = [int]$wizardLayoutParts[16]
    $wizardExtraOneTime3Bottom = [int]$wizardLayoutParts[17]
    $wizardExtrasScrollHeight = [int]$wizardLayoutParts[18]
    $wizardExtraMonthlyName1Bottom = [int]$wizardLayoutParts[19]
    $wizardExtraMonthlyBasis1Top = [int]$wizardLayoutParts[20]
    $wizardExtraOneTimeName1Bottom = [int]$wizardLayoutParts[21]
    $wizardExtraOneTimeBasis1Top = [int]$wizardLayoutParts[22]
    $wizardExtraMonthlyShortLabel = [string]$wizardLayoutParts[23]
    $wizardExtraOneTimeShortLabel = [string]$wizardLayoutParts[24]
    Assert-True ($wizardHeight -ge 700 -and $wizardHeight -le 740) "Enrollment wizard should stay compact enough to fit vertically while keeping room for the page area."
    Assert-True ($wizardWidth -ge 840 -and $wizardWidth -le 900) "Enrollment wizard should stay compact enough to fit horizontally."
    Assert-True ($wizardPageHeight -ge 400) "Enrollment wizard page area is too small and may clip fields."
    Assert-True ($wizardPageWidth -ge 800 -and $wizardPageWidth -lt $wizardWidth) "Enrollment wizard page width is outside the compact window bounds."
    Assert-True ($wizardFirstButtonTop -gt (196 + $wizardPageHeight)) "Enrollment wizard bottom buttons still overlap the page area."
    Assert-True ($wizardSecondButtonTop -gt $wizardFirstButtonTop) "Enrollment wizard secondary action row should be below the primary action row."
    Assert-True ($wizardCloseButtonBottom -le $wizardHeight) "Enrollment wizard bottom buttons may be clipped below the window."
    Assert-True ($wizardSelectButtonRight -le ($wizardWidth - 20)) "Enrollment wizard staff-select button is clipped horizontally."
    Assert-True ([int]$wizardLayoutParts[9] -gt [int]$wizardLayoutParts[8]) "Enrollment wizard monthly controls still overlap."
    Assert-True ($wizardRequisitesNoteRight -le $wizardPageWidth) "Enrollment wizard requisites note field is clipped horizontally."
    Assert-True ($wizardRequisitesNoteBottom -le $wizardPageHeight) "Enrollment wizard requisites note field is clipped vertically."
    Assert-Eq $wizardLayoutParts[12] "Загрузить из поиска" "Enrollment wizard search-load button is not localized."
    Assert-True ($wizardExtraMonthly1Bottom -lt $wizardExtraMonthly2Top) "Enrollment wizard extra monthly controls still overlap."
    Assert-True ($wizardExtraMonthly4Bottom -lt $wizardExtraOneTime1Top) "Enrollment wizard extra one-time controls start before monthly controls end."
    Assert-True ($wizardExtraOneTime3Bottom -le $wizardExtrasScrollHeight) "Enrollment wizard extra payments page does not scroll far enough to the last field."
    Assert-True ($wizardExtraMonthlyName1Bottom -lt $wizardExtraMonthlyBasis1Top) "Enrollment wizard extra monthly basis overlaps the main row controls."
    Assert-True ($wizardExtraOneTimeName1Bottom -lt $wizardExtraOneTimeBasis1Top) "Enrollment wizard extra one-time basis overlaps the main row controls."
    Assert-Eq $wizardExtraMonthlyShortLabel "Ежемес. #1: вид" "Enrollment wizard extra monthly label should be compact and localized."
    Assert-Eq $wizardExtraOneTimeShortLabel "Разовая #1: вид" "Enrollment wizard extra one-time label should be compact and localized."
    $wizardInlineSearch = [string]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeEnrollmentWizardInlineSearch", $ln1.Trim())
    Assert-True ($wizardInlineSearch -like ("1|" + $fio1.Trim() + "|" + $ln1.Trim())) "Enrollment wizard inline search did not find the employee inside the master form."

    $wsPayments.Cells(2, 2).Value = $paymentClassQualification
    $wsPayments.Cells(2, 3).Value = $fio1.Trim()
    $wsPayments.Cells(2, 4).Value = $ln1.Trim()
    $wsPayments.Cells(2, 6).Value = "Order 1"
    $wsPayments.Cells(2, 9).Value = "1"
    $excel.Run("'$($workbook.Name)'!mdlPaymentPackageSupport.EnrichPaymentRow", $wsPayments, 2)
    Assert-Eq (Get-CellText -Worksheet $wsPayments -RowNumber 2 -ColumnNumber 5) "20%" "Direct class qualification did not calculate 20%."

    Write-Output "3a. Number-list import to grouped package"
    $wsPayments.Cells.Clear()
    $excel.Run("'$($workbook.Name)'!mdlDataValidation.HealWorkbookStructure", $true)
    $wsPayments = Get-WorksheetByName -Workbook $workbook -SheetName $sheetPayments

    $sharedImportBasis = "Shared import basis"
    $wsPayments.Cells(2, 2).Value = $paymentClassQualification
    $wsPayments.Cells(2, 6).Value = $sharedImportBasis
    $wsPayments.Cells(2, 7).Value = "PKG-IMPORT"
    $wsPayments.Cells(2, 8).Value = "LIST"
    $wsPayments.Cells(2, 9).Value = "1"
    $wsPayments.Cells(2, 10).Value = $sharedImportBasis
    $wsPayments.Cells(2, 11).Value = $groupExportYes

    $importText = $ln1.Trim() + "`r`n" + $ln2.Trim()
    $importedCount = $excel.Run("'$($workbook.Name)'!mdlPaymentPackageSupport.ImportEmployeesFromTextToSheet", $sheetPayments, 2, $importText)
    Assert-Eq $importedCount 2 "Number-list import did not create expected payment rows."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsPayments -RowNumber 2 -ColumnNumber 3)).Trim() $fio1.Trim() "Number-list import did not resolve FIO for the first row."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsPayments -RowNumber 3 -ColumnNumber 3)).Trim() $fio2.Trim() "Number-list import did not resolve FIO for the second row."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsPayments -RowNumber 3 -ColumnNumber 2)).Trim() $paymentClassQualification "Imported row did not inherit payment type from the anchor row."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsPayments -RowNumber 3 -ColumnNumber 7)).Trim() "PKG-IMPORT" "Imported row did not inherit package id."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsPayments -RowNumber 3 -ColumnNumber 10)).Trim() $sharedImportBasis "Imported row did not inherit shared basis."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsPayments -RowNumber 3 -ColumnNumber 11)).Trim() $groupExportYes "Imported row did not inherit grouped export flag."
    Assert-Eq (Get-CellText -Worksheet $wsPayments -RowNumber 3 -ColumnNumber 5) "20%" "Imported row did not recalculate class qualification amount from inherited parameter."

    Write-Output "3b. Staff-selection import to grouped package"
    $wsPayments.Cells.Clear()
    $excel.Run("'$($workbook.Name)'!mdlDataValidation.HealWorkbookStructure", $true)
    $wsPayments = Get-WorksheetByName -Workbook $workbook -SheetName $sheetPayments

    $sharedStaffImportBasis = "Shared staff import basis"
    $wsPayments.Cells(2, 2).Value = $paymentClassQualification
    $wsPayments.Cells(2, 6).Value = $sharedStaffImportBasis
    $wsPayments.Cells(2, 7).Value = "PKG-STAFF-IMPORT"
    $wsPayments.Cells(2, 8).Value = "LIST"
    $wsPayments.Cells(2, 9).Value = "1"
    $wsPayments.Cells(2, 10).Value = $sharedStaffImportBasis
    $wsPayments.Cells(2, 11).Value = $groupExportYes

    $staffImportedCount = $excel.Run("'$($workbook.Name)'!mdlPaymentPackageSupport.ImportEmployeesFromStaffRange", $sheetStaff, $sheetPayments, "A2:A3", 2)
    Assert-Eq $staffImportedCount 2 "Staff-selection import did not create expected payment rows."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsPayments -RowNumber 2 -ColumnNumber 3)).Trim() $fio1.Trim() "Staff-selection import did not copy the first employee."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsPayments -RowNumber 3 -ColumnNumber 3)).Trim() $fio2.Trim() "Staff-selection import did not copy the second employee."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsPayments -RowNumber 3 -ColumnNumber 2)).Trim() $paymentClassQualification "Staff-selection import did not inherit payment type from anchor row."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsPayments -RowNumber 3 -ColumnNumber 7)).Trim() "PKG-STAFF-IMPORT" "Staff-selection import did not inherit package id."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsPayments -RowNumber 3 -ColumnNumber 10)).Trim() $sharedStaffImportBasis "Staff-selection import did not inherit shared basis."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsPayments -RowNumber 3 -ColumnNumber 11)).Trim() $groupExportYes "Staff-selection import did not inherit grouped export flag."
    Assert-Eq (Get-CellText -Worksheet $wsPayments -RowNumber 3 -ColumnNumber 5) "20%" "Staff-selection import did not recalculate class qualification amount from inherited parameter."

    Write-Output "3c. Twenty-person grouped package"
    $wsPayments.Cells.Clear()
    $excel.Run("'$($workbook.Name)'!mdlDataValidation.HealWorkbookStructure", $true)
    $wsPayments = Get-WorksheetByName -Workbook $workbook -SheetName $sheetPayments

    $sharedTwentyBasis = "Shared basis for twenty"
    $wsPayments.Cells(2, 2).Value = $paymentClassQualification
    $wsPayments.Cells(2, 6).Value = $sharedTwentyBasis
    $wsPayments.Cells(2, 7).Value = "PKG-TWENTY"
    $wsPayments.Cells(2, 8).Value = "LIST"
    $wsPayments.Cells(2, 9).Value = "1"
    $wsPayments.Cells(2, 10).Value = $sharedTwentyBasis
    $wsPayments.Cells(2, 11).Value = $groupExportYes

    $twentyImportedCount = $excel.Run("'$($workbook.Name)'!mdlPaymentPackageSupport.ImportEmployeesFromStaffRange", $sheetStaff, $sheetPayments, "A2:A21", 2)
    Assert-Eq $twentyImportedCount 20 "Twenty-person package import did not copy 20 employees."
    Assert-Eq (Get-LastUsedRow -Worksheet $wsPayments -ColumnNumber 4) 21 "Twenty-person package import did not create the expected row count."

    $allTwentyPackageRowsMatched = $true
    for ($row = 2; $row -le 21; $row++) {
        if (([string](Get-CellValue -Worksheet $wsPayments -RowNumber $row -ColumnNumber 7)).Trim() -ne "PKG-TWENTY") {
            $allTwentyPackageRowsMatched = $false
            break
        }
    }
    Assert-True $allTwentyPackageRowsMatched "Not all rows in the twenty-person package inherited the same package id."

    Write-Output "4. Enrollment generation"
    $wsRef = Get-WorksheetByName -Workbook $workbook -SheetName $sheetPaymentTypes
    $refLastRow = Get-LastUsedRow -Worksheet $wsRef -ColumnNumber 1
    $refTestRow = $refLastRow + 1
    $wsRef.Cells($refTestRow, 1).Value = "Test standard payment"
    $wsRef.Cells($refTestRow, 2).Value = "TEST_STD"
    $wsRef.Cells($refTestRow, 3).Value = ""
    $wsRef.Cells($refTestRow, 6).Value = "STANDARD"
    $wsRef.Cells($refTestRow, 7).Value = "ENROLLMENT"
    $wsRef.Cells($refTestRow, 10).Value = "ORDER"

    $expectedPosition = [string](Get-CellValue -Worksheet $wsStaff -RowNumber 2 -ColumnNumber 30)
    $expectedUnit = [string](Get-CellValue -Worksheet $wsStaff -RowNumber 2 -ColumnNumber 37)

    $wsEnrollment.Cells(2, 3).Value = $ln1.Trim()
    $wsEnrollment.Cells(2, 7).Value = "01.07.2026"
    $wsEnrollment.Cells(2, 8).Value = "100"
    $wsEnrollment.Cells(2, 9).Value = "01.07.2026"
    $wsEnrollment.Cells(2, 10).Value = "02.07.2026"
    $wsEnrollment.Cells(2, 11).Value = "03.07.2026"
    $wsEnrollment.Cells(2, 13).Value = "report 200"
    $wsEnrollment.Cells(2, 14).Value = "assignment 300"
    $wsEnrollment.Cells(2, 15).Value = "1"
    $wsEnrollment.Cells(2, 34).Value = "2"
    $wsEnrollment.Cells(2, 35).Value = "14647"
    Set-EnrollmentRequiredFields -Worksheet $wsEnrollment -RowNumber 2

    $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentRowDirect", $sheetEnrollment, 2)

    Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 2)).Trim() $fio1.Trim() "Enrollment did not resolve FIO from Staff."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 5)).Trim() $expectedPosition.Trim() "Enrollment did not resolve position from Staff."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 6)).Trim() $expectedUnit.Trim() "Enrollment did not resolve unit from Staff."
    Assert-True (([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 18)) -like "*$paymentStdDuty*") "Enrollment did not determine the standard position allowance."
    Assert-True (([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 18)) -like "*$previewStdSpecial*") "Enrollment did not determine the preview marker for the special-conditions allowance."

    $wsEnrollment.Rows(2).Copy($wsEnrollment.Rows(3))
    $wsEnrollment.Cells(3, 5).Value = "Заведомо неверная должность"
    $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentRowDirect", $sheetEnrollment, 3)
    Assert-Eq ([int](Get-CellValue -Worksheet $wsEnrollment -RowNumber 3 -ColumnNumber 24)) 1 "Enrollment staff mismatch should be highlighted as a warning."
    $staffMismatchIssues = [string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 3 -ColumnNumber 25)
    Assert-True ($staffMismatchIssues -like "*Штат*") "Enrollment staff mismatch warning did not mention the Staff sheet."
    Assert-True ($staffMismatchIssues -like "*Штатная должность*") "Enrollment staff mismatch warning did not mention the mismatched position field."
    Assert-True ($staffMismatchIssues -like "*Заведомо неверная должность*") "Enrollment staff mismatch warning did not include the entered position."
    $wsEnrollment.Rows(3).ClearContents()

    $paymentsBeforeGenerateLastRow = Get-LastUsedRow -Worksheet $wsPayments -ColumnNumber 4
    $createdRows = $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.GeneratePaymentsFromEnrollmentRowDirect", 2)
    Assert-True ($createdRows -ge 3) "Enrollment did not generate the expected built-in payment rows."

    $generatedPaymentTypes = @()
    $paymentsLastRow = Get-LastUsedRow -Worksheet $wsPayments -ColumnNumber 4
    for ($row = [Math]::Max(2, $paymentsBeforeGenerateLastRow + 1); $row -le $paymentsLastRow; $row++) {
        $generatedPaymentTypes += ,([string](Get-CellValue -Worksheet $wsPayments -RowNumber $row -ColumnNumber 2))
    }
    Assert-True ($generatedPaymentTypes -contains $paymentClassQualification) "Enrollment did not generate class qualification row."
    Assert-True ($generatedPaymentTypes -contains $paymentStdDuty) "Enrollment did not generate the standard position allowance row."
    Assert-True ($generatedPaymentTypes -contains $paymentStdSpecial) "Enrollment did not generate the standard special-conditions row."

    Write-Output "5. Export split by package id"
    $wsPayments.Cells.Clear()
    $excel.Run("'$($workbook.Name)'!mdlDataValidation.HealWorkbookStructure", $true)
    $wsPayments = Get-WorksheetByName -Workbook $workbook -SheetName $sheetPayments

    $wsPayments.Cells(2, 2).Value = $paymentClassQualification
    $wsPayments.Cells(2, 3).Value = $fio1.Trim()
    $wsPayments.Cells(2, 4).Value = $ln1.Trim()
    $wsPayments.Cells(2, 6).Value = "Order 1"
    $wsPayments.Cells(2, 7).Value = "PKG-A"
    $wsPayments.Cells(2, 8).Value = "LIST"
    $wsPayments.Cells(2, 9).Value = "1"
    $wsPayments.Cells(2, 10).Value = "Order 1"
    $wsPayments.Cells(2, 11).Value = $groupExportYes
    $excel.Run("'$($workbook.Name)'!mdlPaymentPackageSupport.EnrichPaymentRow", $wsPayments, 2)

    $wsPayments.Cells(3, 2).Value = $paymentClassQualification
    $wsPayments.Cells(3, 3).Value = $fio2.Trim()
    $wsPayments.Cells(3, 4).Value = $ln2.Trim()
    $wsPayments.Cells(3, 6).Value = "Order 2"
    $wsPayments.Cells(3, 7).Value = "PKG-B"
    $wsPayments.Cells(3, 8).Value = "LIST"
    $wsPayments.Cells(3, 9).Value = "2"
    $wsPayments.Cells(3, 10).Value = "Order 2"
    $wsPayments.Cells(3, 11).Value = $groupExportYes
    $excel.Run("'$($workbook.Name)'!mdlPaymentPackageSupport.EnrichPaymentRow", $wsPayments, 3)

    $exportedCount = $excel.Run("'$($workbook.Name)'!mdlUniversalPaymentExport.ExportPaymentsWithoutPeriodsCore", $false)
    Assert-Eq $exportedCount 2 "Export did not create separate outputs for different packages."

    $docA = Get-ChildItem -LiteralPath $testDir -Filter "*.docx" | Where-Object { $_.Name -like "*CLASS_QUAL*PKG-A.docx" } | Select-Object -First 1
    $docB = Get-ChildItem -LiteralPath $testDir -Filter "*.docx" | Where-Object { $_.Name -like "*CLASS_QUAL*PKG-B.docx" } | Select-Object -First 1
    Assert-True ($null -ne $docA) "Export file for package PKG-A was not created."
    Assert-True ($null -ne $docB) "Export file for package PKG-B was not created."

    Write-Output "6. Grouped export shared basis"
    $refListRow = (Get-LastUsedRow -Worksheet $wsRef -ColumnNumber 1) + 1
    $wsRef.Cells($refListRow, 1).Value = "Test grouped export"
    $wsRef.Cells($refListRow, 2).Value = "TEST_GROUPED"
    $wsRef.Cells($refListRow, 3).Value = ""

    Get-ChildItem -LiteralPath $testDir -Filter "*.docx" | Where-Object { $_.Name -like "Шаблон_*" } | Remove-Item -Force

    $wsPayments.Cells.Clear()
    $excel.Run("'$($workbook.Name)'!mdlDataValidation.HealWorkbookStructure", $true)
    $wsPayments = Get-WorksheetByName -Workbook $workbook -SheetName $sheetPayments

    $sharedExportBasis = "Shared export basis"
    $wsPayments.Cells(2, 2).Value = "Test grouped export"
    $wsPayments.Cells(2, 3).Value = $fio1.Trim()
    $wsPayments.Cells(2, 4).Value = $ln1.Trim()
    $wsPayments.Cells(2, 5).Value = "10%"
    $wsPayments.Cells(2, 6).Value = $sharedExportBasis
    $wsPayments.Cells(2, 7).Value = "PKG-LIST-DOCX"
    $wsPayments.Cells(2, 8).Value = "LIST"
    $wsPayments.Cells(2, 10).Value = $sharedExportBasis
    $wsPayments.Cells(2, 11).Value = $groupExportYes

    $wsPayments.Cells(3, 2).Value = "Test grouped export"
    $wsPayments.Cells(3, 3).Value = $fio2.Trim()
    $wsPayments.Cells(3, 4).Value = $ln2.Trim()
    $wsPayments.Cells(3, 5).Value = "20%"
    $wsPayments.Cells(3, 6).Value = $sharedExportBasis
    $wsPayments.Cells(3, 7).Value = "PKG-LIST-DOCX"
    $wsPayments.Cells(3, 8).Value = "LIST"
    $wsPayments.Cells(3, 10).Value = $sharedExportBasis
    $wsPayments.Cells(3, 11).Value = $groupExportYes

    $exportedCount = $excel.Run("'$($workbook.Name)'!mdlUniversalPaymentExport.ExportPaymentsWithoutPeriodsCore", $false)
    Assert-Eq $exportedCount 1 "Grouped list export should create one docx package."

    $listDoc = Get-ChildItem -LiteralPath $testDir -Filter "*.docx" | Where-Object { $_.Name -like "*TEST_GROUPED*PKG-LIST-DOCX.docx" } | Select-Object -First 1
    Assert-True ($null -ne $listDoc) "Grouped export docx was not created."
    $listDocText = Get-DocxText -Path $listDoc.FullName
    Assert-Eq (Get-MatchCount -Text $listDocText -Pattern $sharedExportBasis) 1 "Grouped export docx repeated shared basis more than once."

    Write-Output "7. Grouped preview"
    $wsPayments.Cells.Clear()
    $excel.Run("'$($workbook.Name)'!mdlDataValidation.HealWorkbookStructure", $true)
    $wsPayments = Get-WorksheetByName -Workbook $workbook -SheetName $sheetPayments

    $sharedPreviewBasis = "Shared preview basis"
    $wsPayments.Cells(2, 2).Value = $paymentClassQualification
    $wsPayments.Cells(2, 3).Value = $fio1.Trim()
    $wsPayments.Cells(2, 4).Value = $ln1.Trim()
    $wsPayments.Cells(2, 7).Value = "PKG-PREVIEW"
    $wsPayments.Cells(2, 8).Value = "LIST"
    $wsPayments.Cells(2, 9).Value = "1"
    $wsPayments.Cells(2, 10).Value = $sharedPreviewBasis
    $wsPayments.Cells(2, 11).Value = $groupExportYes
    $excel.Run("'$($workbook.Name)'!mdlPaymentPackageSupport.EnrichPaymentRow", $wsPayments, 2)

    $wsPayments.Cells(3, 2).Value = $paymentClassQualification
    $wsPayments.Cells(3, 3).Value = $fio2.Trim()
    $wsPayments.Cells(3, 4).Value = $ln2.Trim()
    $wsPayments.Cells(3, 7).Value = "PKG-PREVIEW"
    $wsPayments.Cells(3, 8).Value = "LIST"
    $wsPayments.Cells(3, 9).Value = "2"
    $wsPayments.Cells(3, 10).Value = $sharedPreviewBasis
    $wsPayments.Cells(3, 11).Value = $groupExportYes
    $excel.Run("'$($workbook.Name)'!mdlPaymentPackageSupport.EnrichPaymentRow", $wsPayments, 3)

    $excel.Run("'$($workbook.Name)'!mdlPaymentPackageSupport.PreparePreviewForRange", $sheetPayments, "A2:M3")
    $wsPreview = Get-WorksheetByName -Workbook $workbook -SheetName $previewSheetName
    Assert-True ($null -ne $wsPreview) "Preview sheet was not created."
    $previewText = [string](Get-CellValue -Worksheet $wsPreview -RowNumber 1 -ColumnNumber 1)
    Assert-True ($previewText -like "*PKG-PREVIEW*") "Preview does not show package identifier."
    Assert-Eq (Get-MatchCount -Text $previewText -Pattern $sharedPreviewBasis) 1 "Grouped preview repeated the shared basis more than once."

    Write-Output "8. Payments warning validation"
    $wsPayments.Cells.Clear()
    $excel.Run("'$($workbook.Name)'!mdlDataValidation.HealWorkbookStructure", $true)
    $wsPayments = Get-WorksheetByName -Workbook $workbook -SheetName $sheetPayments

    $wsPayments.Cells(2, 2).Value = "Test standard payment"
    $wsPayments.Cells(2, 3).Value = $fio1.Trim()
    $wsPayments.Cells(2, 4).Value = $ln1.Trim()
    $wsPayments.Cells(2, 5).Value = "0"
    $wsPayments.Cells(2, 6).Value = "Order warning zero"

    $wsPayments.Cells(3, 2).Value = "Test standard payment"
    $wsPayments.Cells(3, 3).Value = $fio2.Trim()
    $wsPayments.Cells(3, 4).Value = $ln2.Trim()
    $wsPayments.Cells(3, 5).Value = "100"
    $wsPayments.Cells(3, 6).Value = "Order warning package"
    $wsPayments.Cells(3, 7).Value = "PKG-WARN"
    $wsPayments.Cells(3, 8).Value = "LIST"
    $wsPayments.Cells(3, 11).Value = $groupExportYes

    $wsPayments.Cells(4, 2).Value = "Test standard payment"
    $wsPayments.Cells(4, 3).Value = "Тестов Тест Тестович"
    $wsPayments.Cells(4, 4).Value = "ZZ-404"
    $wsPayments.Cells(4, 5).Value = "150"
    $wsPayments.Cells(4, 6).Value = "Order warning missing staff"

    $excel.Run("'$($workbook.Name)'!mdlPaymentValidation.ValidatePaymentsWithoutPeriods", $true)

    Assert-Eq ([string](Get-CellValue -Worksheet $wsPayments -RowNumber 2 -ColumnNumber 13)).Trim() $statusWarning "Zero amount row did not receive warning status."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsPayments -RowNumber 3 -ColumnNumber 13)).Trim() $statusWarning "Grouped row without shared foundation did not receive warning status."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsPayments -RowNumber 4 -ColumnNumber 13)).Trim() $statusWarning "Missing staff row did not receive warning status."

    Write-Output "8a. Crew validation uses VUS-position reference"
    $wsVusCrew = Get-WorksheetByName -Workbook $workbook -SheetName $sheetVusCrew
    Assert-True ($null -ne $wsVusCrew) "VUS crew reference sheet was not created by workbook healing."

    $crewRefPersonalNumber = "CREW-REF-001"
    $crewRefFio = "Crew Reference Employee"
    $crewRefVus = "500543"
    $crewRefPosition = New-Text @(1088,1072,1076,1080,1086,1083,1086,1082,1072,1094,1080,1086,1085,1085,1099,1081,32,1089,1087,1077,1094,1080,1072,1083,1080,1089,1090)

    $crewStaffRow = (Get-LastUsedRow -Worksheet $wsStaff -ColumnNumber 2) + 1
    $wsStaff.Cells($crewStaffRow, 2).Value = $crewRefPersonalNumber
    $wsStaff.Cells($crewStaffRow, 4).Value = $crewRefFio
    $wsStaff.Cells($crewStaffRow, 30).Value = $crewRefPosition
    $wsStaff.Cells($crewStaffRow, 31).Value = $crewRefVus

    $crewReferenceRow = (Get-LastUsedRow -Worksheet $wsVusCrew -ColumnNumber 1) + 1
    $wsVusCrew.Cells($crewReferenceRow, 1).Value = $crewRefVus
    $wsVusCrew.Cells($crewReferenceRow, 2).Value = $crewRefPosition

    $wsPayments.Cells.Clear()
    $excel.Run("'$($workbook.Name)'!mdlDataValidation.HealWorkbookStructure", $true)
    $wsPayments = Get-WorksheetByName -Workbook $workbook -SheetName $sheetPayments
    $wsPayments.Cells(2, 2).Value = $paymentCrew
    $wsPayments.Cells(2, 3).Value = $crewRefFio
    $wsPayments.Cells(2, 4).Value = $crewRefPersonalNumber
    $wsPayments.Cells(2, 5).Value = "100"
    $wsPayments.Cells(2, 6).Value = "VUS-position reference basis"

    $excel.Run("'$($workbook.Name)'!mdlPaymentValidation.ValidatePaymentsWithoutPeriods", $true)
    Assert-Eq ([string](Get-CellValue -Worksheet $wsPayments -RowNumber 2 -ColumnNumber 13)).Trim() $paymentStatusOk "Crew validation should allow a VUS-position reference match even when the position has no crew keyword."

    Write-Output "9. Payment eligibility severity"
    $refWarnPayRow = (Get-LastUsedRow -Worksheet $wsRef -ColumnNumber 1) + 1
    $wsRef.Cells($refWarnPayRow, 1).Value = "Test warning export"
    $wsRef.Cells($refWarnPayRow, 2).Value = "TEST_WARN_PAY"
    $wsRef.Cells($refWarnPayRow, 3).Value = ""
    $wsRef.Cells($refWarnPayRow, 13).Value = "PARAM_REQUIRED"
    $wsRef.Cells($refWarnPayRow, 14).Value = "WARNING"

    $refBlockPayRow = (Get-LastUsedRow -Worksheet $wsRef -ColumnNumber 1) + 1
    $wsRef.Cells($refBlockPayRow, 1).Value = "Test blocked export"
    $wsRef.Cells($refBlockPayRow, 2).Value = "TEST_BLOCK_PAY"
    $wsRef.Cells($refBlockPayRow, 3).Value = ""
    $wsRef.Cells($refBlockPayRow, 13).Value = "PARAM_REQUIRED"
    $wsRef.Cells($refBlockPayRow, 14).Value = "BLOCKED"

    $wsPayments.Cells.Clear()
    $excel.Run("'$($workbook.Name)'!mdlDataValidation.HealWorkbookStructure", $true)
    $wsPayments = Get-WorksheetByName -Workbook $workbook -SheetName $sheetPayments

    $wsPayments.Cells(2, 2).Value = "Test warning export"
    $wsPayments.Cells(2, 3).Value = $fio1.Trim()
    $wsPayments.Cells(2, 4).Value = $ln1.Trim()
    $wsPayments.Cells(2, 5).Value = "100"
    $wsPayments.Cells(2, 6).Value = "Warn export basis"
    $wsPayments.Cells(2, 7).Value = "PKG-WARN-EXPORT"
    $wsPayments.Cells(2, 8).Value = "LIST"
    $wsPayments.Cells(2, 10).Value = "Warn export basis"
    $wsPayments.Cells(2, 11).Value = $groupExportYes

    $wsPayments.Cells(3, 2).Value = "Test blocked export"
    $wsPayments.Cells(3, 3).Value = $fio2.Trim()
    $wsPayments.Cells(3, 4).Value = $ln2.Trim()
    $wsPayments.Cells(3, 5).Value = "100"
    $wsPayments.Cells(3, 6).Value = "Blocked export basis"
    $wsPayments.Cells(3, 7).Value = "PKG-BLOCK-EXPORT"
    $wsPayments.Cells(3, 8).Value = "LIST"
    $wsPayments.Cells(3, 10).Value = "Blocked export basis"
    $wsPayments.Cells(3, 11).Value = $groupExportYes

    $excel.Run("'$($workbook.Name)'!mdlPaymentValidation.ValidatePaymentsWithoutPeriods", $true)
    Assert-Eq ([string](Get-CellValue -Worksheet $wsPayments -RowNumber 2 -ColumnNumber 13)).Trim() $statusWarning "Warning-severity payment row did not remain a warning."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsPayments -RowNumber 3 -ColumnNumber 13)).Trim() $statusBlocked "Blocked-severity payment row did not become blocked."

    $exportedCount = $excel.Run("'$($workbook.Name)'!mdlUniversalPaymentExport.ExportPaymentsWithoutPeriodsCore", $false)
    Assert-Eq $exportedCount 1 "Blocked payment rows were not excluded from export."

    $warnDoc = Get-ChildItem -LiteralPath $testDir -Filter "*.docx" | Where-Object { $_.Name -like "*TEST_WARN_PAY*PKG-WARN-EXPORT.docx" } | Select-Object -First 1
    $blockDoc = Get-ChildItem -LiteralPath $testDir -Filter "*.docx" | Where-Object { $_.Name -like "*TEST_BLOCK_PAY*PKG-BLOCK-EXPORT.docx" } | Select-Object -First 1
    Assert-True ($null -ne $warnDoc) "Warning-severity payment row was not exported."
    Assert-True ($null -eq $blockDoc) "Blocked-severity payment row should not create an export file."

    Write-Output "9b. Advanced payment eligibility rules"
    $refPositionPayRow = (Get-LastUsedRow -Worksheet $wsRef -ColumnNumber 1) + 1
    $wsRef.Cells($refPositionPayRow, 1).Value = "Test position gated"
    $wsRef.Cells($refPositionPayRow, 2).Value = "TEST_POS_PAY"
    $wsRef.Cells($refPositionPayRow, 3).Value = ""
    $wsRef.Cells($refPositionPayRow, 13).Value = "POSITION_KEYWORDS"
    $wsRef.Cells($refPositionPayRow, 14).Value = "WARNING"
    $wsRef.Cells($refPositionPayRow, 17).Value = $expectedPosition

    $refFoundationPayRow = (Get-LastUsedRow -Worksheet $wsRef -ColumnNumber 1) + 1
    $wsRef.Cells($refFoundationPayRow, 1).Value = "Test foundation gated"
    $wsRef.Cells($refFoundationPayRow, 2).Value = "TEST_FOUND_PAY"
    $wsRef.Cells($refFoundationPayRow, 3).Value = ""
    $wsRef.Cells($refFoundationPayRow, 13).Value = "FOUNDATION_KEYWORDS"
    $wsRef.Cells($refFoundationPayRow, 14).Value = "BLOCKED"
    $wsRef.Cells($refFoundationPayRow, 18).Value = "допуск;секрет"

    $wsPayments.Cells.Clear()
    $excel.Run("'$($workbook.Name)'!mdlDataValidation.HealWorkbookStructure", $true)
    $wsPayments = Get-WorksheetByName -Workbook $workbook -SheetName $sheetPayments

    $wsPayments.Cells(2, 2).Value = "Test position gated"
    $wsPayments.Cells(2, 3).Value = $fio1.Trim()
    $wsPayments.Cells(2, 4).Value = $ln1.Trim()
    $wsPayments.Cells(2, 5).Value = "100"
    $wsPayments.Cells(2, 6).Value = "Position basis"

    $wsPayments.Cells(3, 2).Value = "Test foundation gated"
    $wsPayments.Cells(3, 3).Value = $fio2.Trim()
    $wsPayments.Cells(3, 4).Value = $ln2.Trim()
    $wsPayments.Cells(3, 5).Value = "100"
    $wsPayments.Cells(3, 6).Value = "обычное основание"

    $excel.Run("'$($workbook.Name)'!mdlPaymentValidation.ValidatePaymentsWithoutPeriods", $true)
    Assert-Eq ([string](Get-CellValue -Worksheet $wsPayments -RowNumber 2 -ColumnNumber 13)).Trim() $paymentStatusOk "Position-keyword eligibility should allow a matching employee."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsPayments -RowNumber 3 -ColumnNumber 13)).Trim() $statusBlocked "Foundation-keyword eligibility should block a row with invalid basis."

    Write-Output "10. Enrollment V2 preview and transfer"
    $wsEnrollment.Cells.Clear()
    $excel.Run("'$($workbook.Name)'!mdlDataValidation.HealWorkbookStructure", $true)
    $wsEnrollment = Get-WorksheetByName -Workbook $workbook -SheetName $sheetEnrollment
    $wsPayments.Cells.Clear()
    $excel.Run("'$($workbook.Name)'!mdlDataValidation.HealWorkbookStructure", $true)
    $wsPayments = Get-WorksheetByName -Workbook $workbook -SheetName $sheetPayments
    $wsSettings = Get-WorksheetByName -Workbook $workbook -SheetName "Настройки"
    Assert-True ($null -ne $wsSettings) "Settings sheet was not created."

    $wsEnrollment.Cells(2, 3).Value = $ln1.Trim()
    $wsEnrollment.Cells(2, 7).Value = "01.07.2026"
    $wsEnrollment.Cells(2, 8).Value = "500"
    $wsEnrollment.Cells(2, 10).Value = "04.07.2026"
    $wsEnrollment.Cells(2, 11).Value = "02.07.2026"
    $wsEnrollment.Cells(2, 13).Value = "report manual"
    $wsEnrollment.Cells(2, 14).Value = "assignment manual"
    $wsEnrollment.Cells(2, 15).Value = "1 класс"
    $wsEnrollment.Cells(2, 16).Value = "1"
    $wsEnrollment.Cells(2, 17).Value = "2"
    $wsEnrollment.Cells(2, 9).Value = "60"
    $wsEnrollment.Cells(2, 30).Value = "контракт"
    $wsEnrollment.Cells(2, 31).Value = "контракт от 21.09.2022"
    $wsEnrollment.Cells(2, 34).Value = "2"
    $wsEnrollment.Cells(2, 35).Value = "14647"
    $wsEnrollment.Cells(2, 38).Value = "пункт отбора"
    $wsEnrollment.Cells(2, 39).Value = "1312"
    $wsEnrollment.Cells(2, 40).Value = "27.06.2026"
    $wsEnrollment.Cells(2, 41).Value = "881"
    $wsEnrollment.Cells(2, 42).Value = "30.06.2026"
    $wsEnrollment.Cells(2, 43).Value = "02.07.2026"
    $wsEnrollment.Cells(2, 44).Value = "02.07.2026"
    $wsEnrollment.Cells(2, 45).Value = "02.07.2026"
    $wsEnrollment.Cells(2, 61).Value = "YES"
    $wsEnrollment.Cells(2, 62).Value = "1.5"
    $wsEnrollment.Cells(2, 63).Value = "льготная выслуга"
    $wsEnrollment.Cells(2, 64).Value = "YES"
    $wsEnrollment.Cells(2, 65).Value = "25"
    $wsEnrollment.Cells(2, 66).Value = "02.07.2026"
    $wsEnrollment.Cells(2, 67).Value = "31.12.2026"
    $wsEnrollment.Cells(2, 68).Value = "премирование"
    $wsEnrollment.Cells(2, 110).Value = "Надбавка за ученую степень"
    $wsEnrollment.Cells(2, 111).Value = "кандидат наук"
    $wsEnrollment.Cells(2, 112).Value = "15%"
    $wsEnrollment.Cells(2, 113).Value = "02.07.2026"
    $wsEnrollment.Cells(2, 114).Value = "диплом о присуждении степени"
    $wsEnrollment.Cells(2, 115).Value = "YES"
    $wsEnrollment.Cells(2, 134).Value = "Компенсация найма жилья"
    $wsEnrollment.Cells(2, 135).Value = "12000"
    $wsEnrollment.Cells(2, 136).Value = "02.07.2026"
    $wsEnrollment.Cells(2, 137).Value = "договор найма жилого помещения"
    $wsEnrollment.Cells(2, 138).Value = "YES"
    Set-EnrollmentRequiredFields -Worksheet $wsEnrollment -RowNumber 2

    $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentRowDirect", $sheetEnrollment, 2)

    Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 20)).Trim() $statusReady "Enrollment V2 row should become ready after all required fields are filled."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 23)).Trim() "YES" "Enrollment V2 row should be Word-ready."
    Assert-True (([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 18)) -like "*$paymentStdDuty*") "Enrollment V2 preview did not include the standard duty allowance."
    Assert-True (([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 18)) -like "*$previewStdSpecial*") "Enrollment V2 preview did not include the special-conditions marker."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 83)).Trim() "20" "Enrollment V2 did not derive class qualification percent."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 86)).Trim() "90" "Enrollment V2 did not derive FIZO percent."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 89)).Trim() "20" "Enrollment V2 did not derive secrecy percent."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 92)).Trim() "60" "Enrollment V2 did not derive achievement amount."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 115)).Trim() "YES" "Enrollment V2 did not keep extra monthly payment enabled."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 138)).Trim() "YES" "Enrollment V2 did not keep extra one-time payment enabled."

    Write-Output "10aa. Enrollment wizard full-card round trip"
    $wizardRoundTrip = [string]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeEnrollmentWizardFullRoundTrip", 2)
    Assert-True (-not $wizardRoundTrip.StartsWith("ERROR:")) "Enrollment wizard full-card round-trip failed: $wizardRoundTrip"
    Assert-True ($wizardRoundTrip -like "*Надбавка за ученую степень*") "Enrollment wizard round-trip did not load the extra monthly payment name."
    Assert-True ($wizardRoundTrip -like "*Компенсация найма жилья*") "Enrollment wizard round-trip did not load the extra one-time payment name."
    Assert-True ($wizardRoundTrip -like "*40817810000000000001*") "Enrollment wizard round-trip did not preserve the bank account as text."
    Assert-True ($wizardRoundTrip -like "*31.12.2026*") "Enrollment wizard round-trip did not preserve the premium end date."
    Assert-True ($wizardRoundTrip -like "*диплом о присуждении степени*") "Enrollment wizard round-trip did not preserve the extra monthly basis."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 110)).Trim() "Надбавка за ученую степень" "Enrollment wizard save round-trip lost extra monthly payment name on the journal row."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 134)).Trim() "Компенсация найма жилья" "Enrollment wizard save round-trip lost extra one-time payment name on the journal row."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 58)).Trim() "40817810000000000001" "Enrollment wizard save round-trip lost bank account text on the journal row."

    $premiumSourceRow = Get-SettingRow -Worksheet $wsSettings -SettingKey "enrollment.def.premium.start_date_source"
    Assert-True ($premiumSourceRow -gt 0) "Premium definition start_date_source setting was not created."
    $originalPremiumSource = [string]$wsSettings.Cells($premiumSourceRow, 2).Value2
    Set-EnrollmentSetting -Worksheet $wsSettings -SettingKey "enrollment.def.premium.start_date_source" -SettingValue "manual"
    try {
        $wsEnrollment.Rows(2).Copy($wsEnrollment.Rows(3))
        $wsEnrollment.Cells(3, 12).Value = "15.08.2026"
        $wsEnrollment.Cells(3, 66).ClearContents()
        $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentRowDirect", $sheetEnrollment, 3)
        Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 3 -ColumnNumber 66)).Trim() "15.08.2026" "Premium start date did not follow manual start_date_source from settings."

        Set-EnrollmentSetting -Worksheet $wsSettings -SettingKey "enrollment.def.premium.start_date_source" -SettingValue "accept_date"
        $wsEnrollment.Rows(2).Copy($wsEnrollment.Rows(3))
        $wsEnrollment.Cells(3, 10).Value = "16.08.2026"
        $wsEnrollment.Cells(3, 66).ClearContents()
        $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentRowDirect", $sheetEnrollment, 3)
        Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 3 -ColumnNumber 66)).Trim() "16.08.2026" "Premium start date did not follow accept_date start_date_source from settings."

        Set-EnrollmentSetting -Worksheet $wsSettings -SettingKey "enrollment.def.premium.start_date_source" -SettingValue "order_date"
        $wsEnrollment.Rows(2).Copy($wsEnrollment.Rows(3))
        $wsEnrollment.Cells(3, 7).Value = "17.08.2026"
        $wsEnrollment.Cells(3, 66).ClearContents()
        $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentRowDirect", $sheetEnrollment, 3)
        Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 3 -ColumnNumber 66)).Trim() "17.08.2026" "Premium start date did not follow order_date start_date_source from settings."

        Set-EnrollmentSetting -Worksheet $wsSettings -SettingKey "enrollment.def.premium.start_date_source" -SettingValue "unknown_start_source"
        $wsEnrollment.Rows(2).Copy($wsEnrollment.Rows(3))
        $wsEnrollment.Cells(3, 66).ClearContents()
        $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentRowDirect", $sheetEnrollment, 3)
        Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 3 -ColumnNumber 66)).Trim() "" "Unknown start_date_source should not crash or write a premium start date."
    }
    finally {
        Set-EnrollmentSetting -Worksheet $wsSettings -SettingKey "enrollment.def.premium.start_date_source" -SettingValue $originalPremiumSource
        $wsEnrollment.Rows(3).ClearContents()
    }

    $createdRowsV2 = $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.GeneratePaymentsFromEnrollmentRowDirect", 2)
    Assert-True ($createdRowsV2 -ge 10) "Enrollment V2 row did not generate the expected built-in and reserve payment rows. Actual: $createdRowsV2; issues: $([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 25))"

    $paymentsLastRow = Get-LastUsedRow -Worksheet $wsPayments -ColumnNumber 4
    $generatedTypes = @()
    for ($row = 2; $row -le $paymentsLastRow; $row++) {
        $generatedTypes += ,([string](Get-CellValue -Worksheet $wsPayments -RowNumber $row -ColumnNumber 2)).Trim()
    }
    Assert-True ($generatedTypes -contains $paymentClassQualification) "Enrollment V2 transfer did not create class qualification row."
    Assert-True ($generatedTypes -contains $paymentFizo) "Enrollment V2 transfer did not create FIZO row."
    Assert-True ($generatedTypes -contains $paymentSecrecy) "Enrollment V2 transfer did not create secrecy row."
    Assert-True ($generatedTypes -contains $paymentAchievement) "Enrollment V2 transfer did not create achievement row."
    Assert-True ($generatedTypes -contains $paymentStdDuty) "Enrollment V2 transfer did not create standard duty row."
    Assert-True ($generatedTypes -contains $paymentStdSpecial) "Enrollment V2 transfer did not create standard special-conditions row."
    Assert-True ($generatedTypes -contains $paymentStdTariff) "Enrollment V2 transfer did not create tariff row."
    Assert-True ($generatedTypes -contains $paymentStdContract430) "Enrollment V2 transfer did not create 430-contract row."
    Assert-True ($generatedTypes -contains "Надбавка за ученую степень") "Enrollment V2 transfer did not create extra monthly payment row."
    Assert-True ($generatedTypes -contains "Компенсация найма жилья") "Enrollment V2 transfer did not create extra one-time payment row."

    Write-Output "10b. Enrollment backend save and generate"
    $wsPayments.Cells.Clear()
    $excel.Run("'$($workbook.Name)'!mdlDataValidation.HealWorkbookStructure", $true)
    $wsPayments = Get-WorksheetByName -Workbook $workbook -SheetName $sheetPayments

    $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.LoadEnrollmentRowToBackend", 2)
    $createdRowsBackend = $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.SaveEnrollmentFormAndGeneratePayments")
    Assert-True ($createdRowsBackend -ge 10) "Enrollment backend save+generate did not create the expected payment rows."
    Assert-Eq ([string]$excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.GetBackendValue", "current_row")).Trim() "2" "Enrollment backend save+generate did not preserve the current journal row."

    $paymentsLastRow = Get-LastUsedRow -Worksheet $wsPayments -ColumnNumber 4
    $generatedTypesBackend = @()
    for ($row = 2; $row -le $paymentsLastRow; $row++) {
        $generatedTypesBackend += ,([string](Get-CellValue -Worksheet $wsPayments -RowNumber $row -ColumnNumber 2)).Trim()
    }
    Assert-True ($generatedTypesBackend -contains $paymentClassQualification) "Enrollment backend save+generate did not create class qualification row."
    Assert-True ($generatedTypesBackend -contains $paymentStdDuty) "Enrollment backend save+generate did not create standard duty row."

    Write-Output "10c. Enrollment backend save and continue package"
    $continueBackendResult = [string]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeSaveAndContinuePackage")
    $continueDraftId = ($continueBackendResult -split "\|")[0]
    Assert-True ($continueBackendResult -match "^ORD-[^|]+\|\|500\|YES$") "Enrollment backend save+continue package did not preserve shared fields and clear personal fields."
    Assert-Eq ([string]$excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.GetBackendValue", "order_draft_id")).Trim() $continueDraftId "Enrollment backend save+continue package did not preserve OrderDraftId in backend."
    Assert-Eq ([string]$excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.GetBackendValue", "fio")).Trim() "" "Enrollment backend save+continue package did not clear backend FIO."

    Write-Output "10d. Enrollment wizard save and generate"
    $wsPayments.Cells.Clear()
    $excel.Run("'$($workbook.Name)'!mdlDataValidation.HealWorkbookStructure", $true)
    $wsPayments = Get-WorksheetByName -Workbook $workbook -SheetName $sheetPayments

    $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.LoadEnrollmentRowToBackend", 2)
    $wizardSaveGenerateResult = [string]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeEnrollmentWizardSaveGenerate")
    Assert-True ($wizardSaveGenerateResult -match "^2\|ORD-[^|]+\|[0-9]+$") "Enrollment wizard save+generate did not return the expected row/draft/count payload."
    $wizardSaveGenerateParts = $wizardSaveGenerateResult -split "\|"
    Assert-True ([int]$wizardSaveGenerateParts[2] -ge 10) "Enrollment wizard save+generate did not create the expected payment rows."
    Assert-Eq ([string]$excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.GetBackendValue", "current_row")).Trim() "2" "Enrollment wizard save+generate did not preserve the current journal row."

    Write-Output "10e. Enrollment wizard save and continue package"
    $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.LoadEnrollmentRowToBackend", 2)
    $wizardContinueResult = [string]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeEnrollmentWizardSaveContinue")
    $wizardContinueDraftId = ($wizardContinueResult -split "\|")[0]
    Assert-True ($wizardContinueResult -match "^ORD-[^|]+\|\|500\|YES$") "Enrollment wizard save+continue package did not preserve shared fields and clear personal fields."
    Assert-Eq ([string]$excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.GetBackendValue", "order_draft_id")).Trim() $wizardContinueDraftId "Enrollment wizard save+continue package did not preserve OrderDraftId in backend."
    Assert-Eq ([string]$excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.GetBackendValue", "fio")).Trim() "" "Enrollment wizard save+continue package did not clear backend FIO."

    Write-Output "10ea. Enrollment wizard check and preview"
    $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.LoadEnrollmentRowToBackend", 2)
    $wizardCheckResult = [string]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeEnrollmentWizardCheck")
    Assert-True ($wizardCheckResult -match "^(YES|NO)\|[1-9][0-9]*\|[0-9]+\|[0-9]+$") "Enrollment wizard check did not return the expected preview payload."
    $wizardCheckParts = $wizardCheckResult -split "\|"
    Assert-Eq $wizardCheckParts[0] "YES" "Enrollment wizard check should mark the complete draft as Word-ready."
    Assert-True ([int]$wizardCheckParts[1] -gt 100) "Enrollment wizard check did not build a substantial Section 1 preview."
    Assert-True ([string]$excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.GetBackendValue", "preview_section1") -like "*$fio1*") "Enrollment wizard check did not refresh preview_section1 in backend."

    Write-Output "10eb. Enrollment wizard export package"
    Get-ChildItem -LiteralPath $testDir -Filter "Enrollment_Order_*.docx" -ErrorAction SilentlyContinue | Remove-Item -Force
    $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.LoadEnrollmentRowToBackend", 2)
    $wizardExportPath = [string]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeEnrollmentWizardExport")
    Assert-True (Test-Path -LiteralPath $wizardExportPath) "Enrollment wizard export action did not create a .docx file."
    $wizardExportText = Get-DocxText -Path $wizardExportPath
    Assert-True ($wizardExportText -like "*$fio1*") "Enrollment wizard export action did not include the employee FIO."

    Write-Output "10f. Enrollment definition text templates from settings"
    $stdDutyTemplateRow = Get-SettingRow -Worksheet $wsSettings -SettingKey "enrollment.def.std_duty.text_template"
    Assert-True ($stdDutyTemplateRow -gt 0) "Enrollment std_duty text template setting was not created."
    $originalStdDutyTemplate = [string]$wsSettings.Cells($stdDutyTemplateRow, 2).Value2
    Set-EnrollmentSetting -Worksheet $wsSettings -SettingKey "enrollment.def.std_duty.text_template" -SettingValue "CUSTOM_STD_DUTY {percent}% for {fio}."
    try {
        $expectedStdDutyTemplateText = "CUSTOM_STD_DUTY 100% for $($fio1.Trim())."
        $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentRowDirect", $sheetEnrollment, 2)
        Assert-True ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 21) -like ("*" + $expectedStdDutyTemplateText + "*")) "Enrollment preview section did not use the configured std_duty text template."
        Get-ChildItem -LiteralPath $testDir -Filter "Enrollment_Order_*.docx" -ErrorAction SilentlyContinue | Remove-Item -Force
        $templateDrivenExportPath = [string]$excel.Run("'$($workbook.Name)'!mdlEnrollmentOrderExport.ExportEnrollmentOrderByRow", 2)
        $templateDrivenDocText = Get-DocxText -Path $templateDrivenExportPath
        Assert-True ($templateDrivenDocText -like ("*" + $expectedStdDutyTemplateText + "*")) "Enrollment Word export did not use the configured std_duty text template."
    }
    finally {
        Set-EnrollmentSetting -Worksheet $wsSettings -SettingKey "enrollment.def.std_duty.text_template" -SettingValue $originalStdDutyTemplate
        $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentRowDirect", $sheetEnrollment, 2)
    }

    $coreTemplateRow = Get-SettingRow -Worksheet $wsSettings -SettingKey "enrollment.def.core.text_template"
    $requisitesTemplateRow = Get-SettingRow -Worksheet $wsSettings -SettingKey "enrollment.def.requisites.text_template"
    Assert-True ($coreTemplateRow -gt 0) "Enrollment core text template setting was not created."
    Assert-True ($requisitesTemplateRow -gt 0) "Enrollment requisites text template setting was not created."
    $originalCoreTemplate = [string]$wsSettings.Cells($coreTemplateRow, 2).Value2
    $originalRequisitesTemplate = [string]$wsSettings.Cells($requisitesTemplateRow, 2).Value2
    Set-EnrollmentSetting -Worksheet $wsSettings -SettingKey "enrollment.def.core.text_template" -SettingValue "CUSTOM_CORE {fio}: {core_text}"
    Set-EnrollmentSetting -Worksheet $wsSettings -SettingKey "enrollment.def.requisites.text_template" -SettingValue "CUSTOM_REQ {passport_number}: {requisites_text}"
    try {
        $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentRowDirect", $sheetEnrollment, 2)
        $customBlockPreview = [string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 21)
        Assert-True ($customBlockPreview -like ("*CUSTOM_CORE " + $fio1.Trim() + ":*")) "Enrollment preview section did not use the configured core text template."
        Assert-True ($customBlockPreview -like "*CUSTOM_REQ 567890:*") "Enrollment preview section did not use the configured requisites text template."

        Get-ChildItem -LiteralPath $testDir -Filter "Enrollment_Order_*.docx" -ErrorAction SilentlyContinue | Remove-Item -Force
        $customBlockExportPath = [string]$excel.Run("'$($workbook.Name)'!mdlEnrollmentOrderExport.ExportEnrollmentOrderByRow", 2)
        $customBlockDocText = Get-DocxText -Path $customBlockExportPath
        Assert-True ($customBlockDocText -like ("*CUSTOM_CORE " + $fio1.Trim() + ":*")) "Enrollment Word export did not use the configured core text template."
        Assert-True ($customBlockDocText -like "*CUSTOM_REQ 567890:*") "Enrollment Word export did not use the configured requisites text template."
    }
    finally {
        Set-EnrollmentSetting -Worksheet $wsSettings -SettingKey "enrollment.def.core.text_template" -SettingValue $originalCoreTemplate
        Set-EnrollmentSetting -Worksheet $wsSettings -SettingKey "enrollment.def.requisites.text_template" -SettingValue $originalRequisitesTemplate
        $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentRowDirect", $sheetEnrollment, 2)
    }

    Write-Output "10fa. Enrollment payment definition contract"
    $definitionExpectations = @(
        @{ Code = "core"; Kind = "core"; Block = "Section1Core"; Binding = "core"; Required = ""; Start = "manual"; Severity = "blocked" },
        @{ Code = "std_duty"; Kind = "standard"; Block = "Section1MonthlyStandard"; Binding = "std_duty"; Required = "basis"; Start = "standard_start_date"; Severity = "warning" },
        @{ Code = "std_special"; Kind = "standard"; Block = "Section1MonthlyStandard"; Binding = "std_special"; Required = "basis"; Start = "standard_start_date"; Severity = "warning" },
        @{ Code = "std_tariff"; Kind = "standard"; Block = "Section1MonthlyStandard"; Binding = "std_tariff"; Required = "basis"; Start = "standard_start_date"; Severity = "warning" },
        @{ Code = "std_contract430"; Kind = "standard"; Block = "Section1MonthlyStandard"; Binding = "std_contract430"; Required = "basis"; Start = "standard_start_date"; Severity = "warning" },
        @{ Code = "class"; Kind = "personal"; Block = "Section1MonthlyPersonal"; Binding = "class"; Required = "param,basis"; Start = "standard_start_date"; Severity = "blocked" },
        @{ Code = "fizo"; Kind = "personal"; Block = "Section1MonthlyPersonal"; Binding = "fizo"; Required = "param,basis"; Start = "standard_start_date"; Severity = "blocked" },
        @{ Code = "secrecy"; Kind = "personal"; Block = "Section1MonthlyPersonal"; Binding = "secrecy"; Required = "param,basis"; Start = "standard_start_date"; Severity = "blocked" },
        @{ Code = "achievement"; Kind = "personal"; Block = "Section1MonthlyPersonal"; Binding = "achievement"; Required = "param,basis"; Start = "standard_start_date"; Severity = "blocked" },
        @{ Code = "lift"; Kind = "onetime"; Block = "Section1OneTime"; Binding = "lift"; Required = "amount,date,basis"; Start = "enroll_date"; Severity = "warning" },
        @{ Code = "per_diem"; Kind = "onetime"; Block = "Section1OneTime"; Binding = "per_diem"; Required = "amount,date,basis"; Start = "enroll_date"; Severity = "warning" },
        @{ Code = "edv"; Kind = "onetime"; Block = "Section2Edv400k"; Binding = "edv"; Required = "amount,date,basis_section2,contract_basis"; Start = "enroll_date"; Severity = "blocked" },
        @{ Code = "premium"; Kind = "premium"; Block = "Section1Premium"; Binding = "premium"; Required = "premium_end,basis"; Start = "enroll_date"; Severity = "blocked" },
        @{ Code = "requisites"; Kind = "requisites"; Block = "Section1Requisites"; Binding = "requisites"; Required = ""; Start = "manual"; Severity = "blocked" },
        @{ Code = "extra_monthly"; Kind = "personal"; Block = "Section1MonthlyPersonal"; Binding = "extra_monthly"; Required = "name,amount,start,basis"; Start = "manual"; Severity = "blocked" },
        @{ Code = "extra_onetime"; Kind = "onetime"; Block = "Section1OneTime"; Binding = "extra_one_time"; Required = "name,amount,date,basis"; Start = "manual"; Severity = "blocked" }
    )

    foreach ($definition in $definitionExpectations) {
        foreach ($fieldName in @("payment_kind", "word_block_target", "journal_binding", "required_docs", "start_date_source", "validation_severity", "label", "text_template")) {
            $settingRow = Get-SettingRow -Worksheet $wsSettings -SettingKey ("enrollment.def." + $definition["Code"] + "." + $fieldName)
            Assert-True ($settingRow -gt 0) ("Enrollment payment definition setting is missing: enrollment.def." + $definition["Code"] + "." + $fieldName)
        }

        Assert-Eq ([string]$excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.GetEnrollmentSetting", ("enrollment.def." + $definition["Code"] + ".payment_kind"), "")).Trim() $definition["Kind"] ("Wrong payment_kind for enrollment definition " + $definition["Code"])
        Assert-Eq ([string]$excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.GetEnrollmentSetting", ("enrollment.def." + $definition["Code"] + ".word_block_target"), "")).Trim() $definition["Block"] ("Wrong word_block_target for enrollment definition " + $definition["Code"])
        Assert-Eq ([string]$excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.GetEnrollmentSetting", ("enrollment.def." + $definition["Code"] + ".journal_binding"), "")).Trim() $definition["Binding"] ("Wrong journal_binding for enrollment definition " + $definition["Code"])
        Assert-Eq ([string]$excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.GetEnrollmentSetting", ("enrollment.def." + $definition["Code"] + ".required_docs"), "")).Trim() $definition["Required"] ("Wrong required_docs for enrollment definition " + $definition["Code"])
        Assert-Eq ([string]$excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.GetEnrollmentSetting", ("enrollment.def." + $definition["Code"] + ".start_date_source"), "")).Trim() $definition["Start"] ("Wrong start_date_source for enrollment definition " + $definition["Code"])
        Assert-Eq ([string]$excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.GetEnrollmentSetting", ("enrollment.def." + $definition["Code"] + ".validation_severity"), "")).Trim() $definition["Severity"] ("Wrong validation_severity for enrollment definition " + $definition["Code"])
    }
    Assert-Eq ([int]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeEnrollmentDefinitionBlockCount", "Section1Core")) 1 "Enrollment core block is not registered in the payment definition meta-model."
    Assert-Eq ([int]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeEnrollmentDefinitionBlockCount", "Section1Requisites")) 1 "Enrollment requisites block is not registered in the payment definition meta-model."

    Write-Output "10g. Enrollment order header settings and filename template"
    $headerUnitRow = Get-SettingRow -Worksheet $wsSettings -SettingKey "enrollment.unit_number"
    $headerCityRow = Get-SettingRow -Worksheet $wsSettings -SettingKey "enrollment.city"
    $headerSignatoryNameRow = Get-SettingRow -Worksheet $wsSettings -SettingKey "enrollment.signatory_name"
    $headerSignatoryRankRow = Get-SettingRow -Worksheet $wsSettings -SettingKey "enrollment.signatory_rank"
    $headerSignatoryPositionRow = Get-SettingRow -Worksheet $wsSettings -SettingKey "enrollment.signatory_position"
    $headerTextRow = Get-SettingRow -Worksheet $wsSettings -SettingKey "enrollment.header_text"
    $fileTemplateRow = Get-SettingRow -Worksheet $wsSettings -SettingKey "enrollment.filename_template"
    $wordTemplateRow = Get-SettingRow -Worksheet $wsSettings -SettingKey "enrollment.template_file"
    Assert-True ($headerUnitRow -gt 0) "Enrollment unit_number setting was not created."
    Assert-True ($headerCityRow -gt 0) "Enrollment city setting was not created."
    Assert-True ($headerSignatoryNameRow -gt 0) "Enrollment signatory_name setting was not created."
    Assert-True ($headerSignatoryRankRow -gt 0) "Enrollment signatory_rank setting was not created."
    Assert-True ($headerSignatoryPositionRow -gt 0) "Enrollment signatory_position setting was not created."
    Assert-True ($headerTextRow -gt 0) "Enrollment header_text setting was not created."
    Assert-True ($fileTemplateRow -gt 0) "Enrollment filename_template setting was not created."
    Assert-True ($wordTemplateRow -gt 0) "Enrollment template_file setting was not created."
    Assert-Eq ([string]$wsSettings.Cells($wordTemplateRow, 2).Value2) $enrollmentTemplateName "Enrollment template_file setting does not point to the dedicated enrollment template."

    $originalHeaderUnit = [string]$wsSettings.Cells($headerUnitRow, 2).Value2
    $originalHeaderCity = [string]$wsSettings.Cells($headerCityRow, 2).Value2
    $originalHeaderSignatoryName = [string]$wsSettings.Cells($headerSignatoryNameRow, 2).Value2
    $originalHeaderSignatoryRank = [string]$wsSettings.Cells($headerSignatoryRankRow, 2).Value2
    $originalHeaderSignatoryPosition = [string]$wsSettings.Cells($headerSignatoryPositionRow, 2).Value2
    $originalHeaderText = [string]$wsSettings.Cells($headerTextRow, 2).Value2
    $originalFileTemplate = [string]$wsSettings.Cells($fileTemplateRow, 2).Value2

    Set-EnrollmentSetting -Worksheet $wsSettings -SettingKey "enrollment.unit_number" -SettingValue "99999"
    Set-EnrollmentSetting -Worksheet $wsSettings -SettingKey "enrollment.city" -SettingValue "Казань"
    Set-EnrollmentSetting -Worksheet $wsSettings -SettingKey "enrollment.signatory_name" -SettingValue "И.И. Тестов"
    Set-EnrollmentSetting -Worksheet $wsSettings -SettingKey "enrollment.signatory_rank" -SettingValue "подполковник"
    Set-EnrollmentSetting -Worksheet $wsSettings -SettingKey "enrollment.signatory_position" -SettingValue "ИСПОЛНЯЮЩИЙ ОБЯЗАННОСТИ КОМАНДИРА"
    Set-EnrollmentSetting -Worksheet $wsSettings -SettingKey "enrollment.header_text" -SettingValue "ТЕСТОВЫЙ ЗАГОЛОВОК {unit}|ПО СТРОЕВОЙ ЧАСТИ"
    Set-EnrollmentSetting -Worksheet $wsSettings -SettingKey "enrollment.filename_template" -SettingValue "ENROLL_TEST_{orderDraftId}_{date}"
    try {
        $currentHeaderDraftId = ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 22)).Trim()
        Assert-True ($currentHeaderDraftId -ne "") "Enrollment header settings test requires a non-empty OrderDraftId in row 2."
        Get-ChildItem -LiteralPath $testDir -Filter "ENROLL_TEST_*.docx" -ErrorAction SilentlyContinue | Remove-Item -Force
        $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentRowDirect", $sheetEnrollment, 2)
        $settingsExportPath = [string]$excel.Run("'$($workbook.Name)'!mdlEnrollmentOrderExport.ExportEnrollmentOrderByRow", 2)
        Assert-True (Test-Path -LiteralPath $settingsExportPath) "Enrollment export with custom header settings did not create a .docx file."
        Assert-True ((Split-Path -Leaf $settingsExportPath) -like ("ENROLL_TEST_" + $currentHeaderDraftId + "_*.docx")) "Enrollment export did not use the configured filename template."

        $settingsDocText = Get-DocxText -Path $settingsExportPath
        Assert-True ($settingsDocText -like "*ТЕСТОВЫЙ ЗАГОЛОВОК 99999*") "Enrollment export did not use the configured header_text with unit placeholder."
        Assert-True ($settingsDocText -like "*ПО СТРОЕВОЙ ЧАСТИ*") "Enrollment export did not include the configured second header line."
        Assert-True ($settingsDocText -like "*г. Казань*") "Enrollment export did not use the configured city."
        Assert-True ($settingsDocText -like "*ИСПОЛНЯЮЩИЙ ОБЯЗАННОСТИ КОМАНДИРА*") "Enrollment export did not use the configured signatory position."
        Assert-True ($settingsDocText -like "*подполковник И.И. Тестов*") "Enrollment export did not use the configured signatory rank/name."
    }
    finally {
        Set-EnrollmentSetting -Worksheet $wsSettings -SettingKey "enrollment.unit_number" -SettingValue $originalHeaderUnit
        Set-EnrollmentSetting -Worksheet $wsSettings -SettingKey "enrollment.city" -SettingValue $originalHeaderCity
        Set-EnrollmentSetting -Worksheet $wsSettings -SettingKey "enrollment.signatory_name" -SettingValue $originalHeaderSignatoryName
        Set-EnrollmentSetting -Worksheet $wsSettings -SettingKey "enrollment.signatory_rank" -SettingValue $originalHeaderSignatoryRank
        Set-EnrollmentSetting -Worksheet $wsSettings -SettingKey "enrollment.signatory_position" -SettingValue $originalHeaderSignatoryPosition
        Set-EnrollmentSetting -Worksheet $wsSettings -SettingKey "enrollment.header_text" -SettingValue $originalHeaderText
        Set-EnrollmentSetting -Worksheet $wsSettings -SettingKey "enrollment.filename_template" -SettingValue $originalFileTemplate
        $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentRowDirect", $sheetEnrollment, 2)
    }

    Write-Output "11. Enrollment Word export with localization"
    Get-ChildItem -LiteralPath $testDir -Filter "Enrollment_Order_*.docx" -ErrorAction SilentlyContinue | Remove-Item -Force

    $wsEnrollment.Cells(2, 22).Value = "ORD-EXPORT-1"
    $wsEnrollment.Cells(2, 46).Value = "выписка из приказа, предписание, рапорт"
    $wsEnrollment.Cells(2, 47).Value = "основание для ЕДВ"
    $wsEnrollment.Cells(2, 36).Value = "5000"
    $wsEnrollment.Cells(2, 48).Value = "01.01.2000"
    $wsEnrollment.Cells(2, 49).Value = "г. Тестоград"
    $wsEnrollment.Cells(2, 50).Value = "Российская Федерация"
    $wsEnrollment.Cells(2, 73).Value = "YES"
    $wsEnrollment.Cells(2, 74).Value = "2"
    $wsEnrollment.Cells(2, 75).Value = "700 руб."
    $wsEnrollment.Cells(2, 76).Value = "02.07.2026"
    $wsEnrollment.Cells(2, 77).Value = "командировочное предписание"
    $wsEnrollment.Cells(2, 78).Value = "YES"
    $wsEnrollment.Cells(2, 79).Value = "400000"
    $wsEnrollment.Cells(2, 80).Value = "02.07.2026"
    $wsEnrollment.Cells(2, 81).Value = "Указ № 644"

    $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentRowDirect", $sheetEnrollment, 2)
    $exportPath = [string]$excel.Run("'$($workbook.Name)'!mdlEnrollmentOrderExport.ExportEnrollmentOrderByRow", 2)
    Assert-True (Test-Path -LiteralPath $exportPath) "Enrollment Word export did not create a .docx file."

    $docText = Get-DocxText -Path $exportPath
    $docXml = Get-DocxXml -Path $exportPath
    Assert-True ($docText -like "*$fio1*") "Enrollment Word export did not include the employee FIO."
    Assert-True ($docText -like "*400000*") "Enrollment Word export did not include the EDV block."
    Assert-True ($docText -like "*§ 2*") "Enrollment Word export did not include section 2 for EDV."
    Assert-True ($docXml -like "*Times New Roman*") "Enrollment Word export did not apply Times New Roman formatting."
    Assert-True ($docXml -like "*w:val=`"28`"*") "Enrollment Word export did not apply 14 pt font size."
    Assert-True ($docXml -like "*w:jc w:val=`"center`"*") "Enrollment Word export did not center header or section captions."
    Assert-True ($docXml -like "*w:jc w:val=`"both`"*") "Enrollment Word export did not justify body paragraphs."
    Assert-True ($docXml -like "*w:pgMar*") "Enrollment Word export did not apply page margins."
    Assert-True ($docXml -like "*w:firstLine=*") "Enrollment Word export did not apply first-line paragraph indentation."
    Assert-True ($docXml -like "*w:keepNext*") "Enrollment Word export should keep section captions with the following paragraph."
    Assert-True ($docXml -like "*w:keepLines*") "Enrollment Word export should keep section 2/signature paragraphs from splitting into orphan lines."
    Assert-True ($docText -like "*Надбавка за ученую степень*") "Enrollment Word export did not include the extra monthly payment."
    Assert-True ($docText -like "*Компенсация найма жилья*") "Enrollment Word export did not include the extra one-time payment."
    Assert-True ($docText -like "*оклад по воинскому званию*5000*") "Enrollment Word export did not include the rank salary paragraph."
    Assert-True ($docText -like "*Дата рождения:*01.01.2000*") "Enrollment Word export did not include birth date requisites."
    Assert-True ($docText -like "*место рождения:*г. Тестоград*") "Enrollment Word export did not include birth place requisites."
    Assert-True ($docText -like "*гражданство:*Российская Федерация*") "Enrollment Word export did not include citizenship requisites."
    Assert-True ($docText -like "*суточные*700 руб.*за 2 сут.*") "Enrollment Word export did not include per diem days and amount."
    Assert-True ($docText -like "*40817810000000000001*") "Enrollment Word export converted the bank account instead of preserving it as text."
    Assert-True ($docText -notlike "*E+*") "Enrollment Word export leaked scientific notation into requisites."
    Assert-True ($docText -notlike "*enrollment.word.*") "Enrollment Word export leaked localization keys into the document."
    foreach ($legacyToken in $legacyEnrollmentTemplateTokens) {
        Assert-True (-not $docText.Contains($legacyToken)) "Enrollment Word export leaked legacy template placeholder $legacyToken."
    }

    Write-Output "11c. Enrollment Word export inserts body into template marker"
    $markerTemplateName = "EnrollmentMarkerTemplate.docx"
    $markerTemplatePath = Join-Path $testDir $markerTemplateName
    $wordTemplateRowForMarker = Get-SettingRow -Worksheet $wsSettings -SettingKey "enrollment.template_file"
    $bodyMarkerRow = Get-SettingRow -Worksheet $wsSettings -SettingKey "enrollment.template_body_marker"
    Assert-True ($wordTemplateRowForMarker -gt 0) "Enrollment template_file setting was not created."
    Assert-True ($bodyMarkerRow -gt 0) "Enrollment template_body_marker setting was not created."
    $originalTemplateFileForMarker = [string]$wsSettings.Cells($wordTemplateRowForMarker, 2).Value2
    $originalBodyMarker = [string]$wsSettings.Cells($bodyMarkerRow, 2).Value2
    Remove-IfExists -Path $markerTemplatePath
    New-ListTemplateDocx -Path $markerTemplatePath -BodyText "TEMPLATE_PREFIX`r`n[ENROLLMENT_ORDER_BODY]`r`nTEMPLATE_SUFFIX"
    Set-EnrollmentSetting -Worksheet $wsSettings -SettingKey "enrollment.template_file" -SettingValue $markerTemplateName
    Set-EnrollmentSetting -Worksheet $wsSettings -SettingKey "enrollment.template_body_marker" -SettingValue "[ENROLLMENT_ORDER_BODY]"
    try {
        $markerExportPath = [string]$excel.Run("'$($workbook.Name)'!mdlEnrollmentOrderExport.ExportEnrollmentOrderByRow", 2)
        Assert-True (Test-Path -LiteralPath $markerExportPath) "Enrollment marker-template export did not create a .docx file."
        $markerDocText = Get-DocxText -Path $markerExportPath
        Assert-True ($markerDocText -like "*TEMPLATE_PREFIX*") "Enrollment export did not preserve text before the template body marker."
        Assert-True ($markerDocText -like "*TEMPLATE_SUFFIX*") "Enrollment export did not preserve text after the template body marker."
        Assert-True ($markerDocText -like "*$fio1*") "Enrollment export did not insert the generated body at the template marker."
        Assert-True (-not $markerDocText.Contains("[ENROLLMENT_ORDER_BODY]")) "Enrollment export left the template body marker in the output document."
    }
    finally {
        Set-EnrollmentSetting -Worksheet $wsSettings -SettingKey "enrollment.template_file" -SettingValue $originalTemplateFileForMarker
        Set-EnrollmentSetting -Worksheet $wsSettings -SettingKey "enrollment.template_body_marker" -SettingValue $originalBodyMarker
    }

    Write-Output "11a. Enrollment manual draft save without staff"
    $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.ClearEnrollmentForm")
    Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "source_mode" -FieldValue "manual"
    Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "fio" -FieldValue "Тестов Тест Тестович"
    Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "personal_number" -FieldValue "М-000001"
    Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "rank" -FieldValue "Рядовой"
    Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "position" -FieldValue "Тестовая должность"
    Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "section" -FieldValue "Тестовая часть"
    Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "military_unit" -FieldValue "Тестовая часть"
    Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "order_date" -FieldValue "07.07.2026"
    Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "order_number" -FieldValue "MD-1"
    Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "enroll_date" -FieldValue "07.07.2026"
    Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "arrival_source" -FieldValue "ручной ввод"
    $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentForm")
    $manualDraftRow = [int]$excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.SaveEnrollmentFormToSheet", $false)
    Assert-True ($manualDraftRow -ge 2) "Manual enrollment draft did not save into the journal."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber $manualDraftRow -ColumnNumber 2)).Trim() "Тестов Тест Тестович" "Manual enrollment draft did not keep manual FIO."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber $manualDraftRow -ColumnNumber 3)).Trim() "М-000001" "Manual enrollment draft did not keep manual personal number."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber $manualDraftRow -ColumnNumber 26)).Trim() "manual" "Manual enrollment draft did not keep SourceMode=manual."
    Assert-True (([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber $manualDraftRow -ColumnNumber 22)).Trim() -ne "") "Manual enrollment draft did not receive OrderDraftId."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber $manualDraftRow -ColumnNumber 23)).Trim() "NO" "Incomplete manual enrollment draft should not be Word-ready."

    Write-Output "11aa. Enrollment manual mode does not auto-hydrate staff"
    $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.ClearEnrollmentForm")
    Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "source_mode" -FieldValue "manual"
    Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "fio" -FieldValue "Ручной Сотрудник Без Штата"
    Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "personal_number" -FieldValue $ln1.Trim()
    Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "position" -FieldValue "Ручная должность"
    Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "order_date" -FieldValue "07.07.2026"
    Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "order_number" -FieldValue "MD-2"
    Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "enroll_date" -FieldValue "07.07.2026"
    Set-BackendValue -Worksheet $wsEnrollmentForm -FieldKey "arrival_source" -FieldValue "ручной ввод существующего номера"
    $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentForm")
    $manualExistingNumberRow = [int]$excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.SaveEnrollmentFormToSheet", $false)
    Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber $manualExistingNumberRow -ColumnNumber 2)).Trim() "Ручной Сотрудник Без Штата" "Manual enrollment with an existing staff number should keep manual FIO."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber $manualExistingNumberRow -ColumnNumber 4)).Trim() "" "Manual enrollment with an existing staff number should not auto-fill rank from Staff."
    Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber $manualExistingNumberRow -ColumnNumber 26)).Trim() "manual" "Manual enrollment with an existing staff number should keep SourceMode=manual."

    Write-Output "11b. Enrollment grouped export by shared OrderDraftId"
    $wsEnrollment.Cells.Clear()
    $excel.Run("'$($workbook.Name)'!mdlDataValidation.HealWorkbookStructure", $true)
    $wsEnrollment = Get-WorksheetByName -Workbook $workbook -SheetName $sheetEnrollment

    $wsEnrollment.Cells(2, 3).Value = $ln1.Trim()
    $wsEnrollment.Cells(2, 7).Value = "08.07.2026"
    $wsEnrollment.Cells(2, 8).Value = "801"
    $wsEnrollment.Cells(2, 10).Value = "08.07.2026"
    $wsEnrollment.Cells(2, 11).Value = "08.07.2026"
    $wsEnrollment.Cells(2, 13).Value = "report group 1"
    $wsEnrollment.Cells(2, 14).Value = "assignment group 1"
    $wsEnrollment.Cells(2, 22).Value = "ORD-GROUP-OK"
    $wsEnrollment.Cells(2, 38).Value = "пункт отбора"
    $wsEnrollment.Cells(2, 78).Value = "NO"
    Set-EnrollmentRequiredFields -Worksheet $wsEnrollment -RowNumber 2 -OrderIssuer "issuer shared"

    $wsEnrollment.Cells(3, 3).Value = $ln2.Trim()
    $wsEnrollment.Cells(3, 7).Value = "08.07.2026"
    $wsEnrollment.Cells(3, 8).Value = "801"
    $wsEnrollment.Cells(3, 10).Value = "08.07.2026"
    $wsEnrollment.Cells(3, 11).Value = "08.07.2026"
    $wsEnrollment.Cells(3, 13).Value = "report group 2"
    $wsEnrollment.Cells(3, 14).Value = "assignment group 2"
    $wsEnrollment.Cells(3, 22).Value = "ORD-GROUP-OK"
    $wsEnrollment.Cells(3, 38).Value = "пункт отбора"
    $wsEnrollment.Cells(3, 78).Value = "YES"
    $wsEnrollment.Cells(3, 79).Value = "400000"
    $wsEnrollment.Cells(3, 80).Value = "08.07.2026"
    $wsEnrollment.Cells(3, 81).Value = "Указ № 644"
    Set-EnrollmentRequiredFields -Worksheet $wsEnrollment -RowNumber 3 -OrderIssuer "issuer shared"
    $wsEnrollment.Cells(3, 47).Value = "основание для ЕДВ"

    $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentRowDirect", $sheetEnrollment, 2)
    $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentRowDirect", $sheetEnrollment, 3)
    Assert-Eq ([int]$excel.Run("'$($workbook.Name)'!mdlEnrollmentOrderExport.GetExportRowCount", "ORD-GROUP-OK", 0)) 2 "Shared OrderDraftId should collect exactly two enrollment rows."
    $groupExportPath = [string]$excel.Run("'$($workbook.Name)'!mdlEnrollmentOrderExport.ExportEnrollmentOrderByDraftId", "ORD-GROUP-OK", 0)
    Assert-True (Test-Path -LiteralPath $groupExportPath) "Grouped enrollment export did not create a .docx file."

    $groupDocText = Get-DocxText -Path $groupExportPath
    Assert-True ($groupDocText -like "*$fio1*") "Grouped enrollment export did not include the first employee."
    Assert-True ($groupDocText -like "*$fio2*") "Grouped enrollment export did not include the second employee."
    Assert-True ($groupDocText -like "*§ 2*") "Grouped enrollment export should include section 2 when one row has EDV."
    Assert-True ((Get-MatchCount -Text $groupDocText -Pattern $fio1.Trim()) -ge 1) "Grouped enrollment export did not retain the first employee text."
    Assert-True ((Get-MatchCount -Text $groupDocText -Pattern $fio2.Trim()) -ge 1) "Grouped enrollment export did not retain the second employee text."
    $section2Start = $groupDocText.IndexOf("§ 2")
    Assert-True ($section2Start -ge 0) "Grouped enrollment export did not expose a section 2 block for numbering verification."
    $section2OnlyText = $groupDocText.Substring($section2Start)
    Assert-True ($section2OnlyText -like ("*1. *" + $fio2.Trim() + "*")) "Section 2 numbering should start from 1 for the first EDV row, even when earlier package rows have no EDV."
    Assert-True ($section2OnlyText -notlike ("*2. *" + $fio2.Trim() + "*")) "Section 2 numbering skipped the first number when earlier package rows had no EDV."

    Write-Output "12. Enrollment grouped export conflict blocking"
    $wsEnrollment.Cells.Clear()
    $excel.Run("'$($workbook.Name)'!mdlDataValidation.HealWorkbookStructure", $true)
    $wsEnrollment = Get-WorksheetByName -Workbook $workbook -SheetName $sheetEnrollment

    $wsEnrollment.Cells(2, 3).Value = $ln1.Trim()
    $wsEnrollment.Cells(2, 7).Value = "01.07.2026"
    $wsEnrollment.Cells(2, 8).Value = "700"
    $wsEnrollment.Cells(2, 10).Value = "04.07.2026"
    $wsEnrollment.Cells(2, 11).Value = "02.07.2026"
    $wsEnrollment.Cells(2, 13).Value = "report conflict 1"
    $wsEnrollment.Cells(2, 14).Value = "assignment conflict 1"
    $wsEnrollment.Cells(2, 22).Value = "ORD-CONFLICT"
    $wsEnrollment.Cells(2, 38).Value = "пункт отбора"
    Set-EnrollmentRequiredFields -Worksheet $wsEnrollment -RowNumber 2 -OrderIssuer "issuer one"

    $wsEnrollment.Cells(3, 3).Value = $ln2.Trim()
    $wsEnrollment.Cells(3, 7).Value = "01.07.2026"
    $wsEnrollment.Cells(3, 8).Value = "701"
    $wsEnrollment.Cells(3, 10).Value = "04.07.2026"
    $wsEnrollment.Cells(3, 11).Value = "02.07.2026"
    $wsEnrollment.Cells(3, 13).Value = "report conflict 2"
    $wsEnrollment.Cells(3, 14).Value = "assignment conflict 2"
    $wsEnrollment.Cells(3, 22).Value = "ORD-CONFLICT"
    $wsEnrollment.Cells(3, 38).Value = "пункт отбора"
    Set-EnrollmentRequiredFields -Worksheet $wsEnrollment -RowNumber 3 -OrderIssuer "issuer two"

    $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentRowDirect", $sheetEnrollment, 2)
    $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentRowDirect", $sheetEnrollment, 3)

    $conflictResult = [string]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeEnrollmentConflict", "ORD-CONFLICT")
    Assert-True ($conflictResult -like "ERROR:*") "Enrollment grouped export should block conflicting rows with the same OrderDraftId."
    Assert-True ($conflictResult -like "*Номер приказа*700*701*") "Enrollment grouped export conflict did not list the conflicting order numbers."
    Assert-True ($conflictResult -like "*Кем издан приказ*issuer one*issuer two*") "Enrollment grouped export conflict did not list the conflicting order issuers."
    $conflictRawResult = [string]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeEnrollmentConflictRaw", "ORD-CONFLICT")
    Assert-True ($conflictRawResult -like "ERROR:*") "Enrollment grouped export conflict should return a controlled ERROR result without raising a VBA runtime error."
    $conflictPreflightIssues = [string]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeEnrollmentExportBlockingIssues", "ORD-CONFLICT", 0)
    Assert-True (-not $conflictPreflightIssues.StartsWith("ERROR:")) "Enrollment conflict preflight should return blocking issues without raising a VBA runtime error."
    Assert-True ($conflictPreflightIssues -like "*конфликт*") "Enrollment conflict preflight did not return the user-facing conflict message."
    Assert-True ($conflictPreflightIssues -like "*Номер приказа*700*701*") "Enrollment conflict preflight did not list the conflicting order numbers."
    Assert-True ($conflictPreflightIssues -like "*Кем издан приказ*issuer one*issuer two*") "Enrollment conflict preflight did not list the conflicting order issuers."

    Write-Output "13. Enrollment export blocks incomplete draft"
    $wsEnrollment.Cells.Clear()
    $excel.Run("'$($workbook.Name)'!mdlDataValidation.HealWorkbookStructure", $true)
    $wsEnrollment = Get-WorksheetByName -Workbook $workbook -SheetName $sheetEnrollment

    $wsEnrollment.Cells(2, 3).Value = $ln1.Trim()
    $wsEnrollment.Cells(2, 7).Value = "06.07.2026"
    $wsEnrollment.Cells(2, 8).Value = "991"
    $wsEnrollment.Cells(2, 10).Value = "06.07.2026"
    $wsEnrollment.Cells(2, 11).Value = "06.07.2026"
    $wsEnrollment.Cells(2, 22).Value = "ORD-BLOCKED-EXPORT"
    $wsEnrollment.Cells(2, 38).Value = "пункт отбора"
    $wsEnrollment.Cells(2, 43).Value = "06.07.2026"
    $wsEnrollment.Cells(2, 44).Value = "06.07.2026"
    $wsEnrollment.Cells(2, 45).Value = "06.07.2026"
    $wsEnrollment.Cells(2, 46).Value = "черновик без реквизитов"
    $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentRowDirect", $sheetEnrollment, 2)

    $blockedExportResult = [string]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeEnrollmentExportRow", 2)
    Assert-True ($blockedExportResult -like "ERROR:*") "Enrollment export should block an incomplete draft row."
    Assert-True ($blockedExportResult -like "*Word*") "Blocked enrollment export did not return the readiness error."

    Write-Output "13aa. Enrollment export blocks missing position salary"
    $wsEnrollment.Cells.Clear()
    $excel.Run("'$($workbook.Name)'!mdlDataValidation.HealWorkbookStructure", $true)
    $wsEnrollment = Get-WorksheetByName -Workbook $workbook -SheetName $sheetEnrollment

    $wsEnrollment.Cells(2, 3).Value = $ln1.Trim()
    $wsEnrollment.Cells(2, 7).Value = "06.07.2026"
    $wsEnrollment.Cells(2, 8).Value = "991-OKLAD"
    $wsEnrollment.Cells(2, 10).Value = "06.07.2026"
    $wsEnrollment.Cells(2, 11).Value = "06.07.2026"
    $wsEnrollment.Cells(2, 13).Value = "report salary missing"
    $wsEnrollment.Cells(2, 14).Value = "assignment salary missing"
    $wsEnrollment.Cells(2, 22).Value = "ORD-MISSING-SALARY"
    $wsEnrollment.Cells(2, 34).Value = "4"
    $wsEnrollment.Cells(2, 38).Value = "пункт отбора"
    $wsEnrollment.Cells(2, 43).Value = "06.07.2026"
    $wsEnrollment.Cells(2, 44).Value = "06.07.2026"
    $wsEnrollment.Cells(2, 45).Value = "06.07.2026"
    Set-EnrollmentRequiredFields -Worksheet $wsEnrollment -RowNumber 2
    $wsEnrollment.Cells(2, 35).ClearContents()
    $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentRowDirect", $sheetEnrollment, 2)

    Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 23)).Trim() "NO" "Missing position salary should block Word readiness."
    Assert-Eq ([int](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 24)) 2 "Missing position salary should have blocked severity."
    $salaryIssues = [string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 25)
    Assert-True ($salaryIssues -like "*оклад*") "Missing position salary validation did not mention salary."

    $missingSalaryExportResult = [string]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeEnrollmentExportRow", 2)
    Assert-True ($missingSalaryExportResult -like "ERROR:*") "Enrollment export should block rows without position salary."
    Assert-True ($missingSalaryExportResult -like "*оклад*") "Blocked salary export did not include an actionable salary message."

    Write-Output "13a. Enrollment personal payment requires parameter"
    $wsEnrollment.Cells.Clear()
    $excel.Run("'$($workbook.Name)'!mdlDataValidation.HealWorkbookStructure", $true)
    $wsEnrollment = Get-WorksheetByName -Workbook $workbook -SheetName $sheetEnrollment

    $wsEnrollment.Cells(2, 3).Value = $ln1.Trim()
    $wsEnrollment.Cells(2, 7).Value = "06.07.2026"
    $wsEnrollment.Cells(2, 8).Value = "992"
    $wsEnrollment.Cells(2, 10).Value = "06.07.2026"
    $wsEnrollment.Cells(2, 11).Value = "06.07.2026"
    $wsEnrollment.Cells(2, 13).Value = "report personal missing param"
    $wsEnrollment.Cells(2, 14).Value = "assignment personal missing param"
    $wsEnrollment.Cells(2, 22).Value = "ORD-PERSONAL-PARAM-BLOCKED"
    $wsEnrollment.Cells(2, 34).Value = "3"
    $wsEnrollment.Cells(2, 35).Value = "20000"
    $wsEnrollment.Cells(2, 38).Value = "пункт отбора"
    $wsEnrollment.Cells(2, 82).Value = "YES"
    Set-EnrollmentRequiredFields -Worksheet $wsEnrollment -RowNumber 2
    $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentRowDirect", $sheetEnrollment, 2)

    Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 23)).Trim() "NO" "Enabled class qualification without parameter should block Word readiness."
    Assert-Eq ([int](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 24)) 2 "Enabled class qualification without parameter should have blocked severity."
    $personalParamIssues = [string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 25)
    Assert-True ($personalParamIssues -like "*классность*") "Personal payment parameter validation did not mention class qualification."
    Assert-True ($personalParamIssues -like "*параметр*") "Personal payment parameter validation did not mention the missing parameter."

    $personalParamExportResult = [string]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeEnrollmentExportRow", 2)
    Assert-True ($personalParamExportResult -like "ERROR:*") "Enrollment export should block enabled personal payment without a required parameter."

    Write-Output "13ab. Enrollment personal payment requires resolved amount"
    $wsEnrollment.Cells.Clear()
    $excel.Run("'$($workbook.Name)'!mdlDataValidation.HealWorkbookStructure", $true)
    $wsEnrollment = Get-WorksheetByName -Workbook $workbook -SheetName $sheetEnrollment

    $wsEnrollment.Cells(2, 3).Value = $ln1.Trim()
    $wsEnrollment.Cells(2, 7).Value = "06.07.2026"
    $wsEnrollment.Cells(2, 8).Value = "992-RATE"
    $wsEnrollment.Cells(2, 10).Value = "06.07.2026"
    $wsEnrollment.Cells(2, 11).Value = "06.07.2026"
    $wsEnrollment.Cells(2, 13).Value = "report unresolved rate"
    $wsEnrollment.Cells(2, 14).Value = "assignment unresolved rate"
    $wsEnrollment.Cells(2, 22).Value = "ORD-PERSONAL-RATE-BLOCKED"
    $wsEnrollment.Cells(2, 34).Value = "3"
    $wsEnrollment.Cells(2, 35).Value = "20000"
    $wsEnrollment.Cells(2, 38).Value = "пункт отбора"
    $wsEnrollment.Cells(2, 15).Value = "неизвестная классность"
    Set-EnrollmentRequiredFields -Worksheet $wsEnrollment -RowNumber 2
    $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentRowDirect", $sheetEnrollment, 2)

    Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 23)).Trim() "NO" "Enabled class qualification with unresolved amount should block Word readiness."
    Assert-Eq ([int](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 24)) 2 "Enabled class qualification with unresolved amount should have blocked severity."
    $personalRateIssues = [string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 25)
    Assert-True ($personalRateIssues -like "*классность*") "Personal payment amount validation did not mention class qualification."
    Assert-True ($personalRateIssues -like "*размер*") "Personal payment amount validation did not mention the unresolved amount."

    $personalRateExportResult = [string]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeEnrollmentExportRow", 2)
    Assert-True ($personalRateExportResult -like "ERROR:*") "Enrollment export should block enabled personal payment with unresolved amount."
    Assert-True ($personalRateExportResult -like "*размер*") "Blocked personal payment amount export did not include an actionable amount message."

    Write-Output "13b. Enrollment EDV requires contract basis"
    $wsEnrollment.Cells.Clear()
    $excel.Run("'$($workbook.Name)'!mdlDataValidation.HealWorkbookStructure", $true)
    $wsEnrollment = Get-WorksheetByName -Workbook $workbook -SheetName $sheetEnrollment

    $wsEnrollment.Cells(2, 3).Value = $ln1.Trim()
    $wsEnrollment.Cells(2, 7).Value = "06.07.2026"
    $wsEnrollment.Cells(2, 8).Value = "993"
    $wsEnrollment.Cells(2, 10).Value = "06.07.2026"
    $wsEnrollment.Cells(2, 11).Value = "06.07.2026"
    $wsEnrollment.Cells(2, 13).Value = "report edv missing contract basis"
    $wsEnrollment.Cells(2, 14).Value = "assignment edv missing contract basis"
    $wsEnrollment.Cells(2, 22).Value = "ORD-EDV-CONTRACT-BASIS-BLOCKED"
    $wsEnrollment.Cells(2, 34).Value = "3"
    $wsEnrollment.Cells(2, 35).Value = "20000"
    $wsEnrollment.Cells(2, 38).Value = "пункт отбора"
    $wsEnrollment.Cells(2, 47).Value = "основание для §2"
    $wsEnrollment.Cells(2, 78).Value = "YES"
    $wsEnrollment.Cells(2, 79).Value = "400000"
    $wsEnrollment.Cells(2, 80).Value = "06.07.2026"
    $wsEnrollment.Cells(2, 81).Value = "Указ № 644"
    Set-EnrollmentRequiredFields -Worksheet $wsEnrollment -RowNumber 2
    $wsEnrollment.Cells(2, 31).ClearContents()
    $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentRowDirect", $sheetEnrollment, 2)

    Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 23)).Trim() "NO" "Enabled EDV without contract basis should block Word readiness."
    Assert-Eq ([int](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 24)) 2 "Enabled EDV without contract basis should have blocked severity."
    $edvContractIssues = [string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 25)
    Assert-True ($edvContractIssues -like "*ЕДВ*") "EDV contract-basis validation did not mention EDV."
    Assert-True ($edvContractIssues -like "*контракт*") "EDV contract-basis validation did not mention contract basis."

    $edvContractExportResult = [string]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeEnrollmentExportRow", 2)
    Assert-True ($edvContractExportResult -like "ERROR:*") "Enrollment export should block enabled EDV without contract basis."
    $edvContractRawResult = [string]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeEnrollmentExportRowRaw", 2)
    Assert-True ($edvContractRawResult -like "ERROR:*") "Enrollment export should return a controlled ERROR result for enabled EDV without contract basis, not raise a VBA runtime error."
    $edvPreflightIssues = [string]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeEnrollmentExportBlockingIssues", "ORD-EDV-CONTRACT-BASIS-BLOCKED", 2)
    Assert-True (-not $edvPreflightIssues.StartsWith("ERROR:")) "Enrollment export preflight should return blocking issues without raising a VBA runtime error."
    Assert-True ($edvPreflightIssues -like "*Word-приказ*") "Enrollment export preflight did not return the user-facing Word readiness message."
    Assert-True ($edvPreflightIssues -like "*ЕДВ*") "Enrollment export preflight did not include the EDV issue."
    Assert-True ($edvPreflightIssues -like "*контракт*") "Enrollment export preflight did not include the contract-basis issue."

    Write-Output "13c. Enrollment wizard blocked export stays in preview"
    $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.LoadEnrollmentRowToBackend", 2)
    $wizardBlockedExportResult = [string]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeEnrollmentWizardExport")
    Assert-True ($wizardBlockedExportResult -like "ERROR:*") "Enrollment wizard export should return a controlled error for blocked EDV instead of raising a VBA runtime error."
    Assert-Eq ([string]$excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.GetBackendValue", "preview_word_ready")).Trim() "NO" "Blocked wizard export should keep backend preview marked as not Word-ready."
    $wizardBlockedIssues = [string]$excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.GetBackendValue", "preview_issues")
    Assert-True ($wizardBlockedIssues -like "*ЕДВ*") "Blocked wizard export preview did not preserve the EDV issue."
    Assert-True ($wizardBlockedIssues -like "*контракт*") "Blocked wizard export preview did not preserve the contract-basis issue."

    Write-Output "13d. Enrollment premium requires end date"
    $wsEnrollment.Cells.Clear()
    $excel.Run("'$($workbook.Name)'!mdlDataValidation.HealWorkbookStructure", $true)
    $wsEnrollment = Get-WorksheetByName -Workbook $workbook -SheetName $sheetEnrollment

    $wsEnrollment.Cells(2, 3).Value = $ln1.Trim()
    $wsEnrollment.Cells(2, 7).Value = "06.07.2026"
    $wsEnrollment.Cells(2, 8).Value = "994"
    $wsEnrollment.Cells(2, 10).Value = "06.07.2026"
    $wsEnrollment.Cells(2, 11).Value = "06.07.2026"
    $wsEnrollment.Cells(2, 13).Value = "report premium missing end"
    $wsEnrollment.Cells(2, 14).Value = "assignment premium missing end"
    $wsEnrollment.Cells(2, 22).Value = "ORD-PREMIUM-END-BLOCKED"
    $wsEnrollment.Cells(2, 38).Value = "пункт отбора"
    $wsEnrollment.Cells(2, 64).Value = "YES"
    $wsEnrollment.Cells(2, 65).Value = "25"
    $wsEnrollment.Cells(2, 66).Value = "06.07.2026"
    $wsEnrollment.Cells(2, 67).ClearContents()
    $wsEnrollment.Cells(2, 68).Value = "основание премии"
    Set-EnrollmentRequiredFields -Worksheet $wsEnrollment -RowNumber 2
    $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentRowDirect", $sheetEnrollment, 2)
    $wsEnrollment.Cells(2, 67).ClearContents()
    $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentRowDirect", $sheetEnrollment, 2)

    Assert-Eq ([string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 23)).Trim() "NO" "Enabled premium without end date should block Word readiness."
    Assert-Eq ([int](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 24)) 2 "Enabled premium without end date should have blocked severity."
    $premiumEndIssues = [string](Get-CellValue -Worksheet $wsEnrollment -RowNumber 2 -ColumnNumber 25)
    Assert-True ($premiumEndIssues -like "*преми*") "Premium end-date validation did not mention premium."
    Assert-True ($premiumEndIssues -like "*окончан*") "Premium end-date validation did not mention the missing end date."
    $premiumEndExportResult = [string]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeEnrollmentExportRow", 2)
    Assert-True ($premiumEndExportResult -like "ERROR:*") "Enrollment export should block enabled premium without end date."

    Write-Output "14. Selected enrollment row core workflow"
    $wsEnrollment.Cells.Clear()
    $excel.Run("'$($workbook.Name)'!mdlDataValidation.HealWorkbookStructure", $true)
    $wsEnrollment = Get-WorksheetByName -Workbook $workbook -SheetName $sheetEnrollment

    $wsEnrollment.Cells(2, 3).Value = $ln1.Trim()
    $wsEnrollment.Cells(2, 7).Value = "06.07.2026"
    $wsEnrollment.Cells(2, 8).Value = "990"
    $wsEnrollment.Cells(2, 10).Value = "06.07.2026"
    $wsEnrollment.Cells(2, 11).Value = "06.07.2026"
    $wsEnrollment.Cells(2, 13).Value = "report selected core"
    $wsEnrollment.Cells(2, 14).Value = "assignment selected core"
    $wsEnrollment.Cells(2, 22).Value = "ORD-SELECTED-CORE"
    $wsEnrollment.Cells(2, 34).Value = "4"
    $wsEnrollment.Cells(2, 35).Value = "18000"
    $wsEnrollment.Cells(2, 38).Value = "пункт отбора"
    $wsEnrollment.Cells(2, 46).Value = "выписка из приказа, предписание, рапорт"
    $wsEnrollment.Cells(2, 47).Value = "основание для ЕДВ"
    $wsEnrollment.Cells(2, 78).Value = "YES"
    $wsEnrollment.Cells(2, 79).Value = "400000"
    $wsEnrollment.Cells(2, 80).Value = "06.07.2026"
    $wsEnrollment.Cells(2, 81).Value = "Указ № 644"
    Set-EnrollmentRequiredFields -Worksheet $wsEnrollment -RowNumber 2

    $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentRowDirect", $sheetEnrollment, 2)
    $wsEnrollment.Activate() | Out-Null
    $wsEnrollment.Cells(2, 1).Select() | Out-Null

    $selectedRow = [int]$excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.LoadSelectedEnrollmentRowToBackend")
    Assert-Eq $selectedRow 2 "Selected enrollment row loader returned the wrong row number."
    Assert-Eq ([string]$excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.GetBackendValue", "fio")).Trim() $fio1.Trim() "Selected enrollment row loader did not move FIO into backend."

    $selectedExportPath = [string]$excel.Run("'$($workbook.Name)'!mdlEnrollmentOrderExport.ExportSelectedEnrollmentPackage")
    Assert-True (Test-Path -LiteralPath $selectedExportPath) "Selected enrollment row export did not create a .docx file."

    Write-Output "14a. Continue package preserves header and clears person"
    $continuePackageResult = [string]$excel.Run("'$($workbook.Name)'!codex_acceptance_probe.ProbeContinuePackage", 2)
    Assert-True ($continuePackageResult -like "ORD-SELECTED-CORE||990|YES") "Continue package did not preserve shared fields and clear personal fields."
    Assert-Eq ([string]$excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.GetBackendValue", "order_draft_id")).Trim() "ORD-SELECTED-CORE" "Continue package did not preserve order_draft_id in backend."
    Assert-Eq ([string]$excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.GetBackendValue", "fio")).Trim() "" "Continue package did not clear FIO in backend."

    Write-Output "15. Enrollment export without EDV omits section 2"
    $wsEnrollment.Cells.Clear()
    $excel.Run("'$($workbook.Name)'!mdlDataValidation.HealWorkbookStructure", $true)
    $wsEnrollment = Get-WorksheetByName -Workbook $workbook -SheetName $sheetEnrollment

    $wsEnrollment.Cells(2, 3).Value = $ln1.Trim()
    $wsEnrollment.Cells(2, 7).Value = "05.07.2026"
    $wsEnrollment.Cells(2, 8).Value = "911"
    $wsEnrollment.Cells(2, 10).Value = "05.07.2026"
    $wsEnrollment.Cells(2, 11).Value = "05.07.2026"
    $wsEnrollment.Cells(2, 13).Value = "report no edv"
    $wsEnrollment.Cells(2, 14).Value = "assignment no edv"
    $wsEnrollment.Cells(2, 34).Value = "5"
    $wsEnrollment.Cells(2, 35).Value = "20000"
    $wsEnrollment.Cells(2, 38).Value = "пункт отбора"
    $wsEnrollment.Cells(2, 64).Value = "YES"
    $wsEnrollment.Cells(2, 65).Value = "25"
    $wsEnrollment.Cells(2, 66).Value = "05.07.2026"
    $wsEnrollment.Cells(2, 67).Value = "31.12.2026"
    $wsEnrollment.Cells(2, 68).Value = "премирование"
    $wsEnrollment.Cells(2, 78).Value = "NO"
    Set-EnrollmentRequiredFields -Worksheet $wsEnrollment -RowNumber 2

    $excel.Run("'$($workbook.Name)'!mdlEnrollmentWorkflow.RefreshEnrollmentRowDirect", $sheetEnrollment, 2)
    $exportPathNoEdv = [string]$excel.Run("'$($workbook.Name)'!mdlEnrollmentOrderExport.ExportEnrollmentOrderByRow", 2)
    Assert-True (Test-Path -LiteralPath $exportPathNoEdv) "Enrollment export without EDV did not create a .docx file."

    $docTextNoEdv = Get-DocxText -Path $exportPathNoEdv
    Assert-True ($docTextNoEdv -notlike "*§ 2*") "Enrollment export without EDV should not include section 2."

    Remove-TestProbeModule -Workbook $workbook
    $workbook.Close($false)
    $workbook = $null
    $excel.Quit()
    $excel = $null

    Write-Output "ACCEPTANCE_SMOKE_OK"
}
finally {
    if ($workbook -ne $null) {
        Remove-TestProbeModule -Workbook $workbook
        try { $workbook.Close($false) } catch {}
    }
    if ($excel -ne $null) {
        try { $excel.Quit() } catch {}
    }
}
