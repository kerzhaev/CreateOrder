[CmdletBinding()]
param(
    [string]$WorkbookPath,
    [string]$SourceDirectory,
    [string]$SourceComponentName = 'frmEnrollmentWizard',
    [string]$TargetComponentName = 'frmEnrollmentWizardV2'
)

if ([string]::IsNullOrWhiteSpace($WorkbookPath)) { $WorkbookPath = Join-Path $PSScriptRoot '..\CreateOrder.xlsm' }
if ([string]::IsNullOrWhiteSpace($SourceDirectory)) { $SourceDirectory = Join-Path $PSScriptRoot '..\CreateOrder.xlsm.modules' }

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-ConnectLog {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('DEBUG', 'INFO', 'WARN', 'ERROR')][string]$Level,
        [Parameter(Mandatory = $true)][string]$Message,
        [hashtable]$Context = @{}
    )
    $payload = [ordered]@{
        timestamp = (Get-Date).ToString('o')
        level = $Level
        operation = 'Connect-EnrollmentWizardV2Logic'
        message = $Message
    }
    foreach ($key in $Context.Keys) { $payload[$key] = $Context[$key] }
    $line = $payload | ConvertTo-Json -Compress -Depth 5
    if ($Level -eq 'DEBUG') { Write-Verbose $line }
    elseif ($Level -eq 'WARN') { Write-Warning $line }
    elseif ($Level -eq 'ERROR') { Write-Error $line }
    else { Write-Host $line }
}

function Release-ComObject {
    param([object]$Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

if (-not ('EnrollmentV2ConnectNativeMethods' -as [type])) {
    Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class EnrollmentV2ConnectNativeMethods {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@
}

function Get-ExcelProcessId {
    param([Parameter(Mandatory = $true)][object]$ExcelApplication)
    [uint32]$processId = 0
    [void][EnrollmentV2ConnectNativeMethods]::GetWindowThreadProcessId([IntPtr]$ExcelApplication.Hwnd, [ref]$processId)
    return [int]$processId
}

function Stop-OwnedExcelProcessIfNeeded {
    param([int]$ProcessId)
    if ($ProcessId -le 0) { return }
    for ($attempt = 1; $attempt -le 10; $attempt++) {
        if (-not (Get-Process -Id $ProcessId -ErrorAction SilentlyContinue)) { return }
        Start-Sleep -Milliseconds 250
    }
    $process = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
    if ($process -and $process.ProcessName -eq 'EXCEL') {
        Write-ConnectLog WARN 'Excel did not exit after COM Quit; stopping only the process created by the binding probe.' @{ processId = $ProcessId }
        Stop-Process -Id $ProcessId -Force
    }
}

function Normalize-ControlType {
    param([Parameter(Mandatory = $true)][string]$TypeName)
    $typeMap = @{
        IMdcText = 'TextBox'; IMdcList = 'ListBox'; IMdcCombo = 'ComboBox'; IMdcCheckBox = 'CheckBox'
        IMdcLabel = 'Label'; ILabelControl = 'Label'; IMdcCommandButton = 'CommandButton'; ICommandButton = 'CommandButton'
        IMdcFrame = 'Frame'; IOptionFrame = 'Frame'; IMdcMultiPage = 'MultiPage'; IMultiPage = 'MultiPage'
        IMdcPage = 'Page'; IPage = 'Page'
    }
    if ($typeMap.ContainsKey($TypeName)) { return $typeMap[$TypeName] }
    return $TypeName
}

function Replace-VbaProcedure {
    param(
        [Parameter(Mandatory = $true)][string]$Code,
        [Parameter(Mandatory = $true)][string]$ProcedureName,
        [AllowEmptyString()][string]$Replacement
    )
    $pattern = '(?ms)^(?:Private|Public|Friend)\s+(?:Sub|Function)\s+' + [regex]::Escape($ProcedureName) + '\b.*?^End\s+(?:Sub|Function)\s*\r?\n?'
    $matches = [regex]::Matches($Code, $pattern)
    if ($matches.Count -ne 1) { throw "Expected exactly one procedure '$ProcedureName'; found $($matches.Count)." }
    return [regex]::Replace($Code, $pattern, $Replacement, 1)
}

function Escape-VbaString {
    param([AllowEmptyString()][string]$Value)
    return $Value.Replace('"', '""')
}

$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedSource = [IO.Path]::GetFullPath($SourceDirectory)
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Excel must be closed before connecting V2 logic.' }

$v1Path = Join-Path $resolvedSource ($SourceComponentName + '.frm')
$v2Path = Join-Path $resolvedSource ($TargetComponentName + '.frm')
$currentManifestPath = Join-Path $resolvedSource ($TargetComponentName + '.layout.csv')
$bindingPath = Join-Path $resolvedSource ($TargetComponentName + '.bindings.csv')
foreach ($path in @($v1Path, $v2Path, $currentManifestPath)) {
    if (-not (Test-Path -LiteralPath $path)) { throw "Missing required artifact: $path" }
}

$baselineManifestPath = Get-ChildItem -LiteralPath (Join-Path $projectRoot 'CreateOrderBackups') -Directory -Filter 'enrollment-designer-v2-owner-layout-*' |
    Sort-Object Name -Descending |
    ForEach-Object { Join-Path $_.FullName ($TargetComponentName + '.layout.csv') } |
    Where-Object { Test-Path -LiteralPath $_ } |
    Select-Object -First 1
if (-not $baselineManifestPath) { throw 'Could not find the pre-owner V2 layout manifest in CreateOrderBackups.' }

$encoding = [Text.Encoding]::GetEncoding(1251)
$v1Text = $encoding.GetString([IO.File]::ReadAllBytes($v1Path))
$v2Text = $encoding.GetString([IO.File]::ReadAllBytes($v2Path))
$v1CodeIndex = $v1Text.IndexOf('Option Explicit', [StringComparison]::Ordinal)
$v2CodeIndex = $v2Text.IndexOf('Option Explicit', [StringComparison]::Ordinal)
if ($v1CodeIndex -lt 0 -or $v2CodeIndex -lt 0) { throw 'Could not locate the VBA code section in a form export.' }
$v1Code = $v1Text.Substring($v1CodeIndex)
$v2Header = $v2Text.Substring(0, $v2CodeIndex)

$declarations = @()
foreach ($match in [regex]::Matches($v1Code, '(?m)^Private\s+([A-Za-z_][A-Za-z0-9_]*)(?:\(1 To (\d+)\))?\s+As Object\s*$')) {
    $declarations += [pscustomobject]@{
        name = $match.Groups[1].Value
        upperBound = if ($match.Groups[2].Success) { [int]$match.Groups[2].Value } else { 0 }
        declaration = $match.Value
        withEvents = $false
    }
}
foreach ($match in [regex]::Matches($v1Code, '(?m)^Private WithEvents\s+([A-Za-z_][A-Za-z0-9_]*)\s+As MSForms\.(\w+)\s*$')) {
    $declarations += [pscustomobject]@{
        name = $match.Groups[1].Value
        upperBound = 0
        declaration = $match.Value
        withEvents = $true
    }
}

$probeEntries = @()
foreach ($declaration in $declarations) {
    if ($declaration.upperBound -gt 0) {
        for ($index = 1; $index -le $declaration.upperBound; $index++) {
            $probeEntries += [pscustomobject]@{ label = "$($declaration.name)($index)"; expression = "$($declaration.name)($index)"; baseName = $declaration.name; index = $index }
        }
    } else {
        $probeEntries += [pscustomobject]@{ label = $declaration.name; expression = $declaration.name; baseName = $declaration.name; index = 0 }
    }
}

$probeLines = [Collections.Generic.List[string]]::new()
$probeLines.Add('')
$probeLines.Add('Public Function ExportV2DesignerBindings() As String')
$probeLines.Add('    Dim result As String')
foreach ($entry in $probeEntries) {
    $probeLines.Add(('    result = result & V2BindingLine("{0}", {1}) & vbLf' -f $entry.label, $entry.expression))
}
$probeLines.Add('    ExportV2DesignerBindings = result')
$probeLines.Add('End Function')
$probeLines.Add('')
$probeLines.Add('Private Function V2BindingLine(ByVal variableName As String, ByVal controlValue As Object) As String')
$probeLines.Add('    On Error GoTo BindingError')
$probeLines.Add('    If controlValue Is Nothing Then')
$probeLines.Add('        V2BindingLine = variableName & "|<NOTHING>"')
$probeLines.Add('        Exit Function')
$probeLines.Add('    End If')
$probeLines.Add('    If TypeName(controlValue) = "Page" Then')
$probeLines.Add('        V2BindingLine = variableName & "|" & CStr(controlValue.Name) & "|Page|" & CStr(controlValue.Parent.Name) & "|0|0|0|0"')
$probeLines.Add('        Exit Function')
$probeLines.Add('    End If')
$probeLines.Add('    V2BindingLine = variableName & "|" & CStr(controlValue.Name) & "|" & TypeName(controlValue) & "|" & CStr(controlValue.Parent.Name) & "|" & _')
$probeLines.Add('        CStr(controlValue.Left) & "|" & CStr(controlValue.Top) & "|" & CStr(controlValue.Width) & "|" & CStr(controlValue.Height)')
$probeLines.Add('    Exit Function')
$probeLines.Add('BindingError:')
$probeLines.Add('    V2BindingLine = variableName & "|<ERROR>|" & CStr(Err.Number) & "|" & Err.Description')
$probeLines.Add('End Function')

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$probeDirectory = Join-Path $projectRoot ("Trash\enrollment-designer-v2-binding-probe-$stamp")
New-Item -ItemType Directory -Path $probeDirectory -Force | Out-Null
$probeWorkbook = Join-Path $probeDirectory 'CreateOrder.v2-binding-probe.xlsm'
Copy-Item -LiteralPath $resolvedWorkbook -Destination $probeWorkbook

$excel = $null
$excelProcessId = 0
$book = $null
$formComponent = $null
$probeComponent = $null
$probeSheet = $null
$bindingText = ''
try {
    Write-ConnectLog INFO 'Running an isolated V1 binding probe against the accepted runtime form.' @{ workbook = $probeWorkbook; entries = $probeEntries.Count }
    $excel = New-Object -ComObject Excel.Application
    $excelProcessId = Get-ExcelProcessId -ExcelApplication $excel
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($probeWorkbook, 0, $false)
    $formComponent = $book.VBProject.VBComponents.Item($SourceComponentName)
    $formComponent.CodeModule.AddFromString(($probeLines -join "`r`n"))
    $probeComponent = $book.VBProject.VBComponents.Add(1)
    $probeComponent.Name = 'modEnrollmentV2BindingProbe'
    $probeComponent.CodeModule.AddFromString(@'
Option Explicit
Public Sub RunEnrollmentV2BindingProbe()
    Dim bindingText As String
    Load frmEnrollmentWizard
    bindingText = frmEnrollmentWizard.ExportV2DesignerBindings()
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("__V2Bindings").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Dim targetSheet As Worksheet
    Set targetSheet = ThisWorkbook.Worksheets.Add
    targetSheet.Name = "__V2Bindings"
    targetSheet.Range("A1").Value = bindingText
    Unload frmEnrollmentWizard
End Sub
'@)
    $excel.Run("'$($book.Name)'!modEnrollmentV2BindingProbe.RunEnrollmentV2BindingProbe")
    $probeSheet = $book.Worksheets.Item('__V2Bindings')
    $bindingText = [string]$probeSheet.Range('A1').Value2
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit() }
    Release-ComObject $probeSheet
    Release-ComObject $probeComponent
    Release-ComObject $formComponent
    Release-ComObject $book
    Release-ComObject $excel
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Stop-OwnedExcelProcessIfNeeded -ProcessId $excelProcessId
}
if (-not $bindingText) { throw 'The isolated V1 binding probe returned no data.' }

$baselineManifest = @(Import-Csv -LiteralPath $baselineManifestPath)
$currentManifest = @(Import-Csv -LiteralPath $currentManifestPath)
$bindings = @()
foreach ($line in @($bindingText -split "`n" | Where-Object { $_.Trim() })) {
    $parts = @($line.Trim() -split '\|')
    if ($parts.Count -ge 2 -and $parts[1] -eq '<ERROR>') { throw "V1 binding probe failed for $($parts[0]): [$($parts[2])] $($parts[3])" }
    if ($parts.Count -lt 8) { throw "Malformed binding probe line: $line" }
    if ($parts[1] -eq '<NOTHING>') { continue }
    $variableName = $parts[0]
    $sourceName = $parts[1]
    $controlType = Normalize-ControlType $parts[2]
    $parentName = $parts[3]
    $left = [double]::Parse($parts[4], [Globalization.CultureInfo]::CurrentCulture)
    $top = [double]::Parse($parts[5], [Globalization.CultureInfo]::CurrentCulture)
    $width = [double]::Parse($parts[6], [Globalization.CultureInfo]::CurrentCulture)
    $height = [double]::Parse($parts[7], [Globalization.CultureInfo]::CurrentCulture)
    $candidates = @($baselineManifest | Where-Object {
        $_.source_name -eq $sourceName -and $_.control_type -eq $controlType -and
        [Math]::Abs([double]$_.left - $left) -lt 0.1 -and [Math]::Abs([double]$_.top - $top) -lt 0.1 -and
        [Math]::Abs([double]$_.width - $width) -lt 0.1 -and [Math]::Abs([double]$_.height - $height) -lt 0.1
    })
    if ($candidates.Count -gt 1) {
        $parentCandidates = @($candidates | Where-Object { ([string]$_.container_path -split '/')[-1] -eq $parentName })
        if ($parentCandidates.Count -eq 1) { $candidates = $parentCandidates }
    }
    if ($variableName -eq 'btnSaveGenerateDynamic' -and $candidates.Count -eq 0) { continue }
    if ($candidates.Count -ne 1) {
        throw "Could not uniquely map $variableName ($sourceName, $controlType, $left/$top/$width/$height); candidates=$($candidates.Count)."
    }
    $candidate = $candidates[0]
    $bindings += [pscustomobject]@{
        variable_name = $variableName
        source_name = $sourceName
        designer_name = $candidate.designer_name
        control_type = $controlType
        container_path = $candidate.container_path
    }
}
$bindings | Export-Csv -LiteralPath $bindingPath -NoTypeInformation -Encoding utf8

$bindingByVariable = @{}
foreach ($binding in $bindings) { $bindingByVariable[$binding.variable_name] = $binding }
$generatedBindingLines = [Collections.Generic.List[string]]::new()
foreach ($entry in $probeEntries) {
    if (-not $bindingByVariable.ContainsKey($entry.label)) { continue }
    $binding = $bindingByVariable[$entry.label]
    if ($binding.control_type -eq 'Page') {
        $generatedBindingLines.Add(('    Set {0} = mpWizard.Pages("{1}")' -f $entry.expression, (Escape-VbaString $binding.designer_name)))
        continue
    }
    if ($entry.index -eq 0 -and $entry.baseName -eq $binding.designer_name) { continue }
    if ($entry.baseName -like 'btn*Dynamic' -and $binding.designer_name -eq $entry.baseName) { continue }
    $generatedBindingLines.Add(('    Set {0} = FindDesignerControl("{1}")' -f $entry.expression, (Escape-VbaString $binding.designer_name)))
}

$localizationLines = [Collections.Generic.List[string]]::new()
$v1Lines = @($v1Code -split "`r?`n")
foreach ($line in $v1Lines) {
    $match = [regex]::Match($line, 'Set\s+([A-Za-z_][A-Za-z0-9_]*)(?:\(i\))?\s*=\s*AddPage(TextBoxT|ComboBoxT|CheckBoxT)\([^,]+,\s*"([^"]+)",\s*"([^"]*)"')
    if (-not $match.Success) { continue }
    $baseName = $match.Groups[1].Value
    $builderKind = $match.Groups[2].Value
    $key = $match.Groups[3].Value
    $fallback = $match.Groups[4].Value
    $declaration = $declarations | Where-Object name -eq $baseName | Select-Object -First 1
    $indexes = if ($declaration -and $declaration.upperBound -gt 0) { 1..$declaration.upperBound } else { @(0) }
    foreach ($index in $indexes) {
        $variableName = if ($index -gt 0) { "$baseName($index)" } else { $baseName }
        if (-not $bindingByVariable.ContainsKey($variableName)) { continue }
        $controlBinding = $bindingByVariable[$variableName]
        $captionControlName = $controlBinding.designer_name
        if ($builderKind -ne 'CheckBoxT') {
            $controlRowIndex = [Array]::FindIndex($baselineManifest, [Predicate[object]]{ param($row) $row.designer_name -eq $controlBinding.designer_name })
            if ($controlRowIndex -le 0) { throw "Could not locate a label before $variableName." }
            $labelRow = $baselineManifest[$controlRowIndex - 1]
            if ($labelRow.control_type -ne 'Label' -or $labelRow.container_path -ne $controlBinding.container_path) {
                throw "Expected an adjacent label before $variableName; found $($labelRow.control_type)."
            }
            $captionControlName = $labelRow.designer_name
        }
        $localizationLines.Add(('    SetDesignerCaption "{0}", "{1}", "{2}", {3}' -f
            (Escape-VbaString $captionControlName), (Escape-VbaString $key), (Escape-VbaString $fallback), $index))
    }
}

$initializeProcedure = @'
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    ' Infrastructure is prepared by the public workflow command before Show.
    mdlHelper.EnsureStaffColumnsInitialized
    BindDesignerControls
    ConfigureSearchArea
    ConfigureWindow
    ConfigureButtons
    ApplyDesignerLocalization
    PopulateOperatorReferenceLists
    currentSourceMode = "manual"
    ReloadFromBackend
    lblStatus.Caption = t("enrollment.form.status.ready_to_pick", "Выберите сотрудника из листа 'Штат' или заполните карточку вручную. После выбора проверьте страницы мастера.")
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, "frmEnrollmentWizardV2.UserForm_Initialize", Err.Description
End Sub
'@

$bindingProcedure = @"
Private Sub BindDesignerControls()
$($generatedBindingLines -join "`r`n")
End Sub

Private Function FindDesignerControl(ByVal controlName As String) As Object
    Set FindDesignerControl = FindDesignerControlInContainer(Me, controlName)
    If FindDesignerControl Is Nothing Then Err.Raise 5, "frmEnrollmentWizardV2.FindDesignerControl", "Не найден design-time контрол: " & controlName
End Function

Private Function FindDesignerControlInContainer(ByVal containerHost As Object, ByVal controlName As String) As Object
    Dim controlItem As Object
    Dim pageItem As Object
    Dim foundControl As Object

    On Error Resume Next
    Set foundControl = containerHost.Controls.Item(controlName)
    On Error GoTo 0
    If Not foundControl Is Nothing Then
        Set FindDesignerControlInContainer = foundControl
        Exit Function
    End If

    For Each controlItem In containerHost.Controls
        If TypeName(controlItem) = "MultiPage" Then
            For Each pageItem In controlItem.Pages
                Set foundControl = FindDesignerControlInContainer(pageItem, controlName)
                If Not foundControl Is Nothing Then
                    Set FindDesignerControlInContainer = foundControl
                    Exit Function
                End If
            Next pageItem
        ElseIf TypeName(controlItem) = "Frame" Then
            Set foundControl = FindDesignerControlInContainer(controlItem, controlName)
            If Not foundControl Is Nothing Then
                Set FindDesignerControlInContainer = foundControl
                Exit Function
            End If
        End If
    Next controlItem
End Function
"@

$localizationProcedure = @"
Private Sub ApplyDesignerLocalization()
    Me.Caption = t("enrollment.form.title", "Мастер зачисления")
    pgEmployee.Caption = t("enrollment.page.employee", "Военнослужащий")
    pgDocs.Caption = t("enrollment.page.docs", "Документы и даты")
    pgMonthly.Caption = t("enrollment.page.monthly", "Ежемесячные выплаты")
    pgOneTime.Caption = t("enrollment.page.onetime", "Разовые выплаты и реквизиты")
    pgAdvanced.Caption = t("enrollment.page.advanced", "Основания выплат")
    pgExtras.Caption = t("enrollment.page.extras", "Иные выплаты")
    pgPreview.Caption = t("enrollment.page.preview", "Проверка и текст приказа")
$($localizationLines -join "`r`n")
End Sub

Private Sub SetDesignerCaption(ByVal controlName As String, ByVal localizationKey As String, ByVal fallbackText As String, Optional ByVal captionIndex As Long = 0)
    Dim targetControl As Object
    Dim resolvedCaption As String
    Set targetControl = FindDesignerControl(controlName)
    resolvedCaption = t(localizationKey, fallbackText)
    If captionIndex > 0 Then resolvedCaption = Replace`$(resolvedCaption, "{index}", CStr(captionIndex))
    targetControl.Caption = resolvedCaption
End Sub
"@

$configureWindow = @'
Private Sub ConfigureWindow()
    Me.Caption = t("enrollment.form.title", "Мастер зачисления")
    ConfigureInlineSearchUi
End Sub
'@
$configureInlineSearch = @'
Private Sub ConfigureInlineSearchUi()
    txtSearch.Visible = True
    txtSearch.ControlTipText = t("enrollment.form.search.tip", "Введите ФИО, личный или табельный номер.")
    btnLoadFromInlineSearchDynamic.Caption = t("enrollment.form.button.load_from_search", "Загрузить из поиска")
    btnLoadFromInlineSearchDynamic.Visible = True
    btnLoadFromInlineSearchDynamic.Enabled = False
    lstResults.Visible = True
End Sub
'@
$configureButtons = @'
Private Sub ConfigureButtons()
    btnSelect.Caption = t("enrollment.form.button.pick_from_staff", "Выбрать сотрудника из штата")
    btnCheckDynamic.Caption = t("enrollment.form.button.check", "Проверить и показать")
    btnSaveCardDynamic.Caption = t("enrollment.form.button.save", "Сохранить")
    btnExportPackageDynamic.Caption = t("enrollment.form.button.export", "Экспортировать Word")
    btnSaveContinueDynamic.Caption = t("enrollment.form.button.save_continue_package", "Следующий в пакете")
    btnClose.Caption = t("common.close", "Закрыть")
End Sub
'@

$targetCode = $v1Code
$targetCode = Replace-VbaProcedure $targetCode 'UserForm_Initialize' ($initializeProcedure + "`r`n`r`n" + $bindingProcedure + "`r`n" + $localizationProcedure + "`r`n")
$targetCode = Replace-VbaProcedure $targetCode 'ConfigureWindow' ($configureWindow + "`r`n")
$targetCode = Replace-VbaProcedure $targetCode 'ConfigureInlineSearchUi' ($configureInlineSearch + "`r`n")
$targetCode = Replace-VbaProcedure $targetCode 'ConfigureButtons' ($configureButtons + "`r`n")

$removeProcedures = @(
    'HideLegacyControls', 'EnsureDynamicActionButtons', 'CreateWizardUi', 'CreateEmployeePage', 'CreateDocsPage',
    'CreateMonthlyPage', 'CreateOneTimePage', 'CreateAdvancedPage', 'CreateExtrasPage', 'AddPageFrame',
    'CreatePreviewPage', 'RemoveDefaultWizardPages', 'ConfigureScrollablePage', 'AddPageSectionLabel',
    'AddPageComboBoxT', 'AddLabelToPage', 'AddPageTextBoxT', 'AddPageCheckBoxT', 'AddPageTextBox', 'AddPageCheckBox'
)
foreach ($procedureName in $removeProcedures) { $targetCode = Replace-VbaProcedure $targetCode $procedureName '' }

foreach ($declaration in $declarations | Where-Object withEvents) {
    if ($declaration.name -like 'btn*Dynamic') {
        $targetCode = $targetCode.Replace($declaration.declaration + "`r`n", '')
        $targetCode = $targetCode.Replace($declaration.declaration + "`n", '')
    }
}
foreach ($binding in $bindings) {
    if ($binding.variable_name -match '\(') { continue }
    if ($binding.variable_name -ne $binding.designer_name) { continue }
    if ($binding.control_type -eq 'Page') { continue }
    $declaration = $declarations | Where-Object name -eq $binding.variable_name | Select-Object -First 1
    if ($declaration) {
        $targetCode = $targetCode.Replace($declaration.declaration + "`r`n", '')
        $targetCode = $targetCode.Replace($declaration.declaration + "`n", '')
    }
}

$targetCode = $targetCode.Replace('frmEnrollmentWizard.', 'frmEnrollmentWizardV2.')
$targetCode = [regex]::Replace($targetCode, "`r?`n", "`r`n")
if ($targetCode.Contains('Controls.Add(')) { throw 'Generated V2 code still contains Controls.Add.' }
if ([regex]::IsMatch($targetCode, '\.(Left|Top|Width|Height)\s*=')) { throw 'Generated V2 code still contains a geometry assignment.' }
if (-not $targetCode.Contains('Private Sub BindDesignerControls()')) { throw 'Generated V2 code is missing the designer binding procedure.' }

$targetText = [regex]::Replace($v2Header, "`r?`n", "`r`n") + $targetCode
[IO.File]::WriteAllText($v2Path, $targetText, $encoding)

Write-ConnectLog INFO 'Connected V1 behavior and localization to V2 design-time controls.' @{
    bindings = $bindings.Count
    localizedControls = $localizationLines.Count
    source = $v1Path
    target = $v2Path
    bindingManifest = $bindingPath
    baselineManifest = $baselineManifestPath
}
