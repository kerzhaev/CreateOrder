[CmdletBinding()]
param(
    [string]$WorkbookPath,
    [string]$OutputDirectory,
    [string]$TargetComponentName = 'frmGroupedPersonnelOrderWizard'
)

if ([string]::IsNullOrWhiteSpace($WorkbookPath)) { $WorkbookPath = Join-Path $PSScriptRoot '..\CreateOrder.xlsm' }
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) { $OutputDirectory = Join-Path $PSScriptRoot '..\CreateOrder.xlsm.modules' }

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Add-Control {
    param([object]$Designer, [string]$ClassName, [string]$Name, [double]$Left, [double]$Top, [double]$Width, [double]$Height, [string]$Caption = '')
    $control = $Designer.Controls.Add($ClassName, $Name, $true)
    $control.Left = $Left; $control.Top = $Top; $control.Width = $Width; $control.Height = $Height
    if ($Caption) { $control.Caption = $Caption }
    return $control
}

function Add-TextBox {
    param([object]$Designer, [string]$Name, [double]$Left, [double]$Top, [double]$Width, [double]$Height, [bool]$Locked = $false, [bool]$MultiLine = $false)
    $control = Add-Control $Designer 'Forms.TextBox.1' $Name $Left $Top $Width $Height
    $control.Locked = $Locked
    $control.MultiLine = $MultiLine
    if ($MultiLine) { $control.ScrollBars = 3; $control.EnterKeyBehavior = $true; $control.WordWrap = $false }
    return $control
}

if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Excel is running. Close Excel before generating the grouped-order form.' }
$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedOutput = [IO.Path]::GetFullPath($OutputDirectory)
$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$frmOutput = Join-Path $resolvedOutput ($TargetComponentName + '.frm')
$frxOutput = Join-Path $resolvedOutput ($TargetComponentName + '.frx')
$manifestOutput = Join-Path $resolvedOutput ($TargetComponentName + '.layout.csv')
foreach ($path in @($frmOutput, $frxOutput, $manifestOutput)) { if (Test-Path -LiteralPath $path) { throw "Target artifact already exists: $path" } }

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$buildDirectory = Join-Path $projectRoot "Trash\grouped-personnel-order-designer-build-$stamp"
New-Item -ItemType Directory -Path $buildDirectory -Force | Out-Null
$buildWorkbook = Join-Path $buildDirectory 'CreateOrder.grouped-order-designer-build.xlsm'
Copy-Item -LiteralPath $resolvedWorkbook -Destination $buildWorkbook -Force

$formCode = @'
Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = t("personnel.grouped.title", "Единый кадровый приказ")
    lblDescription.Caption = t("personnel.grouped.description", "Выберите сохранённые EventID через запятую. Порядок строк и параграфов будет сохранён.")
    lblSelection.Caption = t("personnel.grouped.selection", "EventID (необязательно)")
    lblReadOnly.Caption = t("personnel.grouped.read_only", "Проверка не изменяет реестры; экспорт блокируется при неполной записи.")
    btnPreview.Caption = t("personnel.grouped.preview", "Проверить")
    btnExport.Caption = t("personnel.grouped.export", "Сформировать DOCX")
    btnClose.Caption = t("common.close", "Закрыть")
    lblSummary.Caption = t("personnel.grouped.not_loaded", "Проверка ещё не запускалась.")
    txtReport.Text = vbNullString
End Sub

Private Sub btnPreview_Click()
    On Error GoTo Failed
    txtReport.Text = mdlGroupedPersonnelOrderExport.BuildGroupedPersonnelOrderReport(Trim$(txtEventIDs.Text))
    If Left$(txtReport.Text, 3) = "OK|" Then
        lblSummary.Caption = t("personnel.grouped.valid", "Данные готовы к формированию DOCX.")
    Else
        lblSummary.Caption = t("personnel.grouped.invalid", "Есть ошибки. Исправьте строки, отмеченные в отчёте.")
    End If
    Exit Sub
Failed:
    lblSummary.Caption = tf("personnel.grouped.failed", "Проверка не выполнена: {error}", "{error}", Err.Description)
    txtReport.Text = vbNullString
End Sub

Private Sub btnExport_Click()
    Dim outputPath As String
    On Error GoTo Failed
    outputPath = mdlGroupedPersonnelOrderExport.ExportGroupedPersonnelOrder(Trim$(txtEventIDs.Text))
    lblSummary.Caption = tf("personnel.grouped.exported", "DOCX сформирован: {path}", "{path}", outputPath)
    txtReport.Text = mdlGroupedPersonnelOrderExport.BuildGroupedPersonnelOrderReport(Trim$(txtEventIDs.Text))
    Exit Sub
Failed:
    lblSummary.Caption = tf("personnel.grouped.failed", "Экспорт не выполнен: {error}", "{error}", Err.Description)
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
'@

$excel = $null
$book = $null
$component = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($buildWorkbook, 0, $false)
    $component = $book.VBProject.VBComponents.Add(3)
    $component.Name = $TargetComponentName
    $designer = $component.Designer
    $component.Properties.Item('Caption').Value = 'Единый кадровый приказ'
    $component.Properties.Item('Width').Value = 720
    $component.Properties.Item('Height').Value = 520
    $component.Properties.Item('StartUpPosition').Value = 1

    $null = Add-Control $designer 'Forms.Label.1' 'lblTitle' 12 10 680 24 'Единый кадровый приказ'
    $null = Add-Control $designer 'Forms.Label.1' 'lblDescription' 12 38 680 30 'Выберите сохранённые EventID через запятую. Порядок строк и параграфов будет сохранён.'
    $null = Add-Control $designer 'Forms.Label.1' 'lblSelection' 12 80 135 18 'EventID (необязательно)'
    $null = Add-TextBox $designer 'txtEventIDs' 154 78 410 20
    $null = Add-Control $designer 'Forms.CommandButton.1' 'btnPreview' 575 76 90 24 'Проверить'
    $null = Add-Control $designer 'Forms.Label.1' 'lblSummary' 12 112 680 20 'Проверка ещё не запускалась.'
    $null = Add-TextBox $designer 'txtReport' 12 140 680 285 $true $true
    $null = Add-Control $designer 'Forms.Label.1' 'lblReadOnly' 12 435 680 22 'Проверка не изменяет реестры; экспорт блокируется при неполной записи.'
    $null = Add-Control $designer 'Forms.CommandButton.1' 'btnExport' 12 468 150 26 'Сформировать DOCX'
    $null = Add-Control $designer 'Forms.CommandButton.1' 'btnClose' 610 468 82 26 'Закрыть'
    [void]$component.CodeModule.AddFromString($formCode)
    if (-not [bool]$designer.Controls.Item('txtReport').Locked) { throw 'Report textbox must be locked.' }
    if (-not [bool]$designer.Controls.Item('txtReport').MultiLine) { throw 'Report textbox must be multiline.' }
    New-Item -ItemType Directory -Path $resolvedOutput -Force | Out-Null
    [void]$component.Export($frmOutput)
    if (-not (Test-Path -LiteralPath $frxOutput)) { throw "Designer export did not create FRX: $frxOutput" }

    $controls = @(
        @{Name='lblTitle'; Type='Label'; Parent='root'; Left=12; Top=10; Width=680; Height=24; Locked=$false; Multiline=$false},
        @{Name='lblDescription'; Type='Label'; Parent='root'; Left=12; Top=38; Width=680; Height=30; Locked=$false; Multiline=$false},
        @{Name='lblSelection'; Type='Label'; Parent='root'; Left=12; Top=80; Width=135; Height=18; Locked=$false; Multiline=$false},
        @{Name='txtEventIDs'; Type='TextBox'; Parent='root'; Left=154; Top=78; Width=410; Height=20; Locked=$false; Multiline=$false},
        @{Name='btnPreview'; Type='CommandButton'; Parent='root'; Left=575; Top=76; Width=90; Height=24; Locked=$false; Multiline=$false},
        @{Name='lblSummary'; Type='Label'; Parent='root'; Left=12; Top=112; Width=680; Height=20; Locked=$false; Multiline=$false},
        @{Name='txtReport'; Type='TextBox'; Parent='root'; Left=12; Top=140; Width=680; Height=285; Locked=$true; Multiline=$true},
        @{Name='lblReadOnly'; Type='Label'; Parent='root'; Left=12; Top=435; Width=680; Height=22; Locked=$false; Multiline=$false},
        @{Name='btnExport'; Type='CommandButton'; Parent='root'; Left=12; Top=468; Width=150; Height=26; Locked=$false; Multiline=$false},
        @{Name='btnClose'; Type='CommandButton'; Parent='root'; Left=610; Top=468; Width=82; Height=26; Locked=$false; Multiline=$false}
    )
    $controls | ForEach-Object { [pscustomobject]$_ } | Export-Csv -LiteralPath $manifestOutput -NoTypeInformation -Encoding UTF8
    $book.Close($false); $book = $null
    $excel.Quit(); $excel = $null
}
finally {
    if ($null -ne $book) { try { $book.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
for ($attempt = 1; $attempt -le 120 -and (Get-Process EXCEL -ErrorAction SilentlyContinue); $attempt++) { Start-Sleep -Milliseconds 500 }
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Office process remained after grouped-order form generation.' }

$frmBytes = [IO.File]::ReadAllBytes($frmOutput)
$frmText = [Text.Encoding]::GetEncoding(1251).GetString($frmBytes)
if ($frmText -notmatch ('Attribute VB_Name = "' + [regex]::Escape($TargetComponentName) + '"')) { throw 'Exported form has an unexpected VB_Name.' }
foreach ($handler in @('btnPreview_Click', 'btnExport_Click', 'btnClose_Click')) { if ($frmText -notmatch ('Private Sub ' + $handler + '\(\)')) { throw "Exported form is missing $handler." } }
if ($frmText -match '(?<!\r)\n') { throw 'Exported form must use CRLF line endings.' }
Write-Output "GROUPED_PERSONNEL_ORDER_DESIGNER_OK|form=$frmOutput|frx=$frxOutput|manifest=$manifestOutput"
