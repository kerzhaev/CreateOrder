[CmdletBinding()]
param(
    [string]$WorkbookPath,
    [string]$OutputDirectory,
    [string]$TargetComponentName = 'frmDataIntegrityCenter'
)

if ([string]::IsNullOrWhiteSpace($WorkbookPath)) { $WorkbookPath = Join-Path $PSScriptRoot '..\CreateOrder.xlsm' }
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) { $OutputDirectory = Join-Path $PSScriptRoot '..\CreateOrder.xlsm.modules' }

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Read-VbaText {
    param([Parameter(Mandatory = $true)][string]$Path)
    [IO.File]::ReadAllText($Path, [Text.Encoding]::GetEncoding(1251))
}

function Add-Label {
    param([object]$Designer, [string]$Name, [double]$Left, [double]$Top, [double]$Width, [double]$Height, [string]$Caption)
    $control = $Designer.Controls.Add('Forms.Label.1', $Name, $true)
    $control.Left = $Left; $control.Top = $Top; $control.Width = $Width; $control.Height = $Height
    $control.Caption = $Caption
    $control.WordWrap = $true
    $control
}

function Add-Button {
    param([object]$Designer, [string]$Name, [double]$Left, [double]$Top, [double]$Width, [double]$Height, [string]$Caption)
    $control = $Designer.Controls.Add('Forms.CommandButton.1', $Name, $true)
    $control.Left = $Left; $control.Top = $Top; $control.Width = $Width; $control.Height = $Height
    $control.Caption = $Caption
    $control
}

function Add-ComboBox {
    param([object]$Designer, [string]$Name, [double]$Left, [double]$Top, [double]$Width, [double]$Height)
    $control = $Designer.Controls.Add('Forms.ComboBox.1', $Name, $true)
    $control.Left = $Left; $control.Top = $Top; $control.Width = $Width; $control.Height = $Height
    $control
}

function Add-TextBox {
    param([object]$Designer, [string]$Name, [double]$Left, [double]$Top, [double]$Width, [double]$Height)
    $control = $Designer.Controls.Add('Forms.TextBox.1', $Name, $true)
    $control.Left = $Left; $control.Top = $Top; $control.Width = $Width; $control.Height = $Height
    $control.MultiLine = $true
    $control.ScrollBars = 3
    $control.Locked = $true
    $control.EnterKeyBehavior = $true
    $control
}

if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Excel is running. Close Excel before generating the data integrity form.' }
$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedOutput = [IO.Path]::GetFullPath($OutputDirectory)
$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$frmOutput = Join-Path $resolvedOutput ($TargetComponentName + '.frm')
$frxOutput = Join-Path $resolvedOutput ($TargetComponentName + '.frx')
$manifestOutput = Join-Path $resolvedOutput ($TargetComponentName + '.layout.csv')
foreach ($path in @($frmOutput, $frxOutput, $manifestOutput)) { if (Test-Path -LiteralPath $path) { throw "Target artifact already exists: $path" } }

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$buildDirectory = Join-Path $projectRoot "Trash\data-integrity-designer-build-$stamp"
New-Item -ItemType Directory -Path $buildDirectory -Force | Out-Null
$buildWorkbook = Join-Path $buildDirectory 'CreateOrder.data-integrity-designer-build.xlsm'
Copy-Item -LiteralPath $resolvedWorkbook -Destination $buildWorkbook

$formCode = @'
Option Explicit

Private mReport As String

Private Sub UserForm_Initialize()
    Me.Caption = t("integrity.form.title", "Data integrity center")
    lblDescription.Caption = t("integrity.form.description", "Read-only diagnostics of personnel registries.")
    lblSeverity.Caption = t("integrity.form.severity", "Severity")
    lblCategory.Caption = t("integrity.form.category", "Category")
    lblReadOnly.Caption = t("integrity.form.readonly", "Read-only: no correction is performed.")
    btnScan.Caption = t("integrity.form.scan", "Scan")
    btnClose.Caption = t("integrity.form.close", "Close")
    cboSeverity.Clear
    cboSeverity.AddItem t("integrity.form.all", "ALL")
    cboSeverity.AddItem "ERROR"
    cboSeverity.AddItem "WARNING"
    cboSeverity.ListIndex = 0
    cboCategory.Clear
    cboCategory.AddItem t("integrity.form.all", "ALL")
    cboCategory.ListIndex = 0
    lblSummary.Caption = t("integrity.form.not_scanned", "Not scanned.")
    txtFindings.Text = vbNullString
    mReport = vbNullString
End Sub

Private Sub btnScan_Click()
    On Error GoTo Failed
    mReport = mdlPersonnelDataIntegrity.BuildPersonnelDataIntegrityReport()
    PopulateCategories
    RenderFiltered
    lblSummary.Caption = Replace$(Split(mReport, vbCrLf)(0), "Integrity scan | ", vbNullString)
    Exit Sub
Failed:
    mReport = vbNullString
    txtFindings.Text = t("integrity.form.scan_failed", "Integrity scan failed.")
    lblSummary.Caption = t("integrity.form.scan_failed", "Integrity scan failed.")
End Sub

Private Sub cboSeverity_Change()
    RenderFiltered
End Sub

Private Sub cboCategory_Change()
    RenderFiltered
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub PopulateCategories()
    Dim lines As Variant
    Dim lineText As String
    Dim parts As Variant
    Dim i As Long
    Dim categoryText As String

    cboCategory.Clear
    cboCategory.AddItem t("integrity.form.all", "ALL")
    If Len(mReport) = 0 Then Exit Sub
    lines = Split(mReport, vbCrLf)
    For i = 1 To UBound(lines)
        lineText = CStr(lines(i))
        parts = Split(lineText, " | ")
        If UBound(parts) >= 1 Then
            categoryText = CStr(parts(1))
            If Not CategoryExists(categoryText) Then cboCategory.AddItem categoryText
        End If
    Next i
    cboCategory.ListIndex = 0
End Sub

Private Function CategoryExists(ByVal categoryText As String) As Boolean
    Dim i As Long
    For i = 0 To cboCategory.ListCount - 1
        If StrComp(CStr(cboCategory.List(i)), categoryText, vbTextCompare) = 0 Then
            CategoryExists = True
            Exit Function
        End If
    Next i
End Function

Private Sub RenderFiltered()
    Dim lines As Variant
    Dim lineText As String
    Dim parts As Variant
    Dim severityFilter As String
    Dim categoryFilter As String
    Dim resultText As String
    Dim i As Long

    If Len(mReport) = 0 Then
        txtFindings.Text = vbNullString
        Exit Sub
    End If
    severityFilter = UCase$(CStr(cboSeverity.Value))
    categoryFilter = UCase$(CStr(cboCategory.Value))
    lines = Split(mReport, vbCrLf)
    resultText = CStr(lines(0))
    For i = 1 To UBound(lines)
        lineText = CStr(lines(i))
        parts = Split(lineText, " | ")
        If UBound(parts) >= 1 Then
            If (severityFilter = "ALL" Or UCase$(CStr(parts(0))) = severityFilter) And (categoryFilter = "ALL" Or UCase$(CStr(parts(1))) = categoryFilter) Then resultText = resultText & vbCrLf & lineText
        End If
    Next i
    txtFindings.Text = resultText
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
    $component.Properties.Item('Caption').Value = 'Data integrity center'
    $component.Properties.Item('Width').Value = 650
    $component.Properties.Item('Height').Value = 430
    $component.Properties.Item('StartUpPosition').Value = 1

    $null = Add-Label -Designer $designer -Name 'lblTitle' -Left 12 -Top 10 -Width 610 -Height 24 -Caption 'Data integrity center'
    $null = Add-Label -Designer $designer -Name 'lblDescription' -Left 12 -Top 38 -Width 610 -Height 30 -Caption 'Read-only diagnostics of personnel registries.'
    $null = Add-Label -Designer $designer -Name 'lblSeverity' -Left 12 -Top 78 -Width 70 -Height 18 -Caption 'Severity'
    $null = Add-ComboBox -Designer $designer -Name 'cboSeverity' -Left 86 -Top 76 -Width 120 -Height 20
    $null = Add-Label -Designer $designer -Name 'lblCategory' -Left 220 -Top 78 -Width 70 -Height 18 -Caption 'Category'
    $null = Add-ComboBox -Designer $designer -Name 'cboCategory' -Left 294 -Top 76 -Width 220 -Height 20
    $null = Add-Label -Designer $designer -Name 'lblSummary' -Left 12 -Top 108 -Width 610 -Height 18 -Caption 'Not scanned.'
    $null = Add-TextBox -Designer $designer -Name 'txtFindings' -Left 12 -Top 134 -Width 610 -Height 230
    $null = Add-Label -Designer $designer -Name 'lblReadOnly' -Left 12 -Top 372 -Width 420 -Height 20 -Caption 'Read-only: no correction is performed.'
    $null = Add-Button -Designer $designer -Name 'btnScan' -Left 452 -Top 368 -Width 80 -Height 26 -Caption 'Scan'
    $null = Add-Button -Designer $designer -Name 'btnClose' -Left 540 -Top 368 -Width 82 -Height 26 -Caption 'Close'

    $codeModule = $component.CodeModule
    $null = $codeModule.AddFromString($formCode)
    if (-not [bool]$designer.Controls.Item('txtFindings').Locked) { throw 'Designer textbox must be locked.' }
    if (-not [bool]$designer.Controls.Item('txtFindings').MultiLine) { throw 'Designer textbox must be multiline.' }
    New-Item -ItemType Directory -Path $resolvedOutput -Force | Out-Null
    $null = $component.Export($frmOutput)
    if (-not (Test-Path -LiteralPath $frxOutput)) { throw "Designer export did not create FRX: $frxOutput" }

    $manifest = @(
        [pscustomobject]@{ designer_name = 'lblTitle'; control_type = 'Label'; parent = 'root'; left = 12; top = 10; width = 610; height = 24; locked = $false; multiline = $false },
        [pscustomobject]@{ designer_name = 'lblDescription'; control_type = 'Label'; parent = 'root'; left = 12; top = 38; width = 610; height = 30; locked = $false; multiline = $false },
        [pscustomobject]@{ designer_name = 'lblSeverity'; control_type = 'Label'; parent = 'root'; left = 12; top = 78; width = 70; height = 18; locked = $false; multiline = $false },
        [pscustomobject]@{ designer_name = 'cboSeverity'; control_type = 'ComboBox'; parent = 'root'; left = 86; top = 76; width = 120; height = 20; locked = $false; multiline = $false },
        [pscustomobject]@{ designer_name = 'lblCategory'; control_type = 'Label'; parent = 'root'; left = 220; top = 78; width = 70; height = 18; locked = $false; multiline = $false },
        [pscustomobject]@{ designer_name = 'cboCategory'; control_type = 'ComboBox'; parent = 'root'; left = 294; top = 76; width = 220; height = 20; locked = $false; multiline = $false },
        [pscustomobject]@{ designer_name = 'lblSummary'; control_type = 'Label'; parent = 'root'; left = 12; top = 108; width = 610; height = 18; locked = $false; multiline = $false },
        [pscustomobject]@{ designer_name = 'txtFindings'; control_type = 'TextBox'; parent = 'root'; left = 12; top = 134; width = 610; height = 230; locked = $true; multiline = $true },
        [pscustomobject]@{ designer_name = 'lblReadOnly'; control_type = 'Label'; parent = 'root'; left = 12; top = 372; width = 420; height = 20; locked = $false; multiline = $false },
        [pscustomobject]@{ designer_name = 'btnScan'; control_type = 'CommandButton'; parent = 'root'; left = 452; top = 368; width = 80; height = 26; locked = $false; multiline = $false },
        [pscustomobject]@{ designer_name = 'btnClose'; control_type = 'CommandButton'; parent = 'root'; left = 540; top = 368; width = 82; height = 26; locked = $false; multiline = $false }
    )
    $manifest | Export-Csv -LiteralPath $manifestOutput -NoTypeInformation -Encoding UTF8
    $book.Close($false)
    $book = $null
    $excel.Quit()
    $excel = $null
} finally {
    if ($null -ne $book) { try { $book.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

for ($attempt = 1; $attempt -le 60 -and (Get-Process EXCEL -ErrorAction SilentlyContinue); $attempt++) { Start-Sleep -Milliseconds 500 }
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Office process remained after data integrity form generation.' }

$frmBytes = [IO.File]::ReadAllBytes($frmOutput)
$frmText = [Text.Encoding]::GetEncoding(1251).GetString($frmBytes)
if ($frmText -notmatch ('Attribute VB_Name = "' + [regex]::Escape($TargetComponentName) + '"')) { throw 'Exported form has an unexpected VB_Name.' }
if ($frmText -notmatch 'Private Sub btnScan_Click\(\)') { throw 'Exported form is missing scan handler.' }
if ($frmText -notmatch 'Private Sub cboSeverity_Change\(\)') { throw 'Exported form is missing severity filter handler.' }
if ($frmText -match '(?<!\r)\n') { throw 'Exported form must use CRLF line endings.' }
Write-Output "DATA_INTEGRITY_DESIGNER_OK|form=$frmOutput|frx=$frxOutput|manifest=$manifestOutput"
