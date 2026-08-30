[CmdletBinding()]
param(
    [string]$WorkbookPath,
    [string]$OutputDirectory,
    [string]$TargetComponentName = 'frmPersonnelHistoryCenter'
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
    $control
}

function Add-TextBox {
    param([object]$Designer, [string]$Name, [double]$Left, [double]$Top, [double]$Width, [double]$Height, [bool]$Locked = $false, [bool]$MultiLine = $false)
    $control = Add-Control $Designer 'Forms.TextBox.1' $Name $Left $Top $Width $Height
    $control.Locked = $Locked
    $control.MultiLine = $MultiLine
    if ($MultiLine) { $control.ScrollBars = 3; $control.EnterKeyBehavior = $true }
    $control
}

if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Excel is running. Close Excel before generating the history center form.' }
$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedOutput = [IO.Path]::GetFullPath($OutputDirectory)
$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$frmOutput = Join-Path $resolvedOutput ($TargetComponentName + '.frm')
$frxOutput = Join-Path $resolvedOutput ($TargetComponentName + '.frx')
$manifestOutput = Join-Path $resolvedOutput ($TargetComponentName + '.layout.csv')
foreach ($path in @($frmOutput, $frxOutput, $manifestOutput)) { if (Test-Path -LiteralPath $path) { throw "Target artifact already exists: $path" } }

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$buildDirectory = Join-Path $projectRoot "Trash\personnel-history-center-designer-build-$stamp"
New-Item -ItemType Directory -Path $buildDirectory -Force | Out-Null
$buildWorkbook = Join-Path $buildDirectory 'CreateOrder.personnel-history-center-designer-build.xlsm'
Copy-Item -LiteralPath $resolvedWorkbook -Destination $buildWorkbook -Force

$formCode = @'
Option Explicit

Private mEmployeeID As String
Private mReport As String

Private Sub UserForm_Initialize()
    Me.Caption = t("history.center.title", "Personnel history and documents")
    lblDescription.Caption = t("history.center.description", "Search and review saved personnel history without changing registries.")
    lblEmployeeQuery.Caption = t("history.center.employee_query", "Employee search")
    lblEventID.Caption = t("history.center.event_id", "EventID")
    lblDocumentID.Caption = t("history.center.document_id", "DocumentID")
    lblReadOnly.Caption = t("history.center.read_only", "Read-only view: actions run only after an explicit click.")
    btnSearch.Caption = t("history.center.search", "Find")
    btnOpenDocument.Caption = t("history.center.open_document", "Open document")
    btnRepeatExport.Caption = t("history.center.repeat_export", "Repeat export")
    btnPrepareCorrection.Caption = t("history.center.prepare_correction", "Prepare correction")
    btnClose.Caption = t("history.center.close", "Close")
    lblSummary.Caption = t("history.center.not_loaded", "Not loaded.")
    txtTimeline.Text = vbNullString
    mEmployeeID = vbNullString
    mReport = vbNullString
End Sub

Private Sub btnSearch_Click()
    On Error GoTo Failed
    mEmployeeID = mdlPersonnelHistoryCenter.ResolvePersonnelHistoryEmployeeID(Trim$(txtEmployeeQuery.Text))
    mReport = mdlPersonnelHistoryCenter.BuildPersonnelHistoryCenterReport(Trim$(txtEmployeeQuery.Text), Trim$(txtEventID.Text), Trim$(txtDocumentID.Text))
    txtTimeline.Text = mReport
    lblSummary.Caption = Replace$(Split(mReport, vbCrLf)(UBound(Split(mReport, vbCrLf))), "SUMMARY | ", vbNullString)
    Exit Sub
Failed:
    mEmployeeID = vbNullString
    mReport = vbNullString
    txtTimeline.Text = vbNullString
    lblSummary.Caption = tf("history.center.action_failed", "History search failed: {error}", "{error}", Err.Description)
End Sub

Private Function EnsureEmployeeLoaded() As Boolean
    On Error GoTo Failed
    If Len(mEmployeeID) = 0 Then
        mEmployeeID = mdlPersonnelHistoryCenter.ResolvePersonnelHistoryEmployeeID(Trim$(txtEmployeeQuery.Text))
        mReport = mdlPersonnelHistoryCenter.BuildPersonnelHistoryCenterReport(Trim$(txtEmployeeQuery.Text), Trim$(txtEventID.Text), Trim$(txtDocumentID.Text))
        txtTimeline.Text = mReport
    End If
    EnsureEmployeeLoaded = (Len(mEmployeeID) > 0)
    Exit Function
Failed:
    lblSummary.Caption = tf("history.center.action_failed", "History action failed: {error}", "{error}", Err.Description)
End Function

Private Sub btnOpenDocument_Click()
    On Error GoTo Failed
    If Not EnsureEmployeeLoaded Then Exit Sub
    mdlPersonnelHistoryCenter.OpenPersonnelHistoryDocument mEmployeeID, Trim$(txtEventID.Text), Trim$(txtDocumentID.Text)
    lblSummary.Caption = t("history.center.action_done", "Document opened.")
    Exit Sub
Failed:
    lblSummary.Caption = tf("history.center.action_failed", "History action failed: {error}", "{error}", Err.Description)
End Sub

Private Sub btnRepeatExport_Click()
    On Error GoTo Failed
    If Not EnsureEmployeeLoaded Then Exit Sub
    lblSummary.Caption = mdlPersonnelHistoryCenter.RepeatPersonnelHistoryExport(mEmployeeID, Trim$(txtEventID.Text))
    Exit Sub
Failed:
    lblSummary.Caption = tf("history.center.action_failed", "History action failed: {error}", "{error}", Err.Description)
End Sub

Private Sub btnPrepareCorrection_Click()
    On Error GoTo Failed
    If Not EnsureEmployeeLoaded Then Exit Sub
    mdlPersonnelHistoryCenter.PreparePersonnelHistoryCorrectionFromCenter mEmployeeID, Trim$(txtEventID.Text)
    lblSummary.Caption = t("history.center.action_done", "Correction is prepared for review.")
    Exit Sub
Failed:
    lblSummary.Caption = tf("history.center.action_failed", "History action failed: {error}", "{error}", Err.Description)
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
    $component.Properties.Item('Caption').Value = 'Personnel history and documents'
    $component.Properties.Item('Width').Value = 680
    $component.Properties.Item('Height').Value = 500
    $component.Properties.Item('StartUpPosition').Value = 1

    $null = Add-Control $designer 'Forms.Label.1' 'lblTitle' 12 10 640 24 'Personnel history and documents'
    $null = Add-Control $designer 'Forms.Label.1' 'lblDescription' 12 38 640 30 'Search and review saved personnel history without changing registries.'
    $null = Add-Control $designer 'Forms.Label.1' 'lblEmployeeQuery' 12 78 100 18 'Employee search'
    $null = Add-TextBox $designer 'txtEmployeeQuery' 118 76 250 20
    $null = Add-Control $designer 'Forms.CommandButton.1' 'btnSearch' 380 74 90 24 'Find'
    $null = Add-Control $designer 'Forms.Label.1' 'lblEventID' 12 108 70 18 'EventID'
    $null = Add-TextBox $designer 'txtEventID' 86 106 170 20
    $null = Add-Control $designer 'Forms.Label.1' 'lblDocumentID' 274 108 80 18 'DocumentID'
    $null = Add-TextBox $designer 'txtDocumentID' 360 106 170 20
    $null = Add-Control $designer 'Forms.Label.1' 'lblSummary' 12 136 640 18 'Not loaded.'
    $null = Add-TextBox $designer 'txtTimeline' 12 162 640 250 $true $true
    $null = Add-Control $designer 'Forms.Label.1' 'lblReadOnly' 12 420 640 22 'Read-only view: actions run only after an explicit click.'
    $null = Add-Control $designer 'Forms.CommandButton.1' 'btnOpenDocument' 12 450 125 26 'Open document'
    $null = Add-Control $designer 'Forms.CommandButton.1' 'btnRepeatExport' 145 450 125 26 'Repeat export'
    $null = Add-Control $designer 'Forms.CommandButton.1' 'btnPrepareCorrection' 278 450 145 26 'Prepare correction'
    $null = Add-Control $designer 'Forms.CommandButton.1' 'btnClose' 570 450 82 26 'Close'

    $codeModule = $component.CodeModule
    [void]$codeModule.AddFromString($formCode)
    if (-not [bool]$designer.Controls.Item('txtTimeline').Locked) { throw 'Timeline textbox must be locked.' }
    if (-not [bool]$designer.Controls.Item('txtTimeline').MultiLine) { throw 'Timeline textbox must be multiline.' }
    New-Item -ItemType Directory -Path $resolvedOutput -Force | Out-Null
    [void]$component.Export($frmOutput)
    if (-not (Test-Path -LiteralPath $frxOutput)) { throw "Designer export did not create FRX: $frxOutput" }

    $controls = @(
        @{Name='lblTitle'; Type='Label'; Parent='root'; Left=12; Top=10; Width=640; Height=24; Locked=$false; Multiline=$false},
        @{Name='lblDescription'; Type='Label'; Parent='root'; Left=12; Top=38; Width=640; Height=30; Locked=$false; Multiline=$false},
        @{Name='lblEmployeeQuery'; Type='Label'; Parent='root'; Left=12; Top=78; Width=100; Height=18; Locked=$false; Multiline=$false},
        @{Name='txtEmployeeQuery'; Type='TextBox'; Parent='root'; Left=118; Top=76; Width=250; Height=20; Locked=$false; Multiline=$false},
        @{Name='btnSearch'; Type='CommandButton'; Parent='root'; Left=380; Top=74; Width=90; Height=24; Locked=$false; Multiline=$false},
        @{Name='lblEventID'; Type='Label'; Parent='root'; Left=12; Top=108; Width=70; Height=18; Locked=$false; Multiline=$false},
        @{Name='txtEventID'; Type='TextBox'; Parent='root'; Left=86; Top=106; Width=170; Height=20; Locked=$false; Multiline=$false},
        @{Name='lblDocumentID'; Type='Label'; Parent='root'; Left=274; Top=108; Width=80; Height=18; Locked=$false; Multiline=$false},
        @{Name='txtDocumentID'; Type='TextBox'; Parent='root'; Left=360; Top=106; Width=170; Height=20; Locked=$false; Multiline=$false},
        @{Name='lblSummary'; Type='Label'; Parent='root'; Left=12; Top=136; Width=640; Height=18; Locked=$false; Multiline=$false},
        @{Name='txtTimeline'; Type='TextBox'; Parent='root'; Left=12; Top=162; Width=640; Height=250; Locked=$true; Multiline=$true},
        @{Name='lblReadOnly'; Type='Label'; Parent='root'; Left=12; Top=420; Width=640; Height=22; Locked=$false; Multiline=$false},
        @{Name='btnOpenDocument'; Type='CommandButton'; Parent='root'; Left=12; Top=450; Width=125; Height=26; Locked=$false; Multiline=$false},
        @{Name='btnRepeatExport'; Type='CommandButton'; Parent='root'; Left=145; Top=450; Width=125; Height=26; Locked=$false; Multiline=$false},
        @{Name='btnPrepareCorrection'; Type='CommandButton'; Parent='root'; Left=278; Top=450; Width=145; Height=26; Locked=$false; Multiline=$false},
        @{Name='btnClose'; Type='CommandButton'; Parent='root'; Left=570; Top=450; Width=82; Height=26; Locked=$false; Multiline=$false}
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
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Office process remained after history center form generation.' }

$frmBytes = [IO.File]::ReadAllBytes($frmOutput)
$frmText = [Text.Encoding]::UTF8.GetString($frmBytes)
if ($frmText -notmatch ('Attribute VB_Name = "' + [regex]::Escape($TargetComponentName) + '"')) { throw 'Exported form has an unexpected VB_Name.' }
foreach ($handler in @('btnSearch_Click', 'btnOpenDocument_Click', 'btnRepeatExport_Click', 'btnPrepareCorrection_Click')) { if ($frmText -notmatch ('Private Sub ' + $handler + '\(\)')) { throw "Exported form is missing $handler." } }
if ($frmText -match '(?<!\r)\n') { throw 'Exported form must use CRLF line endings.' }
Write-Output "PERSONNEL_HISTORY_CENTER_DESIGNER_OK|form=$frmOutput|frx=$frxOutput|manifest=$manifestOutput"
