[CmdletBinding()]
param(
    [string]$WorkbookPath,
    [string]$OutputDirectory,
    [string]$TargetComponentName = 'frmEnrollmentWizardV2'
)

if ([string]::IsNullOrWhiteSpace($WorkbookPath)) { $WorkbookPath = Join-Path $PSScriptRoot '..\CreateOrder.xlsm' }
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) { $OutputDirectory = Join-Path $PSScriptRoot '..\CreateOrder.xlsm.modules' }

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-LayoutLog {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('DEBUG', 'INFO', 'WARN', 'ERROR')][string]$Level,
        [Parameter(Mandatory = $true)][string]$Message,
        [hashtable]$Context = @{}
    )

    $payload = [ordered]@{
        timestamp = (Get-Date).ToString('o')
        level = $Level
        operation = 'New-EnrollmentWizardV2Designer'
        message = $Message
    }
    foreach ($key in $Context.Keys) { $payload[$key] = $Context[$key] }

    $line = $payload | ConvertTo-Json -Compress -Depth 5
    if ($Level -eq 'DEBUG') {
        Write-Verbose $line
    } elseif ($Level -eq 'WARN') {
        Write-Warning $line
    } elseif ($Level -eq 'ERROR') {
        Write-Error $line
    } else {
        Write-Host $line
    }
}

function Release-ComObject {
    param([object]$Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

function Test-VbComponentExists {
    param(
        [Parameter(Mandatory = $true)][object]$VbProject,
        [Parameter(Mandatory = $true)][string]$Name
    )

    $components = $null
    try {
        $components = $VbProject.VBComponents
        for ($index = 1; $index -le $components.Count; $index++) {
            $component = $null
            try {
                $component = $components.Item($index)
                if ($component.Name -eq $Name) { return $true }
            } finally {
                Release-ComObject $component
            }
        }
    } finally {
        Release-ComObject $components
    }
    return $false
}

if (-not ('EnrollmentDesignerNativeMethods' -as [type])) {
    Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class EnrollmentDesignerNativeMethods {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@
}

function Get-ExcelProcessId {
    param([Parameter(Mandatory = $true)][object]$ExcelApplication)
    [uint32]$processId = 0
    [void][EnrollmentDesignerNativeMethods]::GetWindowThreadProcessId([IntPtr]$ExcelApplication.Hwnd, [ref]$processId)
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
        Write-LayoutLog WARN 'Excel did not exit after COM Quit; stopping only the process created by this generator.' @{ processId = $ProcessId }
        Stop-Process -Id $ProcessId -Force
    }
}

$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedOutput = [IO.Path]::GetFullPath($OutputDirectory)
$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))

if (-not $resolvedWorkbook.StartsWith($projectRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Workbook must be inside the CreateOrder project: $resolvedWorkbook"
}
if (-not $resolvedOutput.StartsWith($projectRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Output directory must be inside the CreateOrder project: $resolvedOutput"
}
if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    throw 'Excel is running. Close Excel before generating or importing the designer form.'
}

$frmOutput = Join-Path $resolvedOutput ($TargetComponentName + '.frm')
$frxOutput = Join-Path $resolvedOutput ($TargetComponentName + '.frx')
$manifestOutput = Join-Path $resolvedOutput ($TargetComponentName + '.layout.csv')
foreach ($path in @($frmOutput, $frxOutput, $manifestOutput)) {
    if (Test-Path -LiteralPath $path) {
        throw "Target artifact already exists; refusing to overwrite owner layout: $path"
    }
}

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$buildDirectory = Join-Path $projectRoot ("Trash\enrollment-designer-v2-build-$stamp")
$backupDirectory = Join-Path $projectRoot ("CreateOrderBackups\enrollment-designer-v2-$stamp")
New-Item -ItemType Directory -Path $buildDirectory -Force | Out-Null
New-Item -ItemType Directory -Path $backupDirectory -Force | Out-Null
New-Item -ItemType Directory -Path $resolvedOutput -Force | Out-Null

$buildWorkbook = Join-Path $buildDirectory 'CreateOrder.designer-build.xlsm'
$backupWorkbook = Join-Path $backupDirectory 'CreateOrder.before-enrollment-designer-v2.xlsm'
Copy-Item -LiteralPath $resolvedWorkbook -Destination $buildWorkbook
Copy-Item -LiteralPath $resolvedWorkbook -Destination $backupWorkbook

Write-LayoutLog INFO 'Prepared isolated build workbook and safety backup.' @{
    buildWorkbook = $buildWorkbook
    backupWorkbook = $backupWorkbook
}

$builderCode = @'
Option Explicit

Private Const SOURCE_FORM As String = "frmEnrollmentWizard"
Private Const TARGET_FORM As String = "frmEnrollmentWizardV2"
Private usedNames As Object
Private manifestSheet As Worksheet
Private manifestRow As Long
Private lastStep As String

Public Sub BuildDesignerV2()
    On Error GoTo ErrorHandler

    Dim sourceForm As Object
    Dim targetComponent As Object
    Dim targetDesigner As Object

    lastStep = "Create name registry"
    Set usedNames = CreateObject("Scripting.Dictionary")
    lastStep = "Prepare manifest sheet"
    Set manifestSheet = PrepareManifestSheet()
    manifestRow = 2

    lastStep = "Load source form"
    Load frmEnrollmentWizard
    Set sourceForm = frmEnrollmentWizard

    lastStep = "Check target component"
    If ComponentExists(TARGET_FORM) Then Err.Raise 5, , "Target form already exists: " & TARGET_FORM

    lastStep = "Create target component"
    Set targetComponent = ThisWorkbook.VBProject.VBComponents.Add(3)
    targetComponent.Name = TARGET_FORM
    Set targetDesigner = targetComponent.Designer

    lastStep = "Set target form properties"
    targetComponent.Properties.Item("Caption").Value = sourceForm.Caption & " V2 - layout"
    targetComponent.Properties.Item("Width").Value = sourceForm.Width
    targetComponent.Properties.Item("Height").Value = sourceForm.Height
    targetComponent.Properties.Item("StartUpPosition").Value = 1

    lastStep = "Add layout-only form code"
    targetComponent.CodeModule.AddFromString _
        "Option Explicit" & vbCrLf & vbCrLf & _
        "' Layout-only design form. Business logic is intentionally connected after owner layout acceptance." & vbCrLf & _
        "Private Sub UserForm_Initialize()" & vbCrLf & _
        "    ' Intentionally empty: never override owner-defined geometry at runtime." & vbCrLf & _
        "End Sub" & vbCrLf

    lastStep = "Clone root controls"
    CloneContainer sourceForm, targetDesigner, "root", True
    lastStep = "Finalize manifest"
    Unload frmEnrollmentWizard
    manifestSheet.Columns.AutoFit
    Exit Sub

ErrorHandler:
    Dim capturedNumber As Long
    Dim capturedDescription As String
    capturedNumber = Err.Number
    capturedDescription = Err.Description
    On Error Resume Next
    If Not manifestSheet Is Nothing Then
        manifestSheet.Range("M1").Value = capturedNumber
        manifestSheet.Range("N1").Value = capturedDescription
        manifestSheet.Range("O1").Value = lastStep
    End If
    Unload frmEnrollmentWizard
    On Error GoTo 0
    Exit Sub
End Sub

Private Function PrepareManifestSheet() As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("__EnrollmentV2Layout").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set PrepareManifestSheet = ThisWorkbook.Worksheets.Add
    PrepareManifestSheet.Name = "__EnrollmentV2Layout"
    PrepareManifestSheet.Visible = xlSheetVeryHidden
    PrepareManifestSheet.Range("A1:K1").Value = Array( _
        "container_path", "source_name", "designer_name", "control_type", "caption", _
        "left", "top", "width", "height", "visible", "enabled")
End Function

Private Sub CloneContainer(ByVal sourceContainer As Object, ByVal targetContainer As Object, ByVal containerPath As String, ByVal isRoot As Boolean)
    Dim sourceControl As Object
    Dim sourceParent As Object

    For Each sourceControl In sourceContainer.Controls
        Set sourceParent = Nothing
        On Error Resume Next
        Set sourceParent = sourceControl.Parent
        On Error GoTo 0
        If sourceParent Is sourceContainer Then
            If Not (isRoot And ShouldSkipRootControl(CStr(sourceControl.Name))) Then
                CloneControl sourceControl, targetContainer, containerPath
            End If
        End If
    Next sourceControl
End Sub

Private Sub CloneControl(ByVal sourceControl As Object, ByVal targetContainer As Object, ByVal containerPath As String)
    Dim kind As String
    Dim targetControl As Object
    Dim targetName As String
    Dim childPath As String

    kind = TypeName(sourceControl)
    targetName = UniqueControlName(CStr(sourceControl.Name), ContainerPrefix(containerPath))
    lastStep = "Clone " & containerPath & "/" & sourceControl.Name & " as " & kind

    Select Case kind
        Case "MultiPage"
            Set targetControl = targetContainer.Controls.Add("Forms.MultiPage.1", targetName, True)
            CopyCommonProperties sourceControl, targetControl
            CloneMultiPage sourceControl, targetControl, containerPath & "/" & targetName
        Case "Frame"
            Set targetControl = targetContainer.Controls.Add("Forms.Frame.1", targetName, True)
            CopyCommonProperties sourceControl, targetControl
            childPath = containerPath & "/" & targetName
            WriteManifest containerPath, sourceControl, targetControl
            CloneContainer sourceControl, targetControl, childPath, False
            Exit Sub
        Case "Label"
            Set targetControl = targetContainer.Controls.Add("Forms.Label.1", targetName, True)
            CopyCommonProperties sourceControl, targetControl
        Case "TextBox"
            Set targetControl = targetContainer.Controls.Add("Forms.TextBox.1", targetName, True)
            CopyCommonProperties sourceControl, targetControl
        Case "ComboBox"
            Set targetControl = targetContainer.Controls.Add("Forms.ComboBox.1", targetName, True)
            CopyCommonProperties sourceControl, targetControl
        Case "CheckBox"
            Set targetControl = targetContainer.Controls.Add("Forms.CheckBox.1", targetName, True)
            CopyCommonProperties sourceControl, targetControl
        Case "CommandButton"
            Set targetControl = targetContainer.Controls.Add("Forms.CommandButton.1", targetName, True)
            CopyCommonProperties sourceControl, targetControl
        Case "ListBox"
            Set targetControl = targetContainer.Controls.Add("Forms.ListBox.1", targetName, True)
            CopyCommonProperties sourceControl, targetControl
        Case Else
            Err.Raise 5, , "Unsupported control type: " & kind & " (" & sourceControl.Name & ")"
    End Select

    WriteManifest containerPath, sourceControl, targetControl
End Sub

Private Sub CloneMultiPage(ByVal sourceMultiPage As Object, ByVal targetMultiPage As Object, ByVal containerPath As String)
    Dim pageIndex As Long
    Dim sourcePage As Object
    Dim targetPage As Object
    Dim pageName As String
    Dim pagePath As String

    Do While targetMultiPage.Pages.Count > sourceMultiPage.Pages.Count
        targetMultiPage.Pages.Remove targetMultiPage.Pages.Count - 1
    Loop
    Do While targetMultiPage.Pages.Count < sourceMultiPage.Pages.Count
        targetMultiPage.Pages.Add
    Loop

    For pageIndex = 0 To sourceMultiPage.Pages.Count - 1
        Set sourcePage = sourceMultiPage.Pages(pageIndex)
        Set targetPage = targetMultiPage.Pages(pageIndex)
        pageName = PageDesignerName(pageIndex)
        On Error Resume Next
        targetPage.Name = pageName
        targetPage.Caption = sourcePage.Caption
        targetPage.Visible = sourcePage.Visible
        targetPage.Enabled = sourcePage.Enabled
        targetPage.ScrollBars = sourcePage.ScrollBars
        targetPage.ScrollHeight = sourcePage.ScrollHeight
        targetPage.ScrollWidth = sourcePage.ScrollWidth
        targetPage.KeepScrollBarsVisible = sourcePage.KeepScrollBarsVisible
        On Error GoTo 0
        pagePath = containerPath & "/" & pageName
        WriteManifestPage containerPath, sourcePage, targetPage
        CloneContainer sourcePage, targetPage, pagePath, False
    Next pageIndex

    On Error Resume Next
    targetMultiPage.Value = sourceMultiPage.Value
    On Error GoTo 0
End Sub

Private Sub CopyCommonProperties(ByVal sourceControl As Object, ByVal targetControl As Object)
    On Error Resume Next
    targetControl.Left = sourceControl.Left
    targetControl.Top = sourceControl.Top
    targetControl.Width = sourceControl.Width
    targetControl.Height = sourceControl.Height
    targetControl.Caption = sourceControl.Caption
    targetControl.Visible = sourceControl.Visible
    targetControl.Enabled = sourceControl.Enabled
    targetControl.TabStop = sourceControl.TabStop
    targetControl.TabIndex = sourceControl.TabIndex
    targetControl.ControlTipText = sourceControl.ControlTipText
    targetControl.Tag = sourceControl.Tag
    targetControl.WordWrap = sourceControl.WordWrap
    targetControl.MultiLine = sourceControl.MultiLine
    targetControl.Locked = sourceControl.Locked
    targetControl.TextAlign = sourceControl.TextAlign
    targetControl.Style = sourceControl.Style
    targetControl.MatchEntry = sourceControl.MatchEntry
    targetControl.ListStyle = sourceControl.ListStyle
    targetControl.ScrollBars = sourceControl.ScrollBars
    targetControl.ScrollHeight = sourceControl.ScrollHeight
    targetControl.ScrollWidth = sourceControl.ScrollWidth
    targetControl.KeepScrollBarsVisible = sourceControl.KeepScrollBarsVisible
    targetControl.ColumnCount = sourceControl.ColumnCount
    targetControl.ColumnWidths = sourceControl.ColumnWidths
    targetControl.IntegralHeight = sourceControl.IntegralHeight
    targetControl.BoundColumn = sourceControl.BoundColumn
    targetControl.ListRows = sourceControl.ListRows
    targetControl.SpecialEffect = sourceControl.SpecialEffect
    targetControl.BackStyle = sourceControl.BackStyle
    targetControl.BackColor = sourceControl.BackColor
    targetControl.ForeColor = sourceControl.ForeColor
    targetControl.TakeFocusOnClick = sourceControl.TakeFocusOnClick
    targetControl.Font.Name = sourceControl.Font.Name
    targetControl.Font.Size = sourceControl.Font.Size
    targetControl.Font.Bold = sourceControl.Font.Bold
    targetControl.Font.Italic = sourceControl.Font.Italic
    On Error GoTo 0
End Sub

Private Function ShouldSkipRootControl(ByVal controlName As String) As Boolean
    Select Case controlName
        Case "Frame1", "Frame2", "Frame3", "lstPeriods", "txtPeriodStart", "txtPeriodEnd", _
             "cmbReason", "Label_PeriodStart", "Label_PeriodEnd", "Label_Reason", _
             "lblFIO", "lblZvanie", "lblDolzhnost", "lblChast", _
             "btnAddPeriod", "btnEditPeriod", "btnDeletePeriod", "btnSaveGenerateDynamic"
            ShouldSkipRootControl = True
    End Select
End Function

Private Function UniqueControlName(ByVal sourceName As String, ByVal prefix As String) As String
    Dim candidate As String
    Dim cleanedSourceName As String
    Dim suffix As Long

    cleanedSourceName = CleanControlName(sourceName)
    candidate = Left$(cleanedSourceName, 40)
    If Not usedNames.Exists(LCase$(candidate)) Then
        usedNames.Add LCase$(candidate), True
        UniqueControlName = candidate
        Exit Function
    End If

    candidate = Left$(prefix & "_" & cleanedSourceName, 40)
    suffix = 2
    Do While usedNames.Exists(LCase$(candidate))
        candidate = Left$(prefix & "_" & cleanedSourceName, 36) & "_" & CStr(suffix)
        suffix = suffix + 1
    Loop
    usedNames.Add LCase$(candidate), True
    UniqueControlName = candidate
End Function

Private Function CleanControlName(ByVal sourceName As String) As String
    Dim charIndex As Long
    Dim currentChar As String
    Dim charCode As Long
    Dim result As String

    For charIndex = 1 To Len(sourceName)
        currentChar = Mid$(sourceName, charIndex, 1)
        charCode = AscW(currentChar)
        If (charCode >= 48 And charCode <= 57) Or _
           (charCode >= 65 And charCode <= 90) Or _
           (charCode >= 97 And charCode <= 122) Or currentChar = "_" Then
            result = result & currentChar
        Else
            result = result & "_"
        End If
    Next charIndex

    If Len(result) = 0 Then result = "control"
    If Mid$(result, 1, 1) Like "#" Then result = "ctl_" & result
    CleanControlName = result
End Function

Private Function ContainerPrefix(ByVal containerPath As String) As String
    If InStr(1, containerPath, "pgEmployee", vbTextCompare) > 0 Then
        ContainerPrefix = "emp"
    ElseIf InStr(1, containerPath, "pgDocs", vbTextCompare) > 0 Then
        ContainerPrefix = "doc"
    ElseIf InStr(1, containerPath, "pgMonthly", vbTextCompare) > 0 Then
        ContainerPrefix = "mon"
    ElseIf InStr(1, containerPath, "pgOneTime", vbTextCompare) > 0 Then
        ContainerPrefix = "one"
    ElseIf InStr(1, containerPath, "pgAdvanced", vbTextCompare) > 0 Then
        ContainerPrefix = "adv"
    ElseIf InStr(1, containerPath, "pgExtras", vbTextCompare) > 0 Then
        ContainerPrefix = "ext"
    ElseIf InStr(1, containerPath, "pgPreview", vbTextCompare) > 0 Then
        ContainerPrefix = "pre"
    Else
        ContainerPrefix = "root"
    End If
End Function

Private Function PageDesignerName(ByVal pageIndex As Long) As String
    Select Case pageIndex
        Case 0: PageDesignerName = "pgEmployee"
        Case 1: PageDesignerName = "pgDocs"
        Case 2: PageDesignerName = "pgMonthly"
        Case 3: PageDesignerName = "pgOneTime"
        Case 4: PageDesignerName = "pgAdvanced"
        Case 5: PageDesignerName = "pgExtras"
        Case 6: PageDesignerName = "pgPreview"
        Case Else: PageDesignerName = "pgPage" & CStr(pageIndex + 1)
    End Select
End Function

Private Sub WriteManifest(ByVal containerPath As String, ByVal sourceControl As Object, ByVal targetControl As Object)
    manifestSheet.Cells(manifestRow, 1).Value = containerPath
    manifestSheet.Cells(manifestRow, 2).Value = sourceControl.Name
    manifestSheet.Cells(manifestRow, 3).Value = targetControl.Name
    manifestSheet.Cells(manifestRow, 4).Value = TypeName(sourceControl)
    On Error Resume Next
    manifestSheet.Cells(manifestRow, 5).Value = sourceControl.Caption
    On Error GoTo 0
    manifestSheet.Cells(manifestRow, 6).Value = sourceControl.Left
    manifestSheet.Cells(manifestRow, 7).Value = sourceControl.Top
    manifestSheet.Cells(manifestRow, 8).Value = sourceControl.Width
    manifestSheet.Cells(manifestRow, 9).Value = sourceControl.Height
    manifestSheet.Cells(manifestRow, 10).Value = sourceControl.Visible
    manifestSheet.Cells(manifestRow, 11).Value = sourceControl.Enabled
    manifestRow = manifestRow + 1
End Sub

Private Sub WriteManifestPage(ByVal containerPath As String, ByVal sourcePage As Object, ByVal targetPage As Object)
    manifestSheet.Cells(manifestRow, 1).Value = containerPath
    manifestSheet.Cells(manifestRow, 2).Value = sourcePage.Name
    manifestSheet.Cells(manifestRow, 3).Value = targetPage.Name
    manifestSheet.Cells(manifestRow, 4).Value = "Page"
    manifestSheet.Cells(manifestRow, 5).Value = sourcePage.Caption
    manifestSheet.Cells(manifestRow, 10).Value = sourcePage.Visible
    manifestSheet.Cells(manifestRow, 11).Value = sourcePage.Enabled
    manifestRow = manifestRow + 1
End Sub

Private Function ComponentExists(ByVal componentName As String) As Boolean
    Dim component As Object
    For Each component In ThisWorkbook.VBProject.VBComponents
        If StrComp(component.Name, componentName, vbTextCompare) = 0 Then
            ComponentExists = True
            Exit Function
        End If
    Next component
End Function
'@

$excel = $null
$excelProcessId = 0
$buildBook = $null
$builderComponent = $null
$targetComponent = $null
$manifestSheet = $null
$usedRange = $null
try {
    Write-LayoutLog INFO 'Opening isolated workbook and generating design-time form.' @{ workbook = $buildWorkbook }
    $excel = New-Object -ComObject Excel.Application
    $excelProcessId = Get-ExcelProcessId -ExcelApplication $excel
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $buildBook = $excel.Workbooks.Open($buildWorkbook, 0, $false)

    if ($buildBook.VBProject.Protection -ne 0) { throw 'VBA project is protected; cannot create the designer form.' }
    if (Test-VbComponentExists -VbProject $buildBook.VBProject -Name $TargetComponentName) {
        throw "Target component already exists in the build workbook: $TargetComponentName"
    }

    $builderComponent = $buildBook.VBProject.VBComponents.Add(1)
    $builderComponent.Name = 'modEnrollmentDesignerBuilder'
    $builderComponent.CodeModule.AddFromString($builderCode)
    $excel.Run("'$($buildBook.Name)'!modEnrollmentDesignerBuilder.BuildDesignerV2")

    $manifestSheet = $buildBook.Worksheets.Item('__EnrollmentV2Layout')
    $builderErrorNumber = [int]$manifestSheet.Range('M1').Value2
    if ($builderErrorNumber -ne 0) {
        $builderErrorDescription = [string]$manifestSheet.Range('N1').Value2
        $builderErrorStep = [string]$manifestSheet.Range('O1').Value2
        throw "VBA builder failed at '$builderErrorStep': [$builderErrorNumber] $builderErrorDescription"
    }

    $targetComponent = $buildBook.VBProject.VBComponents.Item($TargetComponentName)
    if ($targetComponent.Type -ne 3) { throw "Generated component has unexpected type: $($targetComponent.Type)" }

    $targetComponent.Export($frmOutput)
    if (-not (Test-Path -LiteralPath $frmOutput) -or -not (Test-Path -LiteralPath $frxOutput)) {
        throw 'Excel did not export the expected .frm/.frx pair.'
    }

    $usedRange = $manifestSheet.UsedRange
    $values = $usedRange.Value2
    $rows = @()
    for ($row = 2; $row -le $usedRange.Rows.Count; $row++) {
        $rows += [pscustomobject]@{
            container_path = [string]$values[$row, 1]
            source_name = [string]$values[$row, 2]
            designer_name = [string]$values[$row, 3]
            control_type = [string]$values[$row, 4]
            caption = [string]$values[$row, 5]
            left = $values[$row, 6]
            top = $values[$row, 7]
            width = $values[$row, 8]
            height = $values[$row, 9]
            visible = $values[$row, 10]
            enabled = $values[$row, 11]
        }
    }
    $rows | Export-Csv -LiteralPath $manifestOutput -NoTypeInformation -Encoding utf8
    $buildBook.Save()

    Write-LayoutLog INFO 'Generated and exported design-time form from isolated workbook.' @{
        controlsAndPages = $rows.Count
        frm = $frmOutput
        frx = $frxOutput
        manifest = $manifestOutput
    }
} catch {
    Write-LayoutLog ERROR 'Designer form generation failed.' @{ error = $_.Exception.Message }
    throw
} finally {
    if ($buildBook) { $buildBook.Close($false) }
    if ($excel) { $excel.Quit() }
    Release-ComObject $usedRange
    Release-ComObject $manifestSheet
    Release-ComObject $targetComponent
    Release-ComObject $builderComponent
    Release-ComObject $buildBook
    Release-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Stop-OwnedExcelProcessIfNeeded -ProcessId $excelProcessId
}

for ($attempt = 1; $attempt -le 20 -and (Get-Process EXCEL -ErrorAction SilentlyContinue); $attempt++) {
    Start-Sleep -Milliseconds 250
}
if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    throw 'Excel remained running after isolated generation; refusing to import into the working workbook.'
}

$excel = $null
$excelProcessId = 0
$workingBook = $null
$importedComponent = $null
try {
    Write-LayoutLog INFO 'Importing verified designer form into working workbook.' @{ workbook = $resolvedWorkbook }
    $excel = New-Object -ComObject Excel.Application
    $excelProcessId = Get-ExcelProcessId -ExcelApplication $excel
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $workingBook = $excel.Workbooks.Open($resolvedWorkbook, 0, $false)

    if ($workingBook.VBProject.Protection -ne 0) { throw 'Working VBA project is protected; cannot import the designer form.' }
    if (Test-VbComponentExists -VbProject $workingBook.VBProject -Name $TargetComponentName) {
        throw "Target component already exists in the working workbook: $TargetComponentName"
    }

    $importedComponent = $workingBook.VBProject.VBComponents.Import($frmOutput)
    if ($importedComponent.Name -ne $TargetComponentName -or $importedComponent.Type -ne 3) {
        throw "Imported component verification failed: name=$($importedComponent.Name), type=$($importedComponent.Type)"
    }
    if (-not (Test-VbComponentExists -VbProject $workingBook.VBProject -Name 'frmEnrollmentWizard')) {
        throw 'Original frmEnrollmentWizard is missing after V2 import.'
    }

    $workingBook.Save()
    Write-LayoutLog INFO 'Designer V2 installed; original form remains present.' @{
        workbook = $resolvedWorkbook
        component = $TargetComponentName
        backup = $backupWorkbook
    }
} catch {
    Write-LayoutLog ERROR 'Designer form import failed.' @{ error = $_.Exception.Message; backup = $backupWorkbook }
    throw
} finally {
    if ($workingBook) { $workingBook.Close($false) }
    if ($excel) { $excel.Quit() }
    Release-ComObject $importedComponent
    Release-ComObject $workingBook
    Release-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Stop-OwnedExcelProcessIfNeeded -ProcessId $excelProcessId
}

for ($attempt = 1; $attempt -le 20 -and (Get-Process EXCEL -ErrorAction SilentlyContinue); $attempt++) {
    Start-Sleep -Milliseconds 250
}
if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    throw 'Excel remained running after import.'
}

Write-LayoutLog INFO 'Enrollment designer V2 generation completed.' @{
    component = $TargetComponentName
    backup = $backupWorkbook
    buildWorkbook = $buildWorkbook
}
