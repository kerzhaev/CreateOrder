VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchFIO 
   Caption         =   "UserForm1"
   ClientHeight    =   8925.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15240
   OleObjectBlob   =   "frmSearchFIO.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmSearchFIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' =====================================================================
' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ пїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ
' пїЅпїЅпїЅпїЅпїЅ: пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ, пїЅпїЅпїЅ "95 пїЅпїЅпїЅ" пїЅпїЅ пїЅпїЅ
' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ: v.1.5.1 (Multi-column Search & Keyboard Navigation)
' =====================================================================

Option Explicit

' === пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ ===
Public selectedLichniyNomer As String     ' пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅ
Private Const DSO_COL_FIO As Long = 2
Private Const DSO_COL_LN As Long = 3
Private Const LONG_PERIOD_WARNING_DAYS As Long = 120

' =====================================================================
' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ
' =====================================================================
Private Sub UserForm_Initialize()
    Me.Caption = t("form.search_fio.title", "Search and SVO periods")
    btnSelect.Caption = t("form.search_fio.button.select", "Select")
    Frame1.Caption = t("form.search_fio.frame.person", "Service member information")
    Frame2.Caption = t("form.search_fio.frame.current_periods", "Current periods")
    Frame3.Caption = t("form.search_fio.frame.add_period", "Add period")
    Label_PeriodStart.Caption = t("form.search_fio.label.period_start", "Start")
    Label_PeriodEnd.Caption = t("form.search_fio.label.period_end", "End")
    Label_Reason.Caption = t("form.search_fio.label.reason", "Reason")
    btnAddPeriod.Caption = t("form.search_fio.button.add_period", "Add period")
    btnEditPeriod.Caption = t("form.search_fio.button.edit_period", "Edit period")
    btnDeletePeriod.Caption = t("form.search_fio.button.delete_period", "Delete period")
    btnClose.Caption = t("form.search_fio.button.close", "Close")

    ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ (5 пїЅпїЅпїЅпїЅпїЅпїЅпїЅ)
    ' 1:пїЅ.пїЅпїЅпїЅпїЅпїЅ, 2:пїЅпїЅпїЅ, 3:пїЅпїЅпїЅпїЅпїЅпїЅ, 4:пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ, 5:пїЅпїЅпїЅпїЅпїЅ
    With lstResults
        .ColumnCount = 5
        .ColumnWidths = "35 pt;120 pt;70 pt;100 pt;30 pt"
    End With

    ' пїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅ пїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ (пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ, пїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ)
    If selectedLichniyNomer <> "" Then
        Call ShowPassportData(selectedLichniyNomer)
        Call LoadPeriodsForLichniy(selectedLichniyNomer)
        ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ, пїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ
        mdlHelper.EnsureStaffColumnsInitialized
    End If
End Sub

Public Sub FillByLichniyNomer()
    If selectedLichniyNomer <> "" Then
        Call ShowPassportData(selectedLichniyNomer)
        Call LoadPeriodsForLichniy(selectedLichniyNomer)
    End If
End Sub

' =====================================================================
' пїЅпїЅпїЅпїЅпїЅ (LIVE SEARCH) пїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ
' =====================================================================

'/**
'* пїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ пїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ 5 пїЅпїЅпїЅпїЅпїЅпїЅпїЅ ListBox.
'*/
Private Sub txtSearch_Change()
    mdlHelper.EnsureStaffColumnsInitialized
    
    Dim wsStaff As Worksheet
    Dim lastRow As Long, i As Long, foundCount As Long
    Dim query As String
    Dim fLichniy As String, fFIO As String, fZvanie As String, fDolzhnost As String, fChast As String
    
    Set wsStaff = mdlHelper.GetStaffWorksheet()
    If wsStaff Is Nothing Then Exit Sub
    
    ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ (пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅ mdlHelper)
    If mdlHelper.colFIO_Global <= 0 Then Exit Sub

    lastRow = wsStaff.Cells(wsStaff.Rows.count, mdlHelper.colFIO_Global).End(xlUp).Row
    
    query = LCase(Trim(txtSearch.Text))
    lstResults.Clear
    
    If Len(query) < 2 Then
        lblStatus.Caption = t("common.status_enter_min_chars", "Enter at least 2 characters...")
        Exit Sub
    End If

    foundCount = 0
    Application.ScreenUpdating = False ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ
    
    For i = 2 To lastRow
        fLichniy = CStr(wsStaff.Cells(i, mdlHelper.colLichniyNomer_Global).value)
        fFIO = CStr(wsStaff.Cells(i, mdlHelper.colFIO_Global).value)
        
        ' пїЅпїЅпїЅпїЅпїЅ пїЅпїЅ пїЅпїЅпїЅ пїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ
        If InStr(LCase(fFIO), query) > 0 Or InStr(LCase(fLichniy), query) > 0 Then
            fZvanie = CStr(wsStaff.Cells(i, mdlHelper.colZvanie_Global).value)
            fDolzhnost = CStr(wsStaff.Cells(i, mdlHelper.colDolzhnost_Global).value)
            fChast = CStr(wsStaff.Cells(i, mdlHelper.colVoinskayaChast_Global).value)
            
            ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ
            lstResults.AddItem fLichniy             ' Col 0 (Hidden ID/Key)
            lstResults.List(foundCount, 1) = fFIO        ' Col 1
            lstResults.List(foundCount, 2) = fZvanie     ' Col 2
            lstResults.List(foundCount, 3) = fDolzhnost  ' Col 3
            lstResults.List(foundCount, 4) = fChast      ' Col 4
            
            foundCount = foundCount + 1
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    If foundCount = 0 Then
        lblStatus.Caption = t("common.status_none", "No matches found.")
    Else
        lblStatus.Caption = tf("common.status_found", "Found: {count}", "{count}", foundCount)
    End If
End Sub

' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ пїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ
Private Sub lstResults_Click()
    If lstResults.ListCount = 0 Or lstResults.ListIndex = -1 Then Exit Sub
    
    Dim lichniyNomer As String
    Dim wsDSO As Worksheet, i As Long, lastRowDSO As Long
    
    ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ пїЅпїЅ 0-пїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ
    lichniyNomer = lstResults.List(lstResults.ListIndex, 0)

    Set wsDSO = mdlHelper.GetDsoWorksheet()
    If wsDSO Is Nothing Then Exit Sub
    lastRowDSO = wsDSO.Cells(wsDSO.Rows.count, DSO_COL_LN).End(xlUp).Row
    
    ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅ пїЅпїЅпїЅ пїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅ пїЅпїЅпїЅпїЅ
    For i = 2 To lastRowDSO
        If Trim(wsDSO.Cells(i, DSO_COL_LN).value) = Trim(lichniyNomer) Then
            NavigateToDSORow wsDSO, i
            Exit For
        End If
    Next i

    selectedLichniyNomer = lichniyNomer
    ShowPassportData selectedLichniyNomer
    LoadPeriodsForLichniy selectedLichniyNomer
    lblStatus.Caption = tf("form.search_fio.status.selected", "Selected: {fio}", _
                           "{fio}", lstResults.List(lstResults.ListIndex, 1))
End Sub

' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ: пїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ
Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅ - пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅ пїЅпїЅпїЅпїЅпїЅпїЅ
    If KeyCode = vbKeyDown Then
        If lstResults.ListCount > 0 Then
            lstResults.SetFocus
            If lstResults.ListIndex = -1 Then lstResults.ListIndex = 0
        End If
        KeyCode = 0
    End If
    
    ' ENTER - пїЅпїЅпїЅпїЅ 1 пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ, пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ
    If KeyCode = vbKeyReturn Then
        If lstResults.ListCount = 1 Then
            lstResults.ListIndex = 0
            Call lstResults_Click
            txtPeriodStart.SetFocus ' пїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅ пїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅ
        ElseIf lstResults.ListCount > 1 Then
            lstResults.SetFocus
            lstResults.ListIndex = 0
        End If
        KeyCode = 0
    End If
End Sub

' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ: пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ
Private Sub lstResults_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' ENTER - пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ
    If KeyCode = vbKeyReturn Then
        If lstResults.ListIndex > -1 Then
            Call lstResults_Click
            txtPeriodStart.SetFocus
        End If
        KeyCode = 0
    End If
End Sub

' пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅ - пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅ пїЅпїЅпїЅ (пїЅпїЅпїЅпїЅ пїЅпїЅпїЅ) пїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ
Private Sub lstResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call AddSearchResultToDSO_LastRow
End Sub

Private Sub btnSelect_Click()
    Call AddSearchResultToDSO_LastRow
End Sub

' =====================================================================
' пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅ пїЅпїЅпїЅ (Add New Logic)
' =====================================================================
Private Sub AddSearchResultToDSO_LastRow()
    If lstResults.ListCount = 0 Or lstResults.ListIndex = -1 Then Exit Sub

    Dim wsDSO As Worksheet
    Set wsDSO = mdlHelper.GetDsoWorksheet()
    If wsDSO Is Nothing Then Exit Sub

    Dim fioVal As String, lnVal As String
    ' пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ ListBox
    lnVal = lstResults.List(lstResults.ListIndex, 0)
    fioVal = lstResults.List(lstResults.ListIndex, 1)

    ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ
    Dim exists As Boolean, lastRowDSO As Long, i As Long
    exists = False
    lastRowDSO = wsDSO.Cells(wsDSO.Rows.count, DSO_COL_LN).End(xlUp).Row
    
    For i = 2 To lastRowDSO
        If Trim(wsDSO.Cells(i, DSO_COL_LN).value) = Trim(lnVal) Then
            exists = True
            Exit For
        End If
    Next i

    If exists Then
        selectedLichniyNomer = lnVal
        NavigateToDSORow wsDSO, i
        ShowPassportData selectedLichniyNomer
        LoadPeriodsForLichniy selectedLichniyNomer
        lblStatus.Caption = tf("form.search_fio.message.employee_exists", _
                               "Employee already exists in DSO (row {row}).", _
                               "{row}", i)
        txtPeriodStart.SetFocus
        Exit Sub
    End If

    ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ
    lastRowDSO = lastRowDSO + 1
    wsDSO.Cells(lastRowDSO, 1).value = lastRowDSO - 1      ' пїЅ пїЅ/пїЅ
    wsDSO.Cells(lastRowDSO, 2).value = fioVal              ' пїЅпїЅпїЅ
    wsDSO.Cells(lastRowDSO, 3).value = lnVal               ' пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ
    wsDSO.Cells(lastRowDSO, 4).value = ""                  ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ (пїЅпїЅпїЅпїЅпїЅ)

    selectedLichniyNomer = lnVal
    NavigateToDSORow wsDSO, lastRowDSO
    ShowPassportData selectedLichniyNomer
    LoadPeriodsForLichniy selectedLichniyNomer
    lblStatus.Caption = tf("form.search_fio.message.employee_added", _
                           "Employee added to row {row}.", _
                           "{row}", lastRowDSO)
    txtPeriodStart.SetFocus
End Sub

' =====================================================================
' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ ("пїЅпїЅпїЅпїЅпїЅпїЅпїЅ")
' =====================================================================
Private Sub ShowPassportData(lichniyNomer As String)
    mdlHelper.EnsureStaffColumnsInitialized
    
    Dim wsStaff As Worksheet, lastRow As Long, i As Long
    Set wsStaff = mdlHelper.GetStaffWorksheet()
    If wsStaff Is Nothing Then Exit Sub
    
    lastRow = wsStaff.Cells(wsStaff.Rows.count, mdlHelper.colLichniyNomer_Global).End(xlUp).Row
    
    For i = 2 To lastRow
        If Trim(wsStaff.Cells(i, mdlHelper.colLichniyNomer_Global).value) = Trim(lichniyNomer) Then
            lblFIO.Caption = wsStaff.Cells(i, mdlHelper.colFIO_Global).value
            lblZvanie.Caption = t("form.search_fio.label.rank_prefix", "Rank: ") & wsStaff.Cells(i, mdlHelper.colZvanie_Global).value
            lblDolzhnost.Caption = t("form.search_fio.label.position_prefix", "Position: ") & wsStaff.Cells(i, mdlHelper.colDolzhnost_Global).value
            lblChast.Caption = t("form.search_fio.label.unit_prefix", "Unit: ") & Trim(wsStaff.Cells(i, mdlHelper.colVoinskayaChast_Global).value)
            Exit Sub
        End If
    Next i
    
    ' пїЅпїЅпїЅпїЅ пїЅпїЅ пїЅпїЅпїЅпїЅпїЅ (пїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ, пїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅ пїЅпїЅпїЅпїЅпїЅ, пїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅ пїЅпїЅпїЅ)
    lblFIO.Caption = t("form.search_fio.label.fio_missing", "FIO: not found on Staff sheet")
    lblZvanie.Caption = t("form.search_fio.label.rank_prefix", "Rank: ") & "-"
    lblDolzhnost.Caption = t("form.search_fio.label.position_prefix", "Position: ") & "-"
    lblChast.Caption = t("form.search_fio.label.unit_prefix", "Unit: ") & "-"
End Sub

' =====================================================================
' пїЅпїЅпїЅпїЅпїЅпїЅ пїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ
' =====================================================================
Private Sub LoadPeriodsForLichniy(lichniyNomer As String)
    Dim wsDSO As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long, periodCounter As Integer
    Dim baseReasonRaw As String, baseReasonArr() As String
    Dim rowIdx As Long
  
    Set wsDSO = mdlHelper.GetDsoWorksheet()
    If wsDSO Is Nothing Then Exit Sub
    lastRow = wsDSO.Cells(wsDSO.Rows.count, DSO_COL_LN).End(xlUp).Row
    lstPeriods.Clear
    lstPeriods.ColumnCount = 4
    lstPeriods.ColumnWidths = "35 pt;75 pt;75 pt;200 pt"
    
    periodCounter = 0

    For i = 2 To lastRow
        If Trim(wsDSO.Cells(i, DSO_COL_LN).value) = Trim(lichniyNomer) Then
            ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ
            baseReasonRaw = wsDSO.Cells(i, 4).value
            baseReasonArr = Split(baseReasonRaw, ",")
            
            lastCol = wsDSO.Cells(i, wsDSO.Columns.count).End(xlToLeft).Column
            j = 5
            Do While j + 1 <= lastCol
                If wsDSO.Cells(i, j).value <> "" And wsDSO.Cells(i, j + 1).value <> "" Then
                    periodCounter = periodCounter + 1
                    rowIdx = lstPeriods.ListCount
                    lstPeriods.AddItem periodCounter
                    lstPeriods.List(rowIdx, 1) = Format(wsDSO.Cells(i, j).value, "dd.mm.yyyy")
                    lstPeriods.List(rowIdx, 2) = Format(wsDSO.Cells(i, j + 1).value, "dd.mm.yyyy")
                    
                    If UBound(baseReasonArr) >= periodCounter - 1 Then
                        lstPeriods.List(rowIdx, 3) = Trim(baseReasonArr(periodCounter - 1))
                    Else
                        lstPeriods.List(rowIdx, 3) = ""
                    End If
                End If
                j = j + 2
            Loop
            Exit For
        End If
    Next i
    
    If lstPeriods.ListCount = 0 Then
        lstPeriods.AddItem t("form.search_fio.status.no_periods", "No saved periods")
    End If
End Sub

' пїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅ пїЅпїЅпїЅпїЅпїЅпїЅ (пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ)
Private Sub lstPeriods_Click()
    If lstPeriods.ListIndex = -1 Then Exit Sub
    If lstPeriods.List(lstPeriods.ListIndex, 0) = t("form.search_fio.status.no_periods", "No saved periods") Then Exit Sub
    
    txtPeriodStart.Text = lstPeriods.List(lstPeriods.ListIndex, 1) ' пїЅпїЅпїЅпїЅпїЅпїЅ
    txtPeriodEnd.Text = lstPeriods.List(lstPeriods.ListIndex, 2)   ' пїЅпїЅпїЅпїЅпїЅ
    cmbReason.Text = lstPeriods.List(lstPeriods.ListIndex, 3)      ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ
End Sub

' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ ComboBox пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ
Private Sub cmbReason_Enter()
    Dim wsDSO As Worksheet
    Dim lastRow As Long, i As Long
    Dim reasonsDict As Object
    Dim arrBase As Variant, part As Variant
    
    Set reasonsDict = CreateObject("Scripting.Dictionary")
    Set wsDSO = mdlHelper.GetDsoWorksheet()
    If wsDSO Is Nothing Then Exit Sub
    lastRow = wsDSO.Cells(wsDSO.Rows.count, 4).End(xlUp).Row

    cmbReason.Clear

    For i = 2 To lastRow
        If wsDSO.Cells(i, 4).value <> "" Then
            arrBase = Split(wsDSO.Cells(i, 4).value, ",")
            For Each part In arrBase
                part = Trim(part)
                If part <> "" Then
                    If Not reasonsDict.exists(part) Then
                        reasonsDict.Add part, 1
                    End If
                End If
            Next part
        End If
    Next i

    Dim key As Variant
    For Each key In reasonsDict.keys
        cmbReason.AddItem key
    Next key
End Sub

' =====================================================================
' пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ (ADD, EDIT, DELETE)
' =====================================================================

' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ
Private Sub btnAddPeriod_Click()
    If txtPeriodStart.Text = "" Or txtPeriodEnd.Text = "" Or cmbReason.Text = "" Then
        MsgBox t("form.search_fio.message.fill_period", "Fill in dates and basis."), vbExclamation
        Exit Sub
    End If
    If selectedLichniyNomer = "" Then
        MsgBox t("form.search_fio.message.employee_not_selected", "Employee is not selected."), vbExclamation
        Exit Sub
    End If
    If Not ConfirmPeriodEntry() Then Exit Sub

    Dim wsDSO As Worksheet
    Set wsDSO = mdlHelper.GetDsoWorksheet()
    If wsDSO Is Nothing Then Exit Sub
    Dim lastRow As Long, rowNum As Long, found As Boolean, i As Long

    lastRow = wsDSO.Cells(wsDSO.Rows.count, DSO_COL_LN).End(xlUp).Row
    found = False
    
    ' пїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ
    For i = 2 To lastRow
        If Trim(wsDSO.Cells(i, DSO_COL_LN).value) = Trim(selectedLichniyNomer) Then
            found = True
            rowNum = i
            Exit For
        End If
    Next i

    If Not found Then
        MsgBox t("form.search_fio.message.employee_not_found_in_dso", _
                 "Employee was not found in DSO. Add him first."), vbCritical
        Exit Sub
    End If

    ' пїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ
    Dim lastCol As Long
    lastCol = wsDSO.Cells(rowNum, wsDSO.Columns.count).End(xlToLeft).Column
    If lastCol < 4 Then lastCol = 4

    ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅ
    wsDSO.Cells(rowNum, lastCol + 1).value = txtPeriodStart.Text
    wsDSO.Cells(rowNum, lastCol + 2).value = txtPeriodEnd.Text

    ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ
    Dim oldReasonRaw As String, newReasonRaw As String
    oldReasonRaw = CStr(wsDSO.Cells(rowNum, 4).value)
    
    If Trim(oldReasonRaw) = "" Then
        newReasonRaw = cmbReason.Text
    Else
        If Right(Trim(oldReasonRaw), 1) <> "," Then oldReasonRaw = oldReasonRaw & ","
        newReasonRaw = oldReasonRaw & " " & cmbReason.Text
    End If
    wsDSO.Cells(rowNum, 4).value = newReasonRaw

    Call LoadPeriodsForLichniy(selectedLichniyNomer)
    NavigateToDSORow wsDSO, rowNum
    txtPeriodStart.Text = ""
    txtPeriodEnd.Text = ""
    lblStatus.Caption = t("form.search_fio.status.period_added", "Period added.")
    txtPeriodStart.SetFocus
End Sub

' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ
Private Sub btnEditPeriod_Click()
    If selectedLichniyNomer = "" Or lstPeriods.ListIndex = -1 Then Exit Sub
    If Not ConfirmPeriodEntry() Then Exit Sub
    
    Dim wsDSO As Worksheet, lastRow As Long, i As Long
    Dim periodNum As Integer, reasonsArr() As String
    
    Set wsDSO = mdlHelper.GetDsoWorksheet()
    If wsDSO Is Nothing Then Exit Sub
    lastRow = wsDSO.Cells(wsDSO.Rows.count, DSO_COL_LN).End(xlUp).Row
    
    ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ ListBox
    If Not IsNumeric(lstPeriods.List(lstPeriods.ListIndex, 0)) Then Exit Sub
    periodNum = CInt(lstPeriods.List(lstPeriods.ListIndex, 0))

    For i = 2 To lastRow
        If Trim(wsDSO.Cells(i, DSO_COL_LN).value) = Trim(selectedLichniyNomer) Then
            ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅ
            Dim colIndex As Long
            colIndex = 5 + (periodNum - 1) * 2
            wsDSO.Cells(i, colIndex).value = txtPeriodStart.Text
            wsDSO.Cells(i, colIndex + 1).value = txtPeriodEnd.Text

            ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ
            reasonsArr = Split(wsDSO.Cells(i, 4).value, ",")
            
            ' пїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ (пїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ)
            If UBound(reasonsArr) < periodNum - 1 Then
                ReDim Preserve reasonsArr(periodNum - 1)
            End If
            
            reasonsArr(periodNum - 1) = cmbReason.Text
            wsDSO.Cells(i, 4).value = Join(reasonsArr, ",")

            Call LoadPeriodsForLichniy(selectedLichniyNomer)
            NavigateToDSORow wsDSO, i
            lblStatus.Caption = t("form.search_fio.status.period_updated", "Period updated.")
            txtPeriodStart.SetFocus
            Exit Sub
        End If
    Next i
End Sub

' пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ
Private Sub btnDeletePeriod_Click()
    If selectedLichniyNomer = "" Or lstPeriods.ListIndex = -1 Then Exit Sub

    Dim wsDSO As Worksheet
    Dim lastRow As Long, lastCol As Long, i As Long
    Dim periodNum As Integer, reasonsArr() As String, maxPairs As Integer, k As Integer
    
    Set wsDSO = mdlHelper.GetDsoWorksheet()
    If wsDSO Is Nothing Then Exit Sub
    lastRow = wsDSO.Cells(wsDSO.Rows.count, DSO_COL_LN).End(xlUp).Row
    
    If Not IsNumeric(lstPeriods.List(lstPeriods.ListIndex, 0)) Then Exit Sub
    periodNum = CInt(lstPeriods.List(lstPeriods.ListIndex, 0))

    For i = 2 To lastRow
        If Trim(wsDSO.Cells(i, DSO_COL_LN).value) = Trim(selectedLichniyNomer) Then
            lastCol = wsDSO.Cells(i, wsDSO.Columns.count).End(xlToLeft).Column
            maxPairs = (lastCol - 5 + 1) \ 2
            
            ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ
            For k = periodNum To maxPairs - 1
                wsDSO.Cells(i, 5 + (k - 1) * 2).value = wsDSO.Cells(i, 5 + k * 2).value
                wsDSO.Cells(i, 5 + (k - 1) * 2 + 1).value = wsDSO.Cells(i, 5 + k * 2 + 1).value
            Next k
            
            ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ
            wsDSO.Cells(i, 5 + (maxPairs - 1) * 2).value = vbNullString
            wsDSO.Cells(i, 5 + (maxPairs - 1) * 2 + 1).value = vbNullString

            ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ
            reasonsArr = Split(wsDSO.Cells(i, 4).value, ",")
            If UBound(reasonsArr) >= periodNum - 1 Then
                For k = periodNum - 1 To UBound(reasonsArr) - 1
                    reasonsArr(k) = reasonsArr(k + 1)
                Next k
                
                If UBound(reasonsArr) > 0 Then
                    ReDim Preserve reasonsArr(UBound(reasonsArr) - 1)
                    wsDSO.Cells(i, 4).value = Join(reasonsArr, ",")
                Else
                    wsDSO.Cells(i, 4).value = "" ' пїЅпїЅпїЅпїЅ пїЅпїЅпїЅ пїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ
                End If
            End If

            Call CleanPeriodsForLichniyRow(wsDSO, i)
            Call LoadPeriodsForLichniy(selectedLichniyNomer)
            NavigateToDSORow wsDSO, i
            
            txtPeriodStart.Text = ""
            txtPeriodEnd.Text = ""
            lblStatus.Caption = t("form.search_fio.status.period_deleted", "Period deleted.")
            txtPeriodStart.SetFocus
            Exit Sub
        End If
    Next i
End Sub

' пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ (пїЅпїЅпїЅпїЅпїЅпїЅ) пїЅ пїЅпїЅпїЅпїЅпїЅпїЅ
Private Sub CleanPeriodsForLichniyRow(wsDSO As Worksheet, rowNum As Long)
    Dim lastCol As Long, lastPeriodCol As Long, j As Long
    lastCol = wsDSO.Cells(rowNum, wsDSO.Columns.count).End(xlToLeft).Column
    lastPeriodCol = 4

    For j = 5 To lastCol Step 2
        If wsDSO.Cells(rowNum, j).value <> "" And wsDSO.Cells(rowNum, j + 1).value <> "" Then
            lastPeriodCol = j + 1
        End If
    Next j

    For j = lastPeriodCol + 1 To lastCol
        wsDSO.Cells(rowNum, j).value = vbNullString
    Next j
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub NavigateToDSORow(ByVal wsDSO As Worksheet, ByVal rowNum As Long)
    On Error Resume Next
    Application.Goto wsDSO.Cells(rowNum, DSO_COL_FIO), True
    If Not ActiveWindow Is Nothing Then
        ActiveWindow.ScrollColumn = 1
        If rowNum > 3 Then ActiveWindow.ScrollRow = rowNum - 2
    End If
    On Error GoTo 0
End Sub

Private Function ConfirmPeriodEntry() As Boolean
    Dim dStart As Date
    Dim dEnd As Date
    Dim dayCount As Long

    dStart = mdlHelper.ParseDateSafe(txtPeriodStart.Text)
    dEnd = mdlHelper.ParseDateSafe(txtPeriodEnd.Text)

    If dStart = 0 Or dEnd = 0 Then
        MsgBox LocalizeDirect("form.search_fio.message.invalid_dates", "Check the entered dates."), vbExclamation, LocalizeDirect("common.attention", "Attention")
        Exit Function
    End If

    If dEnd < dStart Then
        MsgBox LocalizeDirect("form.search_fio.message.end_before_start", "End date cannot be earlier than start date."), vbExclamation, LocalizeDirect("common.attention", "Attention")
        Exit Function
    End If

    dayCount = DateDiff("d", dStart, dEnd) + 1
    If dayCount > LONG_PERIOD_WARNING_DAYS Then
        If MsgBox(Replace$(LocalizeDirect("form.search_fio.message.long_period_confirm", _
                     "The continuous period is {days} days. Check month and dates. Continue saving?"), _
                     "{days}", CStr(dayCount)), _
                  vbYesNo + vbExclamation, LocalizeDirect("common.attention", "Attention")) <> vbYes Then
            Exit Function
        End If
    End If

    ConfirmPeriodEntry = True
End Function

Private Function LocalizeDirect(ByVal localizationKey As String, ByVal fallback As String) As String
    LocalizeDirect = ModuleLocalization.t(localizationKey, fallback)
End Function

Private Function IsLikelyBrokenDirectText(ByVal value As String) As Boolean
    Dim sample As String

    sample = Trim$(value)
    If Len(sample) = 0 Then Exit Function

    IsLikelyBrokenDirectText = _
        InStr(1, sample, "РїС—", vbTextCompare) > 0 Or _
        InStr(1, sample, "Гђ", vbTextCompare) > 0 Or _
        InStr(1, sample, "Г‘", vbTextCompare) > 0 Or _
        InStr(1, sample, "пїЅпїЅпїЅпїЅ", vbTextCompare) > 0
End Function
