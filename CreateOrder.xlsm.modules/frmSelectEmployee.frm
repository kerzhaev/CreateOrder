VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectEmployee 
   Caption         =   "UserForm1"
   ClientHeight    =   7245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16335
   OleObjectBlob   =   "frmSelectEmployee.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'version 5#
'Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectEmployee
'   Caption = "пїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ"
'   ClientHeight = 6000
'   ClientLeft = 120
'   ClientTop = 465
'   ClientWidth = 10000
'   OleObjectBlob   =   "frmSelectEmployee.frx":0000
'   StartUpPosition = 2    'CenterScreen
'End
'Attribute VB_Name = "frmSelectEmployee"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = True
'Attribute VB_Exposed = False

' =====================================================================
' пїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ "пїЅпїЅпїЅпїЅпїЅпїЅпїЅ_пїЅпїЅпїЅ_пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ"
' пїЅпїЅпїЅпїЅпїЅ: пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ, пїЅпїЅпїЅ "95 пїЅпїЅпїЅ" пїЅпїЅ пїЅпїЅ
' =====================================================================

Option Explicit

Private Const RESULT_COL_PERSONAL_NUMBER As Long = 0
Private Const RESULT_COL_FIO As Long = 1
Private Const RESULT_COL_RANK As Long = 2
Private Const RESULT_COL_POSITION As Long = 3
Private Const RESULT_COL_SECTION As Long = 4

Public selectedLichniyNomer As String
Public selectedFIO As String
Public isCancelled As Boolean

Private lblSectionDynamic As MSForms.Label

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    mdlHelper.EnsureStaffColumnsInitialized

    isCancelled = True
    selectedLichniyNomer = ""
    selectedFIO = ""

    ConfigureLayout

    With lstResults
        .ColumnCount = 5
        .ColumnHeads = False
        .BoundColumn = 1
        .ColumnWidths = "75 pt;180 pt;95 pt;180 pt;150 pt"
        .ListStyle = fmListStylePlain
        .MultiSelect = fmMultiSelectSingle
        .IntegralHeight = False
        .Clear
    End With

    Me.Caption = t("form.select_employee.title", "Р’С‹Р±РѕСЂ СЃРѕС‚СЂСѓРґРЅРёРєР°")
    btnSelect.Caption = t("form.select_employee.button.select", "Р’С‹Р±СЂР°С‚СЊ")
    btnCancel.Caption = t("form.select_employee.button.cancel", "РћС‚РјРµРЅР°")
    txtSearch.ControlTipText = t("form.select_employee.search_hint", "Enter FIO, personal number, or table number")
    lblStatus.Caption = t("form.select_employee.status.start", "Р’РІРµРґРёС‚Рµ РЅРµ РјРµРЅРµРµ 2 СЃРёРјРІРѕР»РѕРІ РґР»СЏ РїРѕРёСЃРєР°.")
    ClearSelectedDetails
    Exit Sub

ErrorHandler:
    MsgBox tf("form.select_employee.message.init_error", _
              "РћС€РёР±РєР° РёРЅРёС†РёР°Р»РёР·Р°С†РёРё С„РѕСЂРјС‹: {error}", _
              "{error}", Err.Description), vbCritical, t("common.error", "РћС€РёР±РєР°")
End Sub

Private Sub txtSearch_Change()
    On Error GoTo ErrorHandler

    Dim wsStaff As Worksheet
    Dim lastRow As Long
    Dim rowNum As Long
    Dim foundCount As Long
    Dim query As String
    Dim colTableNumber As Long
    Dim tableValue As String
    Dim fioValue As String
    Dim lnValue As String

    mdlHelper.EnsureStaffColumnsInitialized
    Set wsStaff = mdlHelper.GetStaffWorksheet()
    If wsStaff Is Nothing Then Exit Sub

    query = LCase$(Trim$(txtSearch.Text))
    lstResults.Clear
    ClearSelectedDetails

    If Len(query) < 2 Then
        lblStatus.Caption = t("common.status_enter_min_chars", "Р’РІРµРґРёС‚Рµ РЅРµ РјРµРЅРµРµ 2 СЃРёРјРІРѕР»РѕРІ.")
        Exit Sub
    End If

    If mdlHelper.colFIO_Global <= 0 Or mdlHelper.colLichniyNomer_Global <= 0 Then
        MsgBox t("form.select_employee.message.staff_columns_error", _
                 "РќРµ СѓРґР°Р»РѕСЃСЊ РѕРїСЂРµРґРµР»РёС‚СЊ РѕР±СЏР·Р°С‚РµР»СЊРЅС‹Рµ СЃС‚РѕР»Р±С†С‹ Р»РёСЃС‚Р° 'РЁС‚Р°С‚'."), vbCritical, t("common.error", "РћС€РёР±РєР°")
        Exit Sub
    End If

    lastRow = wsStaff.Cells(wsStaff.Rows.Count, mdlHelper.colFIO_Global).End(xlUp).Row
    colTableNumber = mdlHelper.FindTableNumberColumn(wsStaff)

    For rowNum = 2 To lastRow
        fioValue = LCase$(Trim$(CStr(wsStaff.Cells(rowNum, mdlHelper.colFIO_Global).Value)))
        lnValue = LCase$(Trim$(CStr(wsStaff.Cells(rowNum, mdlHelper.colLichniyNomer_Global).Value)))
        tableValue = ""
        If colTableNumber > 0 Then tableValue = LCase$(Trim$(CStr(wsStaff.Cells(rowNum, colTableNumber).Value)))

        If InStr(1, fioValue, query, vbTextCompare) > 0 _
            Or InStr(1, lnValue, query, vbTextCompare) > 0 _
            Or (tableValue <> "" And InStr(1, tableValue, query, vbTextCompare) > 0) Then
            AddEmployeeRow wsStaff, rowNum, foundCount
            foundCount = foundCount + 1
        End If
    Next rowNum

    If foundCount = 0 Then
        lblStatus.Caption = t("common.status_none", "РЎРѕРІРїР°РґРµРЅРёР№ РЅРµ РЅР°Р№РґРµРЅРѕ.")
        Exit Sub
    End If

    lblStatus.Caption = tf("common.status_found", "РќР°Р№РґРµРЅРѕ: {count}", "{count}", foundCount)
    If foundCount = 1 Then
        lstResults.ListIndex = 0
        lstResults_Click
    End If
    Exit Sub

ErrorHandler:
    lstResults.Clear
    ClearSelectedDetails
    lblStatus.Caption = tf("form.select_employee.message.search_error", _
                           "РћС€РёР±РєР° РїРѕРёСЃРєР°: {error}", _
                           "{error}", Err.Description)
End Sub

Private Sub lstResults_Click()
    On Error GoTo ErrorHandler

    If lstResults.ListCount = 0 Or lstResults.ListIndex < 0 Then Exit Sub

    lblFIO.Caption = t("form.select_employee.label.fio_prefix", "FIO:") & " " & CStr(lstResults.List(lstResults.ListIndex, RESULT_COL_FIO))
    lblZvanie.Caption = t("form.select_employee.label.rank_prefix", "Р—РІР°РЅРёРµ:") & " " & CStr(lstResults.List(lstResults.ListIndex, RESULT_COL_RANK))
    lblDolzhnost.Caption = t("form.select_employee.label.position_prefix", "Р”РѕР»Р¶РЅРѕСЃС‚СЊ:") & " " & CStr(lstResults.List(lstResults.ListIndex, RESULT_COL_POSITION))
    lblSectionDynamic.Caption = t("form.select_employee.label.section_prefix", "Р Р°Р·РґРµР» РїРµСЂСЃРѕРЅР°Р»Р°:") & " " & CStr(lstResults.List(lstResults.ListIndex, RESULT_COL_SECTION))
    lblStatus.Caption = tf("form.select_employee.status.selected", _
                           "Р’С‹Р±СЂР°РЅ: {fio}", _
                           "{fio}", CStr(lstResults.List(lstResults.ListIndex, RESULT_COL_FIO)))
    Exit Sub

ErrorHandler:
    ClearSelectedDetails
End Sub

Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstResults.ListCount > 0 Then
            lstResults.SetFocus
            If lstResults.ListIndex < 0 Then lstResults.ListIndex = 0
        End If
        KeyCode = 0
    ElseIf KeyCode = vbKeyReturn Then
        If lstResults.ListCount = 1 Then
            lstResults.ListIndex = 0
            btnSelect_Click
        ElseIf lstResults.ListCount > 1 Then
            lstResults.SetFocus
            If lstResults.ListIndex < 0 Then lstResults.ListIndex = 0
        End If
        KeyCode = 0
    End If
End Sub

Private Sub lstResults_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btnSelect_Click
        KeyCode = 0
    End If
End Sub

Private Sub lstResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btnSelect_Click
End Sub

Private Sub btnSelect_Click()
    On Error GoTo ErrorHandler

    If lstResults.ListCount = 0 Or lstResults.ListIndex < 0 Then
        MsgBox t("form.select_employee.message.choose_from_list", "Р’С‹Р±РµСЂРёС‚Рµ СЃРѕС‚СЂСѓРґРЅРёРєР° РёР· СЃРїРёСЃРєР°."), _
               vbExclamation, t("common.attention", "Р’РЅРёРјР°РЅРёРµ")
        Exit Sub
    End If

    selectedLichniyNomer = CStr(lstResults.List(lstResults.ListIndex, RESULT_COL_PERSONAL_NUMBER))
    selectedFIO = CStr(lstResults.List(lstResults.ListIndex, RESULT_COL_FIO))
    isCancelled = False
    Me.Hide
    Exit Sub

ErrorHandler:
    MsgBox tf("form.select_employee.message.select_error", _
              "РћС€РёР±РєР° РІС‹Р±РѕСЂР° СЃРѕС‚СЂСѓРґРЅРёРєР°: {error}", _
              "{error}", Err.Description), vbCritical, t("common.error", "РћС€РёР±РєР°")
End Sub

Private Sub btnCancel_Click()
    isCancelled = True
    selectedLichniyNomer = ""
    selectedFIO = ""
    Me.Hide
End Sub

Private Sub ConfigureLayout()
    Dim detailsLeft As Single
    Dim detailsWidth As Single

    Me.Width = 820
    Me.Height = 380

    txtSearch.Left = 18
    txtSearch.Top = 24
    txtSearch.Width = 310

    lstResults.Left = 18
    lstResults.Top = txtSearch.Top + txtSearch.Height + 8
    lstResults.Width = 320
    lstResults.Height = 170

    detailsLeft = lstResults.Left + lstResults.Width + 18
    detailsWidth = Me.InsideWidth - detailsLeft - 18

    lblStatus.Left = detailsLeft
    lblStatus.Top = txtSearch.Top
    lblStatus.Width = detailsWidth
    lblStatus.Height = 18

    lblFIO.Left = detailsLeft
    lblFIO.Top = lstResults.Top + 4
    lblFIO.Width = detailsWidth
    lblFIO.Height = 36
    lblFIO.WordWrap = True

    lblZvanie.Left = detailsLeft
    lblZvanie.Top = lblFIO.Top + lblFIO.Height + 10
    lblZvanie.Width = detailsWidth
    lblZvanie.Height = 30
    lblZvanie.WordWrap = True

    lblDolzhnost.Left = detailsLeft
    lblDolzhnost.Top = lblZvanie.Top + lblZvanie.Height + 10
    lblDolzhnost.Width = detailsWidth
    lblDolzhnost.Height = 54
    lblDolzhnost.WordWrap = True

    Set lblSectionDynamic = EnsureRuntimeLabel("lblSectionDynamic")
    lblSectionDynamic.Left = detailsLeft
    lblSectionDynamic.Top = lblDolzhnost.Top + lblDolzhnost.Height + 10
    lblSectionDynamic.Width = detailsWidth
    lblSectionDynamic.Height = 50
    lblSectionDynamic.WordWrap = True

    btnSelect.Left = 30
    btnSelect.Top = lstResults.Top + lstResults.Height + 28
    btnSelect.Width = 92

    btnCancel.Left = btnSelect.Left + btnSelect.Width + 26
    btnCancel.Top = btnSelect.Top
    btnCancel.Width = 92
End Sub

Private Function EnsureRuntimeLabel(ByVal controlName As String) As MSForms.Label
    On Error Resume Next
    Set EnsureRuntimeLabel = Me.Controls(controlName)
    On Error GoTo 0

    If EnsureRuntimeLabel Is Nothing Then
        Set EnsureRuntimeLabel = Me.Controls.Add("Forms.Label.1", controlName, True)
        EnsureRuntimeLabel.BackStyle = fmBackStyleTransparent
        EnsureRuntimeLabel.Font.Name = "Times New Roman"
        EnsureRuntimeLabel.Font.Size = 10
    End If
End Function

Private Sub AddEmployeeRow(ByVal wsStaff As Worksheet, ByVal rowNum As Long, ByVal listRow As Long)
    lstResults.AddItem Trim$(CStr(wsStaff.Cells(rowNum, mdlHelper.colLichniyNomer_Global).Value))
    lstResults.List(listRow, RESULT_COL_FIO) = Trim$(CStr(wsStaff.Cells(rowNum, mdlHelper.colFIO_Global).Value))
    lstResults.List(listRow, RESULT_COL_RANK) = Trim$(CStr(wsStaff.Cells(rowNum, mdlHelper.colZvanie_Global).Value))
    lstResults.List(listRow, RESULT_COL_POSITION) = Trim$(CStr(wsStaff.Cells(rowNum, mdlHelper.colDolzhnost_Global).Value))
    lstResults.List(listRow, RESULT_COL_SECTION) = Trim$(CStr(wsStaff.Cells(rowNum, mdlHelper.colVoinskayaChast_Global).Value))
End Sub

Private Sub ClearSelectedDetails()
    lblFIO.Caption = t("form.select_employee.label.fio_prefix", "FIO:")
    lblZvanie.Caption = t("form.select_employee.label.rank_prefix", "Р—РІР°РЅРёРµ:")
    lblDolzhnost.Caption = t("form.select_employee.label.position_prefix", "Р”РѕР»Р¶РЅРѕСЃС‚СЊ:")
    If Not lblSectionDynamic Is Nothing Then
        lblSectionDynamic.Caption = t("form.select_employee.label.section_prefix", "Р Р°Р·РґРµР» РїРµСЂСЃРѕРЅР°Р»Р°:")
    End If
End Sub


