VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDataIntegrityCenter 
   Caption         =   "Data integrity center"
   ClientHeight    =   8010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12765
   OleObjectBlob   =   "frmDataIntegrityCenter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDataIntegrityCenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    severityFilter = UCase$(CStr(cboSeverity.value))
    categoryFilter = UCase$(CStr(cboCategory.value))
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
