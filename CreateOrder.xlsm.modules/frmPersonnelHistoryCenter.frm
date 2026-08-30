VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPersonnelHistoryCenter 
   Caption         =   "Personnel history and documents"
   ClientHeight    =   9420.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13365
   OleObjectBlob   =   "frmPersonnelHistoryCenter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPersonnelHistoryCenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    lblSummary.Caption = tf("history.center.action_failed", "History search failed: {error}", "{error}", Err.description)
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
    lblSummary.Caption = tf("history.center.action_failed", "History action failed: {error}", "{error}", Err.description)
End Function

Private Sub btnOpenDocument_Click()
    On Error GoTo Failed
    If Not EnsureEmployeeLoaded Then Exit Sub
    mdlPersonnelHistoryCenter.OpenPersonnelHistoryDocument mEmployeeID, Trim$(txtEventID.Text), Trim$(txtDocumentID.Text)
    lblSummary.Caption = t("history.center.action_done", "Document opened.")
    Exit Sub
Failed:
    lblSummary.Caption = tf("history.center.action_failed", "History action failed: {error}", "{error}", Err.description)
End Sub

Private Sub btnRepeatExport_Click()
    On Error GoTo Failed
    If Not EnsureEmployeeLoaded Then Exit Sub
    lblSummary.Caption = mdlPersonnelHistoryCenter.RepeatPersonnelHistoryExport(mEmployeeID, Trim$(txtEventID.Text))
    Exit Sub
Failed:
    lblSummary.Caption = tf("history.center.action_failed", "History action failed: {error}", "{error}", Err.description)
End Sub

Private Sub btnPrepareCorrection_Click()
    On Error GoTo Failed
    If Not EnsureEmployeeLoaded Then Exit Sub
    mdlPersonnelHistoryCenter.PreparePersonnelHistoryCorrectionFromCenter mEmployeeID, Trim$(txtEventID.Text)
    lblSummary.Caption = t("history.center.action_done", "Correction is prepared for review.")
    Exit Sub
Failed:
    lblSummary.Caption = tf("history.center.action_failed", "History action failed: {error}", "{error}", Err.description)
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
