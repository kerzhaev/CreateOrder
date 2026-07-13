Attribute VB_Name = "mdlEnrollmentEventLink"
Option Explicit

' Explicit bridge from the existing enrollment Word exporter to the
' personnel-event ledger. No matching by names, dates, or order numbers.

Private Const EVENTS_SHEET As String = "PersonnelEvents"
Private Const DOCUMENTS_SHEET As String = "DocumentRegistry"
Private Const EVENT_TYPE_ENROLLMENT As String = "ENROLLMENT"
Private Const DOCUMENT_TYPE_ENROLLMENT_ORDER As String = "ENROLLMENT_ORDER"

Public Function RegisterEnrollmentOrderForEvent(ByVal eventID As String, ByVal filePath As String, Optional ByVal documentNumber As String = "", Optional ByVal templateName As String = "EnrollmentOrderTemplate") As String
    Dim eventRow As Long
    Dim documentRow As Long
    Dim events As Worksheet
    Dim documents As Worksheet

    mdlPersonnelEvents.EnsurePersonnelEventInfrastructure
    eventID = LinkText(eventID)
    filePath = LinkText(filePath)
    If eventID = "" Then Err.Raise vbObjectError + 790, "mdlEnrollmentEventLink", "EventID is required."
    If filePath = "" Or Dir$(filePath) = "" Then Err.Raise vbObjectError + 791, "mdlEnrollmentEventLink", "Enrollment Word file was not found."

    Set events = ThisWorkbook.Worksheets(EVENTS_SHEET)
    eventRow = FindEventRow(events, eventID)
    If eventRow = 0 Then Err.Raise vbObjectError + 792, "mdlEnrollmentEventLink", "Personnel event was not found: " & eventID
    If UCase$(LinkText(events.Cells(eventRow, 3).Value)) <> EVENT_TYPE_ENROLLMENT Then Err.Raise vbObjectError + 793, "mdlEnrollmentEventLink", "Only ENROLLMENT events can be linked to an enrollment order."
    If documentNumber = "" Then documentNumber = LinkText(events.Cells(eventRow, 9).Value)

    Set documents = ThisWorkbook.Worksheets(DOCUMENTS_SHEET)
    documentRow = FindDocumentRow(documents, eventID, filePath)
    If documentRow = 0 Then
        documentRow = documents.Cells(documents.Rows.Count, 1).End(xlUp).Row + 1
        If documentRow < 2 Then documentRow = 2
        documents.Cells(documentRow, 1).Value = BuildDocumentID()
    End If
    documents.Cells(documentRow, 2).Value = eventID
    documents.Cells(documentRow, 3).Value = DOCUMENT_TYPE_ENROLLMENT_ORDER
    documents.Cells(documentRow, 4).Value = documentNumber
    documents.Cells(documentRow, 5).Value = Now
    documents.Cells(documentRow, 6).Value = filePath
    documents.Cells(documentRow, 7).Value = templateName
    documents.Cells(documentRow, 8).Value = "ExistingEnrollmentExport"
    documents.Cells(documentRow, 10).Value = "EXPORTED"
    documents.Cells(documentRow, 11).ClearContents
    documents.Cells(documentRow, 12).Value = Now
    mdlPersonnelEvents.SetPersonnelEventStatus eventID, mdlPersonnelEvents.EVENT_STATUS_EXPORTED

    RegisterEnrollmentOrderForEvent = LinkText(documents.Cells(documentRow, 1).Value)
End Function

Public Sub RegisterEnrollmentOrderForEventPrompt()
    Dim eventID As String
    Dim filePath As String
    Dim documentID As String

    On Error GoTo ErrorHandler
    eventID = Trim$(InputBox("Enter saved ENROLLMENT EventID:", "Link enrollment Word order"))
    If eventID = "" Then Exit Sub
    filePath = Trim$(InputBox("Enter full path to the already exported Word order:", "Link enrollment Word order"))
    If filePath = "" Then Exit Sub
    documentID = RegisterEnrollmentOrderForEvent(eventID, filePath)
    MsgBox "Enrollment order was linked: " & documentID, vbInformation
    Exit Sub
ErrorHandler:
    MsgBox "Enrollment order link failed:" & vbCrLf & Err.Description, vbCritical
End Sub

Private Function FindEventRow(ByVal ws As Worksheet, ByVal eventID As String) As Long
    Dim rowNum As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowNum = 2 To lastRow
        If StrComp(LinkText(ws.Cells(rowNum, 1).Value), eventID, vbTextCompare) = 0 Then
            FindEventRow = rowNum
            Exit Function
        End If
    Next rowNum
End Function

Private Function FindDocumentRow(ByVal ws As Worksheet, ByVal eventID As String, ByVal filePath As String) As Long
    Dim rowNum As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowNum = 2 To lastRow
        If StrComp(LinkText(ws.Cells(rowNum, 2).Value), eventID, vbTextCompare) = 0 And StrComp(LinkText(ws.Cells(rowNum, 6).Value), filePath, vbTextCompare) = 0 Then
            FindDocumentRow = rowNum
            Exit Function
        End If
    Next rowNum
End Function

Private Function BuildDocumentID() As String
    BuildDocumentID = "DOC-" & Format$(Now, "yyyymmdd-hhnnss") & "-" & CStr(Int((Timer * 100) Mod 100))
End Function

Private Function LinkText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    LinkText = Trim$(CStr(value))
End Function
